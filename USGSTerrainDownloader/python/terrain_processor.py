"""Auditable DEM processing. Horizontal PROJ pipeline never transforms Z."""
import json
from contextlib import suppress
import math
import subprocess
from pathlib import Path

from pyproj import CRS
from pyproj.exceptions import ProjError
from pyproj.crs import CompoundCRS
from crs_utils import UNIT_METERS, horizontal_crs, linear_factor, require, transformer


def gis():
    try:
        from osgeo import gdal, ogr, osr
        import numpy as np
    except ImportError as exc:
        raise RuntimeError("GDAL Python bindings and numpy are required. Use the configured QGIS/OSGeo4W Python environment.") from exc
    require(int(gdal.VersionInfo()) >= 3080000, "GDAL >= 3.8 is required.")
    gdal.UseExceptions()
    ogr.UseExceptions()
    osr.UseExceptions()
    return gdal, ogr, osr, np


def unit_name(value):
    value = value.strip().lower()
    if value in ("m", "meter", "metre", "meters", "metres"):
        return "m"
    if value in ("us-ft", "us survey foot", "us survey feet", "foot_us", "ftus"):
        return "us-ft"
    if value in ("ft", "international foot", "international feet"):
        return "ft"
    require(not value, f"Ambiguous/unsupported embedded elevation unit: {value!r}. Verify outside this tool.")
    return None


def vertical_contract(confirmation):
    require(isinstance(confirmation, dict), "Each source needs a vertical confirmation.")
    require(confirmation.get("unit") in UNIT_METERS, "Confirm source elevation unit: m, ft, or us-ft.")
    datum = CRS.from_user_input(confirmation.get("crs", ""))
    require(datum.is_vertical, "Supply an explicit vertical CRS such as EPSG:5703 (only when verified).")
    require(len(confirmation.get("evidence", "").strip()) >= 12,
            "Document source-specific vertical datum/unit evidence (metadata URL, report, or embedded tags).")
    return datum


def output_vertical(datum, unit):
    if abs(datum.axis_info[0].unit_conversion_factor - UNIT_METERS[unit]) < 1e-14:
        return datum
    obj = datum.to_json_dict()
    obj["datum"] = datum.datum.to_json_dict()
    obj["datum"].pop("$schema", None)
    obj.pop("id", None)
    obj["name"] = f"{datum.datum.name} height ({unit})"
    obj["coordinate_system"]["axis"][0]["unit"] = {
        "type": "LinearUnit", "name": {"m": "metre", "ft": "foot", "us-ft": "US survey foot"}[unit],
        "conversion_factor": UNIT_METERS[unit],
        "id": {"authority": "EPSG", "code": {"m": 9001, "ft": 9002, "us-ft": 9003}[unit]}}
    return CRS.from_json_dict(obj)


def inspect_raster(path, confirmation):
    gdal, _, _, _ = gis()
    ds = gdal.Open(str(path), gdal.GA_ReadOnly)
    require(ds is not None and ds.RasterCount == 1, "DEM must be a readable, single-band raster.")
    require(ds.GetProjection(), f"Source raster has no CRS: {path}")
    declared = CRS.from_wkt(ds.GetProjection())
    hcrs = declared.to_2d()
    require(hcrs.is_projected or hcrs.is_geographic, "Unsupported raster horizontal CRS.")
    band = ds.GetRasterBand(1)
    scale, offset = band.GetScale(), band.GetOffset()
    require(scale in (None, 1) and offset in (None, 0), "Scaled/offset raster bands require explicit unpacking outside this tool.")
    require(not gdal.DataTypeIsComplex(band.DataType), "Complex-valued DEM is unsupported.")
    datum = vertical_contract(confirmation)
    embedded_unit = unit_name(band.GetUnitType())
    if embedded_unit:
        require(embedded_unit == confirmation["unit"], "Source elevation unit contradicts its embedded band unit.")
    for part in declared.sub_crs_list:
        if part.is_vertical:
            require(part.datum == datum.datum, "Source vertical datum contradicts confirmed vertical datum.")
            factor = part.axis_info[0].unit_conversion_factor
            require(abs(factor - UNIT_METERS[confirmation["unit"]]) < 1e-12,
                    "Source vertical CRS unit contradicts confirmed elevation unit.")
    gt = ds.GetGeoTransform()
    require(gt[1] > 0 and gt[5] < 0 and gt[2] == 0 and gt[4] == 0,
            "Only north-up unrotated source DEMs are supported.")
    info = {"path": str(path), "crs_wkt": declared.to_wkt(), "horizontal_wkt": hcrs.to_wkt(),
            "geotransform": list(gt), "width": ds.RasterXSize, "height": ds.RasterYSize,
            "band_type": gdal.GetDataTypeName(band.DataType),
            "nodata": str(band.GetNoDataValue()), "band_unit": band.GetUnitType(),
            "vertical_confirmation": confirmation, "dataset_metadata": ds.GetMetadata()}
    ds = None
    return info, hcrs, datum


def compatible_tiles(infos, crss, confirmations, overlap_confirmed):
    first = infos[0]
    gt = first["geotransform"]
    for info, crs, confirmation in zip(infos, crss, confirmations):
        require(crs.equals(crss[0], ignore_axis_order=True),
                "Mosaic source horizontal CRSs differ; select one compatible collection.")
        other = info["geotransform"]
        require(abs(other[1]/gt[1]-1) < 1e-9 and abs(other[5]/gt[5]-1) < 1e-9,
                "Conflicting source resolutions; choose compatible products.")
        require(confirmation["unit"] == confirmations[0]["unit"], "Mixed source elevation units are unsupported.")
        for origin, step, coordinate in ((gt[0], gt[1], other[0]), (gt[3], gt[5], other[3])):
            delta = (coordinate-origin)/step
            require(abs(delta-round(delta)) < 1e-6, "Source grids are misaligned; mosaic requires aligned tiles.")
    def extent(info):
        g = info["geotransform"]
        return g[0], g[3]+g[5]*info["height"], g[0]+g[1]*info["width"], g[3]
    overlaps = []
    for i, a in enumerate(infos):
        ea = extent(a)
        for j in range(i):
            eb = extent(infos[j])
            if min(ea[2], eb[2])-max(ea[0], eb[0]) > gt[1]*1e-6 and min(ea[3], eb[3])-max(ea[1], eb[1]) > -gt[5]*1e-6:
                overlaps.append([j, i])
    require(not overlaps or overlap_confirmed,
            "Selected DEMs overlap. Review dates/collections and explicitly confirm list-order priority (last selected wins).")
    return overlaps


def process_dem(paths, products, config, boundary, folder, check=lambda: None, progress=lambda message: None):
    gdal, ogr, osr, np = gis()
    folder = Path(folder)
    require(folder.is_dir(), "Processing folder does not exist.")
    target = horizontal_crs(boundary["crs_wkt"])
    confirmations = [config.get("vertical_confirmations", {}).get(p["id"], config.get("source_vertical")) for p in products]
    infos, crss, datums = [], [], []
    for path, confirmation in zip(paths, confirmations):
        check()
        info, crs, datum = inspect_raster(path, confirmation)
        infos.append(info); crss.append(crs); datums.append(datum)
    require(all(d.datum == datums[0].datum for d in datums),
            "Mixed vertical datums: no validated vertical transformation is implemented.")
    overlaps = compatible_tiles(infos, crss, confirmations, config.get("overlap_order_confirmed") is True)
    output_unit = config.get("elevation_unit", "m")
    factor = linear_factor(confirmations[0]["unit"], output_unit)
    resolution = float(config["resolution"])
    require(math.isfinite(resolution) and resolution > 0, "Output resolution must be positive in target CRS native units.")
    ring = boundary["ring"]
    polygon = ogr.Geometry(ogr.wkbPolygon)
    linear = ogr.Geometry(ogr.wkbLinearRing)
    for x, y in ring:
        linear.AddPoint_2D(x, y)
    polygon.AddGeometry(linear)
    require(polygon.IsValid(), "Boundary is self-intersecting or invalid. Redraw the boundary.")
    xs, ys = zip(*ring)
    bounds = [min(xs), min(ys), max(xs), max(ys)]
    width, height = math.ceil((bounds[2]-bounds[0])/resolution), math.ceil((bounds[3]-bounds[1])/resolution)
    require(width > 0 and height > 0 and width*height <= 100_000_000,
            "Output exceeds 100 million cells or is empty. Increase cell size or reduce boundary.")
    srs = osr.SpatialReference()
    srs.ImportFromWkt(target.to_wkt())
    srs.SetAxisMappingStrategy(osr.OAMS_TRADITIONAL_GIS_ORDER)
    cut = folder / "boundary.geojson"
    cut.write_text(json.dumps({"type": "FeatureCollection", "crs": {"type": "name", "properties": {"name": target.to_wkt()}},
        "features": [{"type": "Feature", "properties": {}, "geometry": json.loads(polygon.ExportToJson())}]}), encoding="utf-8")
    memory_driver = ogr.GetDriverByName("MEM") or ogr.GetDriverByName("Memory")
    layer_ds = memory_driver.CreateDataSource("")
    layer = layer_ds.CreateLayer("boundary", srs, ogr.wkbPolygon)
    feature = ogr.Feature(layer.GetLayerDefn()); feature.SetGeometry(polygon); layer.CreateFeature(feature)
    feature = None
    operation = transformer(crss[0].to_2d(), target)
    # Resolve the operation with an actual point before passing the exact pipeline to GDAL.
    source_gt = infos[0]["geotransform"]
    operation.transform(source_gt[0], source_gt[3], errcheck=True)
    try:
        operation = operation.get_last_used_operation()
    except ProjError:
        require(operation.definition not in ("", "unavailable until proj_trans is called"),
                "PROJ could not resolve the horizontal operation.")
    vrt_path, temp, output = folder / "mosaic.vrt", folder / "terrain.partial.tif", folder / "terrain.tif"
    working = folder / "terrain.working.tif"
    require(not output.exists() and not temp.exists() and not working.exists(), "Output already exists; use a new run folder.")
    callback_error = []
    def callback(fraction, message, data):
        try:
            check()
            progress(f"Processing: horizontal warp {fraction:.0%}")
            return 1
        except Exception as error:
            callback_error.append(error)
            return 0
    nodata, output_nodata = -1e300, -999999.0
    vrt = warped = band = mask = verify = delivery = output_band = None
    try:
        vrt = gdal.BuildVRT(str(vrt_path), [str(p) for p in paths],
                            options=gdal.BuildVRTOptions(resolution="highest", strict=True))
        require(vrt is not None, "GDAL could not construct the mosaic.")
        vrt.FlushCache()
        progress("Processing: mosaic and horizontal reprojection")
        try:
            warped = gdal.Warp(str(working), vrt, options=gdal.WarpOptions(
                format="GTiff", dstSRS=target.to_wkt(), coordinateOperation=operation.definition,
                outputBounds=bounds, xRes=resolution, yRes=resolution,
                outputType=gdal.GDT_Float64, dstNodata=nodata,
                resampleAlg="bilinear", errorThreshold=0.0, multithread=False,
                options=["-novshift"], callback=callback,
                creationOptions=["TILED=YES", "COMPRESS=DEFLATE", "BIGTIFF=IF_SAFER", "GEOTIFF_VERSION=1.1"]))
        except RuntimeError:
            if callback_error:
                raise callback_error[0]
            raise
        require(warped is not None, "GDAL warp produced no output.")
        # Rasterize expected coverage separately: NoData inside the polygon is a coverage gap.
        mask = gdal.GetDriverByName("MEM").Create("", warped.RasterXSize, warped.RasterYSize, 1, gdal.GDT_Byte)
        mask.SetGeoTransform(warped.GetGeoTransform()); mask.SetProjection(target.to_wkt())
        gdal.RasterizeLayer(mask, [1], layer, burn_values=[1])
        band = warped.GetRasterBand(1)
        # ORD accepted Float32 with finite NoData where the prior Float64 delivery failed.
        delivery = gdal.GetDriverByName("GTiff").Create(str(temp), warped.RasterXSize, warped.RasterYSize, 1,
            gdal.GDT_Float32, options=["TILED=YES", "COMPRESS=DEFLATE", "BIGTIFF=IF_SAFER", "GEOTIFF_VERSION=1.1"])
        delivery.SetGeoTransform(warped.GetGeoTransform())
        output_band = delivery.GetRasterBand(1)
        output_band.SetNoDataValue(output_nodata)
        valid_count = expected = gaps = 0
        minimum, maximum = math.inf, -math.inf
        rounding_error = 0.0
        for y in range(0, warped.RasterYSize, 256):
            for x in range(0, warped.RasterXSize, 256):
                check()
                w, h = min(256, warped.RasterXSize-x), min(256, warped.RasterYSize-y)
                a = band.ReadAsArray(x, y, w, h)
                inside = mask.GetRasterBand(1).ReadAsArray(x, y, w, h) == 1
                # Clip after resampling: source-grid cutlines can erase covered target edge cells.
                if config.get("clip", True):
                    a[~inside] = nodata
                valid = a != nodata
                require(np.isfinite(a[valid]).all(), "Nonfinite raster elevation detected.")
                expected += int(inside.sum()); gaps += int((inside & ~valid).sum())
                # Float64 conversion applies only to valid samples, independently of XY warp.
                a[valid] *= factor
                require(np.isfinite(a[valid]).all() and not np.any(a[valid] == nodata), "Elevation conversion overflow/NoData collision.")
                with np.errstate(over="ignore"):
                    rounded = a[valid].astype(np.float32)
                require(np.isfinite(rounded).all() and not np.any(rounded == output_nodata),
                        "Float32 elevation overflow/NoData collision; output not released.")
                if rounded.size:
                    rounding_error = max(rounding_error, float(np.max(np.abs(rounded.astype(np.float64) - a[valid]))))
                a[valid] = rounded
                a[~valid] = output_nodata
                samples = a[valid & inside]
                if samples.size:
                    minimum = min(minimum, float(samples.min())); maximum = max(maximum, float(samples.max()))
                    valid_count += int(samples.size)
                output_band.WriteArray(a, x, y)
            progress(f"Processing: elevation units / coverage {min(y+256, warped.RasterYSize)}/{warped.RasterYSize} rows")
        require(expected > 0 and valid_count > 0, "No valid DEM cell centers inside boundary. Reduce cell size or verify coverage.")
        require(gaps == 0 or config.get("allow_gaps") is True,
                f"Coverage gap: {gaps}/{expected} boundary cells are NoData. Select more coverage; output not released.")
        vertical = output_vertical(datums[0], output_unit)
        compound = CompoundCRS("DGN horizontal + verified source vertical (no datum shift)", [target, vertical])
        output_srs = osr.SpatialReference()
        output_srs.ImportFromWkt(compound.to_wkt())
        # GDAL GeoTIFF needs the vertical UNIT authority to retain non-metre units.
        output_srs.SetAuthority("COMPD_CS|VERT_CS|UNIT", "EPSG",
                                {"m": 9001, "ft": 9002, "us-ft": 9003}[output_unit])
        delivery.SetSpatialRef(output_srs)
        output_band.SetUnitType(output_unit)
        delivery.SetMetadataItem("VERTICAL_DATUM", datums[0].datum.name)
        delivery.SetMetadataItem("ELEVATION_UNIT", output_unit)
        delivery.SetMetadataItem("VERTICAL_OPERATION", "Unit scaling only; no vertical datum transformation")
        delivery.SetMetadataItem("SOURCE_REPORT", "report.json")
        output_band.FlushCache(); delivery.FlushCache(); output_band = None
        delivery.Close(); delivery = None
        band = None
        warped.Close(); warped = None
        mask.Close(); mask = None
        vrt.Close(); vrt = None
        verify = gdal.Open(str(temp))
        require(verify is not None and CRS.from_wkt(verify.GetProjection()).equals(compound), "Output CRS verification failed.")
        require(verify.GetRasterBand(1).GetUnitType() == output_unit, "Output elevation unit tag verification failed.")
        require(verify.GetRasterBand(1).DataType == gdal.GDT_Float32 and
                verify.GetRasterBand(1).GetNoDataValue() == output_nodata, "Output storage verification failed.")
        require(verify.GetRasterBand(1).Checksum() >= 0, "Output read-back checksum failed.")
        geotransform = list(verify.GetGeoTransform())
        verify.Close(); verify = None
        check()
        temp.replace(output)
        return {"output": str(output), "sources": infos, "output_crs_wkt": compound.to_wkt(),
                "horizontal_pipeline": operation.definition, "horizontal_accuracy_m": operation.accuracy,
                "vertical_datum": datums[0].datum.name, "vertical_unit": output_unit,
                "elevation_multiplier": factor, "vertical_transformation": "none; unit conversion only",
                "band_type": "Float32", "nodata": output_nodata,
                "max_storage_rounding_error": rounding_error, "rounding_error_unit": output_unit,
                "resolution": resolution, "resampling": "bilinear", "geotransform": geotransform,
                "clipping": "target-grid polygon cell centers after horizontal resampling" if config.get("clip", True) else "disabled",
                "valid_boundary_cells": valid_count, "expected_boundary_cells": expected,
                "gap_cells": gaps, "overlaps": overlaps, "overlap_policy": "selected list order, last wins",
                "elevation_range": [minimum, maximum], "gdal_version": gdal.VersionInfo("--version"),
                "warning": "Not survey-grade. No vertical datum shift performed. Validate with project control before design use."}
    finally:
        band = output_band = None
        for dataset in (verify, delivery, warped, mask, vrt):
            if dataset is not None:
                with suppress(RuntimeError):
                    dataset.Close()
        verify = delivery = warped = mask = vrt = None
        # A failed/cancelled raster is never named terrain.tif.
        temp.unlink(missing_ok=True)
        working.unlink(missing_ok=True)


def inspect_lidar(path, executable, check=lambda: None):
    """Optional PDAL metadata/schema only. No ground classification or rasterization implied."""
    require(Path(executable).is_file(), "Configure a full PDAL executable path for point-cloud inspection.")
    check()
    result = subprocess.run([str(executable), "info", "--metadata", "--schema", str(path)],
                            capture_output=True, text=True, timeout=180, check=True)
    check()
    return json.loads(result.stdout)
