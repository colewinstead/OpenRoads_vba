"""Clipped LandXML TIN from unmasked DEM samples; no retriangulation across gaps."""
import math
import xml.etree.ElementTree as ET
from contextlib import suppress
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import quoteattr

from crs_utils import UNIT_METERS, horizontal_crs, linear_factor, require, ring_valid
from terrain_processor import gis

NS = "http://www.landxml.org/schema/LandXML-1.2"
MAX_GRID_CELLS = 1_000_000
MAX_FACES = 2_000_000


def cross(a, b, c):
    return (b[0]-a[0])*(c[1]-a[1]) - (b[1]-a[1])*(c[0]-a[0])


def triangulate_ring(points):
    """Ear clipping retains concave edges; inputs are simple rings without holes."""
    points = list(points)
    area = math.fsum(cross(points[0], a, b) for a, b in zip(points[1:], points[2:]))
    if area < 0:
        points.reverse()
    while len(points) > 3:
        for i, b in enumerate(points):
            a, c = points[i-1], points[(i+1) % len(points)]
            if cross(a, b, c) <= 0:
                continue
            if any(cross(a, b, p) >= 0 and cross(b, c, p) >= 0 and cross(c, a, p) >= 0
                   for j, p in enumerate(points) if j not in ((i-1) % len(points), i, (i+1) % len(points))):
                continue
            yield a, b, c
            del points[i]
            break
        else:
            raise ValueError("Cannot triangulate clipped boundary; no LandXML released.")
    if len(points) == 3 and cross(*points) != 0:
        yield tuple(points)


def polygon_parts(geometry, ogr):
    if geometry.IsEmpty() or geometry.GetDimension() < 2:
        return
    if ogr.GT_Flatten(geometry.GetGeometryType()) == ogr.wkbPolygon:
        require(geometry.GetGeometryCount() == 1, "Unexpected hole in clipped triangle; no LandXML released.")
        yield geometry
    else:
        for child in geometry:
            yield from polygon_parts(child, ogr)


def verify_landxml(path, point_count, face_count, check):
    points = faces = 0
    for event, element in ET.iterparse(path, events=("end",)):
        if element.tag == f"{{{NS}}}P":
            points += 1
            values = [float(v) for v in element.text.split()]
            require(int(element.attrib['id']) == points and len(values) == 3 and
                    all(math.isfinite(v) for v in values), "LandXML point read-back failed.")
        elif element.tag == f"{{{NS}}}F":
            faces += 1
            ids = [int(v) for v in element.text.split()]
            require(len(ids) == 3 and len(set(ids)) == 3 and all(1 <= v <= point_count for v in ids),
                    "LandXML face read-back failed.")
        if (points + faces) % 10000 == 0:
            check()
        element.clear()
    require(points == point_count and faces == face_count, "LandXML count read-back failed.")


def export_landxml(paths, config, boundary, processing, folder, check=lambda: None, progress=lambda message: None):
    gdal, ogr, _, np = gis()
    folder = Path(folder)
    output, partial = folder/'terrain.xml', folder/'terrain.partial.xml'
    samples_path = folder/'terrain.mesh.working.tif'
    points_path, faces_path = folder/'terrain.points.tmp', folder/'terrain.faces.tmp'
    require(not any(p.exists() for p in (output, partial, samples_path, points_path, faces_path)),
            "LandXML output already exists; use a new run folder.")
    target = horizontal_crs(boundary['crs_wkt'])
    native_factor = target.axis_info[0].unit_conversion_factor
    unit = next((u for u, f in UNIT_METERS.items() if abs(f-native_factor) < 1e-14), None)
    require(unit is not None, "LandXML requires a metre, international-foot or US-survey-foot horizontal CRS.")
    factor = linear_factor(processing['sources'][0]['vertical_confirmation']['unit'], unit)
    resolution = processing['resolution']
    source = gdal.Open(processing['output'])
    gt = source.GetGeoTransform()
    width, height = source.RasterXSize, source.RasterYSize
    source.Close()
    require((width+1)*(height+1) <= MAX_GRID_CELLS,
            "LandXML exceeds the one-million grid-cell limit. Increase cell size or reduce the boundary; verified GeoTIFF retained.")
    if config.get('clip', True):
        ring = [[x*boundary['model_to_native_scale'], y*boundary['model_to_native_scale']]
                for x, y in ring_valid(config['boundary'])]
    else:
        ring = [[gt[0], gt[3]], [gt[0]+width*resolution, gt[3]],
                [gt[0]+width*resolution, gt[3]-height*resolution], [gt[0], gt[3]-height*resolution]]
        ring.append(ring[0])
    bounds = [gt[0]-resolution, gt[3]-(height+1)*resolution,
              gt[0]+(width+1)*resolution, gt[3]+resolution]
    origin_x, origin_y = bounds[0], bounds[3]
    linear = ogr.Geometry(ogr.wkbLinearRing)
    for x, y in ring:
        linear.AddPoint_2D((x-origin_x)/resolution, (origin_y-y)/resolution)
    polygon = ogr.Geometry(ogr.wkbPolygon); polygon.AddGeometry(linear)
    require(polygon.IsValid(), "Invalid LandXML clip polygon.")
    expected_area = polygon.GetArea()*resolution**2
    # Joining at 1e-10 of a grid cell is below the precision of typical projected coordinates.
    xy_tolerance = max(resolution*1e-10, 8*math.ulp(origin_x), 8*math.ulp(origin_y))
    area_tolerance = polygon.Boundary().Length()*resolution*xy_tolerance + 128*math.ulp(expected_area)
    nodata = -1e300
    vrt = samples = None
    callback_errors = []
    def callback(fraction, message, data):
        try:
            check()
            progress(f"LandXML: sampling boundary elevations {fraction:.0%}")
            return 1
        except Exception as error:
            callback_errors.append(error)
            return 0
    try:
        check()
        vrt = gdal.BuildVRT('', [str(p) for p in paths], options=gdal.BuildVRTOptions(resolution='highest', strict=True))
        require(vrt is not None, "Cannot construct LandXML source mosaic.")
        try:
            samples = gdal.Warp(str(samples_path), vrt, options=gdal.WarpOptions(
                format='GTiff', dstSRS=target.to_wkt(), coordinateOperation=processing['horizontal_pipeline'],
                outputBounds=bounds, xRes=resolution, yRes=resolution, outputType=gdal.GDT_Float64,
                dstNodata=nodata, resampleAlg='bilinear', errorThreshold=0.0, options=['-novshift'],
                callback=callback, creationOptions=['TILED=YES', 'COMPRESS=DEFLATE', 'BIGTIFF=IF_SAFER']))
        except RuntimeError:
            if callback_errors:
                raise callback_errors[0]
            raise
        require(samples is not None, "LandXML elevation sampling failed.")
        point_ids = {}
        faces = 0
        area_rows = []
        minimum, maximum = math.inf, -math.inf
        skipped_nodata_triangles = 0
        with points_path.open('w', encoding='utf-8') as point_stream, faces_path.open('w', encoding='utf-8') as face_stream:
            for y in range(samples.RasterYSize-1):
                check()
                row_area = []
                values = samples.GetRasterBand(1).ReadAsArray(0, y, samples.RasterXSize, 2)
                require(np.isfinite(values[values != nodata]).all(), "Nonfinite LandXML source elevation; no XML released.")
                for x in range(samples.RasterXSize-1):
                    if x % 256 == 0:
                        check()
                    quad = [(x+0.5,y+0.5), (x+1.5,y+0.5), (x+1.5,y+1.5), (x+0.5,y+1.5)]
                    elevations = [values[0,x], values[0,x+1], values[1,x+1], values[1,x]]
                    for indexes in ((0,1,2), (0,2,3)):
                        z = [elevations[i] for i in indexes]
                        if any(v == nodata or not math.isfinite(v) for v in z):
                            skipped_nodata_triangles += 1
                            continue
                        abc = [quad[i] for i in indexes]
                        edge = ogr.Geometry(ogr.wkbLinearRing)
                        for px, py in (*abc, abc[0]): edge.AddPoint_2D(px, py)
                        triangle = ogr.Geometry(ogr.wkbPolygon); triangle.AddGeometry(edge)
                        if polygon.Contains(triangle):
                            clipped = triangle
                        elif polygon.Intersects(triangle):
                            clipped = polygon.Intersection(triangle)
                        else:
                            continue
                        for part in polygon_parts(clipped, ogr):
                            vertices = [p[:2] for p in part.GetGeometryRef(0).GetPoints()[:-1]]
                            for face in triangulate_ring(vertices):
                                face = [(round(px,10), round(py,10)) for px, py in face]
                                signed_area = cross(*face)/2
                                if signed_area == 0:
                                    continue
                                ids = []
                                for px, py in face:
                                    a, b, c = abc
                                    u = cross(a, (px,py), c)/cross(a,b,c)
                                    v = cross(a, b, (px,py))/cross(a,b,c)
                                    elevation = float((z[0]+u*(z[1]-z[0])+v*(z[2]-z[0]))*factor)
                                    require(math.isfinite(elevation), "Nonfinite LandXML elevation.")
                                    key = (px, py)
                                    if key in point_ids:
                                        ident, previous_z = point_ids[key]
                                        require(abs(previous_z-elevation) <= max(1e-8, 32*math.ulp(elevation)),
                                                "Inconsistent elevations at a shared LandXML vertex.")
                                    else:
                                        ident = len(point_ids)+1
                                        point_ids[key] = (ident, elevation)
                                        point_stream.write(f'<P id="{ident}">{origin_y-py*resolution:.12f} {origin_x+px*resolution:.12f} {elevation:.12f}</P>\n')
                                        minimum, maximum = min(minimum,elevation), max(maximum,elevation)
                                    ids.append(ident)
                                # Local raster Y runs south; reverse to counterclockwise world XY.
                                if signed_area > 0: ids[1], ids[2] = ids[2], ids[1]
                                face_stream.write(f'<F>{ids[0]} {ids[1]} {ids[2]}</F>\n')
                                row_area.append(abs(signed_area)*resolution**2)
                                faces += 1
                                require(faces <= MAX_FACES, "LandXML exceeds two million faces; increase cell size.")
                area_rows.append(math.fsum(row_area))
                progress(f"LandXML: clipping triangles {y+1}/{samples.RasterYSize-1} rows")
        count = len(point_ids)
        point_ids.clear()
        covered_area = math.fsum(area_rows)
        gap_area = max(0.0, expected_area-covered_area)
        require(count >= 3 and faces > 0, "No valid LandXML triangles; verify coverage or reduce cell size.")
        require(covered_area <= expected_area+area_tolerance, "LandXML triangle area exceeds its boundary.")
        require(gap_area <= area_tolerance or config.get('allow_gaps') is True,
                f"LandXML coverage gap: {gap_area:.9g} square {unit}. Boundary interpolation needs surrounding source samples. Select more coverage or explicitly allow/report gaps; verified GeoTIFF retained.")
        now = datetime.now(timezone.utc)
        attrs = lambda values: ' '.join(f'{k}={quoteattr(str(v))}' for k,v in values.items())
        unit_attrs = {'linearUnit': {'m':'meter','ft':'foot','us-ft':'USSurveyFoot'}[unit],
                      'areaUnit': 'squareMeter' if unit == 'm' else 'squareFoot',
                      'volumeUnit': 'cubicMeter' if unit == 'm' else 'cubicFeet',
                      'temperatureUnit': 'celsius' if unit == 'm' else 'fahrenheit',
                      'pressureUnit': 'milliBars' if unit == 'm' else 'inHG',
                      'angularUnit': 'decimal degrees', 'directionUnit': 'decimal degrees'}
        cs_attrs = {'name': target.name, 'horizontalDatum': target.datum.name,
                    'verticalDatum': processing['vertical_datum']}
        if target.to_epsg(): cs_attrs['epsgCode'] = target.to_epsg()
        with partial.open('w', encoding='utf-8') as stream:
            stream.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            stream.write(f'<LandXML xmlns="{NS}" version="1.2" date="{now.date()}" time="{now:%H:%M:%S}" readOnly="true" language="English">\n')
            stream.write(f'<Units><{"Metric" if unit == "m" else "Imperial"} {attrs(unit_attrs)}/></Units>\n')
            stream.write(f'<CoordinateSystem {attrs(cs_attrs)}/>\n')
            stream.write('<Surfaces><Surface name="USGS Terrain" desc="Clipped DEM triangles; import Definition without retriangulation">\n')
            stream.write(f'<Definition surfType="TIN" elevMin="{minimum:.12f}" elevMax="{maximum:.12f}"><Pnts>\n')
            for path, closing in ((points_path, '</Pnts><Faces>\n'), (faces_path, '</Faces></Definition></Surface></Surfaces></LandXML>\n')):
                with path.open(encoding='utf-8') as source_stream:
                    while chunk := source_stream.read(1024*1024):
                        check(); stream.write(chunk)
                stream.write(closing)
        verify_landxml(partial, count, faces, check)
        check()
        partial.replace(output)
        return {'output': str(output), 'version': '1.2', 'points': count, 'faces': faces,
                'coordinate_order': 'northing easting elevation', 'horizontal_unit': unit, 'vertical_unit': unit,
                'horizontal_crs_wkt': target.to_wkt(), 'vertical_datum': processing['vertical_datum'],
                'elevation_multiplier_from_source': factor, 'elevation_range': [minimum,maximum],
                'clipped': config.get('clip', True), 'expected_area': expected_area, 'triangle_area': covered_area,
                'gap_area': gap_area, 'area_unit': f'square {unit}', 'area_tolerance': area_tolerance,
                'xy_join_tolerance': xy_tolerance, 'skipped_nodata_triangles': skipped_nodata_triangles,
                'interpolation': 'Float64 cell-center lattice; northwest-southeast diagonal; planar edge interpolation; one-cell sampling halo',
                'ord_build_terrain_from': 'Definition', 'native_ord_import_validated': False}
    finally:
        for dataset in (samples, vrt):
            if dataset is not None:
                with suppress(RuntimeError): dataset.Close()
        samples = vrt = None
        for path in (partial, samples_path, points_path, faces_path): path.unlink(missing_ok=True)