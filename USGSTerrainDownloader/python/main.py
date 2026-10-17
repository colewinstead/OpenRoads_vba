"""One request per subprocess. Atomic JSON plus UTF-16 UI files; cancellation is cooperative."""
import hashlib
import json
import os
import shutil
import sys
import time
import traceback
from datetime import datetime, timezone
from pathlib import Path

DEPENDENCY_ERROR = None
try:
    from crs_utils import prepare_boundary, horizontal_crs, require
    from usgs_client import TNMClient, download
except ImportError as error:
    DEPENDENCY_ERROR = str(error)


def atomic_json(path, value):
    temp = Path(str(path) + ".tmp")
    temp.write_text(json.dumps(value, indent=2, ensure_ascii=True, allow_nan=False), encoding="utf-8")
    temp.replace(path)


def clean(value):
    return str(value if value is not None else "unknown").replace("\t", " ").replace("\r", " ").replace("\n", " ")


def ui_file(path, lines):
    temp = Path(str(path) + ".tmp")
    temp.write_text("\r\n".join(lines), encoding="utf-16", newline="")
    # An ORD reader may briefly hold the old file on Windows.
    for attempt in range(20):
        try:
            temp.replace(path)
            return
        except PermissionError:
            if attempt == 19:
                raise
            time.sleep(0.025)


def fingerprint(boundary):
    return hashlib.sha256(json.dumps({k: boundary[k] for k in ("crs_wkt", "ring", "model_to_native_scale")},
                                    sort_keys=True).encode()).hexdigest()


class Job:
    def __init__(self, folder):
        self.folder = Path(folder)
        self.last = 0
        self.state = "starting"

    def check(self):
        if (self.folder / "cancel.flag").exists():
            raise InterruptedError("Cancelled; original completed downloads retained. No partial raster released.")

    def status(self, message, force=False):
        if force or time.monotonic()-self.last > 0.2:
            self.last = time.monotonic()
            atomic_json(self.folder / "status.json", {"state": self.state, "message": message})
            ui_file(self.folder / "status.txt", [self.state, clean(message)])


def run(config, job):
    require(config.get("schema_version") == 1, "Unsupported request schema version.")
    require(config.get("action") in ("search", "process"), "Action must be search or process.")
    job.check()
    boundary = prepare_boundary(config)
    if config["action"] == "search":
        job.state = "searching"
        job.status("Validating boundary / TNM catalog", True)
        client = TNMClient(check=job.check, progress=job.status)
        products = client.search(config["dataset"], boundary["bbox"])
        manifest = {"schema_version": 1, "utc": datetime.now(timezone.utc).isoformat(),
                    "boundary": boundary, "boundary_fingerprint": fingerprint(boundary),
                    "dataset": config["dataset"], "products": products}
        atomic_json(job.folder / "search.json", manifest)
        ui_file(job.folder / "products.tsv", ["\t".join(clean(v) for v in (
            p["id"], p["title"], p["bytes"], p["publication"], p["resolution"], p["bbox"],
            p["vertical"], p["metadata_url"], json.dumps(p["raw"], ensure_ascii=True),
            (p.get("vertical_confirmation") or {}).get("unit", ""),
            (p.get("vertical_confirmation") or {}).get("crs", ""),
            (p.get("vertical_confirmation") or {}).get("evidence", ""))) for p in products])
        ui_file(job.folder / "extent.txt", [f"WGS84 lon/lat bbox: {boundary['bbox']}",
            f"{boundary['crs_name']} | {boundary['horizontal_unit']} | {boundary['area_m2']/1e6:.3f} km2"])
        # The angular 1/3-arc-second collection has no single metre spacing to assume.
        ui_file(job.folder / "cell-size.txt", [format(1 / horizontal_crs(boundary["crs_wkt"]).axis_info[0].unit_conversion_factor, ".12g")]
                if config["dataset"] == "dem1m" else [])
        job.state = "search_complete"
        job.status(f"{len(products)} overlapping candidates. Review metadata and select products; search bbox may overselect irregular boundaries.", True)
        return manifest
    manifest_path = Path(config["search_manifest"])
    require(manifest_path.is_file() and manifest_path.stat().st_size < 32*1024*1024, "Missing/oversized search manifest.")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    require(manifest["boundary_fingerprint"] == fingerprint(boundary), "Boundary/CRS changed since search. Search again.")
    require(manifest["dataset"] == config["dataset"], "Dataset changed since search. Search again.")
    ids = config.get("selected_ids")
    require(isinstance(ids, list) and len(ids) > 0 and len(ids) == len(set(ids)), "Select unique product IDs.")
    selected = [p for p in manifest["products"] if p["id"] in ids]
    require(len(selected) == len(ids), "Selected IDs are absent from the search results.")
    budget = config.get("max_download_bytes", 5*1024**3)
    require(isinstance(budget, int) and 0 < budget <= 100*1024**3, "Download cap must be > 0 and <= 100 GiB.")
    known = sum(p["bytes"] or 0 for p in selected)
    require(known <= budget, "Selected catalog estimate exceeds download cap.")
    require(all(p["bytes"] for p in selected) or config.get("unknown_size_confirmed") is True,
            "A selected product has unknown size. Explicit confirmation is required; hard transfer cap still applies.")
    require(shutil.disk_usage(job.folder).free > max(known, budget if not known else known)*2 + 256*1024**2,
            "Insufficient free space for sources and processing. Choose a different destination or reduce selection/cap.")
    if config["dataset"] != "lidar":
        # Validate vertical declarations before potentially large transfers.
        from terrain_processor import gis, vertical_contract
        gis()
        for p in selected:
            confirmation = config.get("vertical_confirmations", {}).get(p["id"], config.get("source_vertical"))
            datum = vertical_contract(confirmation)
            detected = p.get("vertical_confirmation")
            if detected:
                require(confirmation["unit"] == detected["unit"] and datum.datum == vertical_contract(detected).datum,
                        f"Vertical declaration contradicts selected product metadata: {p['title']}")
    report = {"schema_version": 1, "utc": datetime.now(timezone.utc).isoformat(), "state": "downloading",
              "request": config, "search_manifest": str(manifest_path.resolve()), "boundary": boundary,
              "products": selected, "downloads": [], "sources_preserved": True,
              "limits": "No survey-grade accuracy claimed. Vertical datum transformations are not implemented."}
    atomic_json(job.folder / "report.json", report)
    try:
        job.state = "downloading"
        paths, used = [], 0
        # Sources live in the destination cache, outside individual job folders; receipts enable safe retry.
        sources = job.folder.parent / "originals"
        for p in selected:
            job.check()
            job.status(f"Download: {p['title']}", True)
            path, receipt = download(p, sources, budget-used, job.check, job.status)
            used += receipt["bytes"]; paths.append(path)
            report["downloads"].append({"product_id": p["id"], "path": str(path), **receipt})
            atomic_json(job.folder / "report.json", report)
        if config["dataset"] == "lidar":
            report["lidar_inspection"] = []
            if config.get("pdal_executable"):
                from terrain_processor import inspect_lidar
                job.state = "inspecting_lidar"
                for path in paths:
                    report["lidar_inspection"].append(inspect_lidar(path, config["pdal_executable"], job.check))
                report["state"] = "lidar_downloaded_inspected"
            else:
                report["state"] = "lidar_downloaded_uninspected"
            report["warning"] = "Original LAS/LAZ only; no ground classification, bare-earth raster, or terrain model created. Review point CRS, vertical metadata, and classification schema with PDAL/QGIS."
        else:
            from terrain_processor import process_dem
            job.state = "processing"
            report["processing"] = process_dem(paths, selected, config, boundary, job.folder, job.check, job.status)
            report["state"] = "dem_complete"
            (job.folder / "ORD_IMPORT.txt").write_text(
                "Import terrain.tif manually in a disposable ORD design model first.\n"
                "Review report.json: horizontal CRS/native units, vertical datum, separate elevation unit, gaps and overlaps.\n"
                "GeoTIFF is a terrain raster, not a native ORD Terrain Model or a certified survey.\n"
                "Delivery uses Float32 samples and finite NoData; report.json records the maximum storage rounding in output elevation units.\n"
                "Use ORD's installed Terrain > Create From File workflow if GeoTIFF is supported in your build; otherwise use an approved GIS conversion.\n"
                "Set/verify the import elevation units; compare known control and metadata before project use.\n"
                "Raster attachment/terrain creation are not automated by this VBA tool.\n", encoding="utf-8")
        atomic_json(job.folder / "report.json", report)
        job.state = report["state"]
        job.status("DEM verified: terrain.tif and report.json ready for review." if config["dataset"] != "lidar" else report["warning"], True)
        return report
    except Exception as error:
        report["state"] = "cancelled" if isinstance(error, InterruptedError) else "failed"
        report["error"] = str(error)
        atomic_json(job.folder / "report.json", report)
        raise


def main():
    if len(sys.argv) != 2:
        print("Usage: python main.py request.json", file=sys.stderr)
        return 2
    request = Path(sys.argv[1]).resolve()
    job = Job(request.parent)
    try:
        if DEPENDENCY_ERROR:
            raise RuntimeError("Missing Python dependency. Use QGIS/OSGeo4W Python with pyproj and numpy: " + DEPENDENCY_ERROR)
        require(request.is_file() and request.stat().st_size <= 4*1024*1024, "Missing/oversized request.json")
        config = json.loads(request.read_text(encoding="utf-8-sig"))
        run(config, job)
        return 0
    except Exception as error:
        job.state = "cancelled" if isinstance(error, InterruptedError) else "failed"
        job.status(str(error), True)
        (job.folder / "error.log").write_text(traceback.format_exc(), encoding="utf-8")
        return 2


if __name__ == "__main__":
    sys.exit(main())
