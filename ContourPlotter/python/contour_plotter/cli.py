# SPDX-License-Identifier: GPL-2.0-or-later
"""Independent command line engine and file-based ORD bridge."""
import argparse
from datetime import datetime, timezone
import json
import math
from pathlib import Path
import sys
import time
import traceback
import uuid
from . import __version__
from .tin_reader import inspect, read_tin
from .contour_engine import extract, level_values
from .label_engine import place_labels
from .output_serializer import atomic_json, write_geometry

SETTING_KEYS = {"minor", "major", "base", "label_mode", "label_spacing", "paper_height",
                "scale", "clearance_paper", "break_lines", "max_labels", "dimension", "text_widths", "text_box_height", "paper_height_major", "paper_height_minor", "text_box_height_minor"}
SOURCE_KEYS = {"surface_index", "surface_name", "axis_order", "horizontal_unit",
               "vertical_unit", "model_unit", "crs", "vertical_datum"}


def safe_tsv(text):
    return str(text or "").replace("\t", " ").replace("\r", " ").replace("\n", " ")


def run(request):
    if request.get("version") != 1:
        raise ValueError("Unsupported request version")
    action = request.get("action", "generate")
    source = Path(request["input"])
    directory = Path(request["output_dir"])
    if source.absolute().parent == directory.absolute() or source.resolve().parent == directory.resolve():
        raise ValueError("Keep the LandXML source outside the output directory")
    directory.mkdir(parents=True, exist_ok=True)
    run_id = str(uuid.UUID(request.get("run_id", str(uuid.uuid4()))))
    cancel_file = directory / "cancel"
    def cancelled():
        return cancel_file.exists()
    def status(stage):
        atomic_json(directory / "progress.json", {"run_id": run_id, "stage": stage})
    started = time.perf_counter()
    status("Reading LandXML")
    if action == "inspect":
        report = inspect(source)
        atomic_json(directory / "inspection.json", report)
        with (directory / "surfaces.tsv").open("w", encoding="utf-16") as stream:
            stream.write("INSPECT1\t" + safe_tsv(report["horizontal_unit"]) + "\t" +
                         safe_tsv(report["vertical_unit"]) + "\t" + safe_tsv(report["declared_crs"]) + "\n")
            for i, surface in enumerate(report["surfaces"]):
                stream.write("\t".join(map(safe_tsv, [i, surface["name"] or f"Surface {i+1}",
                                                     surface["points"], surface["triangles"]])) + "\n")
        atomic_json(directory / "result.json", {"ok": True, "run_id": run_id, "inspection": report})
        return report
    if action not in {"analyze", "generate"}:
        raise ValueError("Unsupported action")
    raw_settings = request.get("settings", {})
    if set(raw_settings) - SETTING_KEYS:
        raise ValueError("Unknown plotting settings")
    spatial = request.get("spatial", {})
    if set(spatial) - SOURCE_KEYS:
        raise ValueError("Unknown spatial settings")
    points, faces, metadata = read_tin(source, **spatial)
    if request.get("expected_source_sha256") and request["expected_source_sha256"] != metadata["source_sha256"]:
        raise ValueError("LandXML changed after analysis; analyze again")
    parsed = time.perf_counter()
    settings = dict(minor=1, major=5, base=0, label_mode="major", label_spacing=250,
                    paper_height=0.10, scale=50, clearance_paper=0.02, break_lines=True,
                    max_labels=100, dimension=2, text_widths={})
    settings.update(raw_settings)
    for key in ("minor", "major", "label_spacing", "paper_height", "scale", "clearance_paper"):
        value = settings[key]
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
            raise ValueError(f"{key} must be finite and positive")
    if settings["dimension"] not in (2, 3) or not isinstance(settings["break_lines"], bool):
        raise ValueError("Invalid output dimension or gap setting")
    for key in ("paper_height_major", "paper_height_minor"):
        settings.setdefault(key, settings["paper_height"])
        value = settings[key]
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
            raise ValueError(f"{key} must be finite and positive")
    height = settings["paper_height_major"] * settings["scale"]
    minor_height = settings["paper_height_minor"] * settings["scale"]
    levels = level_values(metadata["zmin"], metadata["zmax"], settings["minor"],
                          settings["major"], settings["base"])
    # ORD measures these exact strings using its chosen DGN text style.
    with (directory / "levels.tsv").open("w", encoding="ascii") as stream:
        for _, _major, text in levels:
            stream.write(text + "\t" + ("M" if _major else "m") + "\n")
    if action == "analyze":
        report = {"ok": True, "run_id": run_id, "source": metadata, "height": height,
                  "minor_height": minor_height, "elevations": len(levels), "parse_seconds": parsed - started}
        with (directory / "analysis.tsv").open("w", encoding="utf-16") as stream:
            stream.write("\t".join(map(safe_tsv, [metadata["surface_name"], metadata["point_count"],
                metadata["face_count"], metadata["zmin"], metadata["zmax"], metadata["declared_crs"],
                *metadata["bounds"], height, minor_height, metadata["vertical_unit"], metadata["source_sha256"]])) + "\n")
        atomic_json(directory / "result.json", report)
        return report
    status("Extracting contours on source triangles")
    contours = extract(points, faces, settings["minor"], settings["major"], settings["base"], cancelled)
    extracted = time.perf_counter()
    status("Placing labels and physical gaps")
    contours = place_labels(contours, mode=settings["label_mode"], spacing=settings["label_spacing"],
                            height=height, clearance=settings["clearance_paper"] * settings["scale"],
                            break_lines=settings["break_lines"], widths=settings["text_widths"],
                            max_labels=settings["max_labels"], box_height=settings.get("text_box_height", height),
                            minor_height=minor_height, minor_box_height=settings.get("text_box_height_minor"), cancel=cancelled)
    labelled = time.perf_counter()
    status("Validating and writing native CAD geometry")
    stats = write_geometry(directory / "geometry.cp2", run_id, contours, metadata, height,
                           settings["dimension"], cancelled, minor_height=minor_height)
    completed = time.perf_counter()
    warnings = list(metadata["warnings"])
    warnings += ["No smoothing or simplification applied.",
                 "CRS, vertical datum and axis mapping are user-confirmed; no reprojection performed."]
    if not settings["text_widths"] and stats["text_elements"]:
        warnings.append("Text widths estimated; production ORD bridge supplies measured font widths.")
    report = {"ok": True, "version": 1, "engine_version": __version__, "run_id": run_id,
              "timestamp_utc": datetime.now(timezone.utc).isoformat(), "source": metadata,
              "settings": settings, "statistics": {**stats, "contours": len(contours),
                  "short_fragments": sum(bool(c["short"]) for c in contours),
                  "parse_seconds": parsed - started, "contour_seconds": extracted - parsed,
                  "label_seconds": labelled - extracted, "serialize_seconds": completed - labelled,
                  "total_seconds": completed - started}, "warnings": warnings}
    atomic_json(directory / "result.json", report)
    status("Complete")
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description="Direct LandXML TIN contours for native ORD VBA insertion")
    parser.add_argument("--request", type=Path, help="Version 1 JSON request from ORD")
    parser.add_argument("--input", type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--inspect", action="store_true")
    parser.add_argument("--analyze", action="store_true")
    parser.add_argument("--surface-index", type=int)
    parser.add_argument("--surface-name")
    parser.add_argument("--axis-order", choices=["xy", "ne"])
    parser.add_argument("--horizontal-unit", choices=["m", "ft", "us-ft"])
    parser.add_argument("--vertical-unit", choices=["m", "ft", "us-ft"])
    parser.add_argument("--model-unit", choices=["m", "ft", "us-ft"])
    parser.add_argument("--crs")
    parser.add_argument("--vertical-datum")
    parser.add_argument("--minor", type=float, default=1)
    parser.add_argument("--major", type=float, default=5)
    parser.add_argument("--base", type=float, default=0)
    parser.add_argument("--label-mode", choices=["major", "both", "none"], default="major")
    parser.add_argument("--label-spacing", type=float, default=250)
    parser.add_argument("--paper-height", type=float, default=.1)
    parser.add_argument("--major-paper-height", dest="paper_height_major", type=float)
    parser.add_argument("--minor-paper-height", dest="paper_height_minor", type=float)
    parser.add_argument("--scale", type=float, default=50, help="Model master units per paper inch")
    parser.add_argument("--dimension", type=int, choices=[2, 3], default=2)
    parser.add_argument("--no-gaps", action="store_true")
    parser.add_argument("--version", action="version", version=__version__)
    args = parser.parse_args(argv)
    request = None
    try:
        if args.request:
            with args.request.open(encoding="utf-8-sig") as stream:
                request = json.load(stream)
        else:
            if not args.input or not args.output_dir:
                parser.error("--input and --output-dir are required without --request")
            request = {"version": 1, "input": str(args.input), "output_dir": str(args.output_dir),
                       "action": "inspect" if args.inspect else "analyze" if args.analyze else "generate",
                       "spatial": {k: getattr(args, k) for k in SOURCE_KEYS if getattr(args, k, None) is not None},
                       "settings": {k: getattr(args, k) for k in ("minor", "major", "base", "label_mode",
                           "label_spacing", "paper_height", "scale", "dimension")}}
            request["settings"]["break_lines"] = not args.no_gaps
            for key in ("paper_height_major", "paper_height_minor"):
                if getattr(args, key) is not None:
                    request["settings"][key] = getattr(args, key)
        report = run(request)
        print(json.dumps({"ok": True, "statistics": report.get("statistics", {}), "version": __version__}))
        return 0
    except Exception as exc:
        error = {"ok": False, "error": str(exc), "traceback": traceback.format_exc()}
        if isinstance(request, dict) and request.get("output_dir"):
            directory = Path(request["output_dir"])
            directory.mkdir(parents=True, exist_ok=True)
            error_path = directory / "error.json"
            source_path = Path(request.get("input", ""))
            if source_path.absolute() != error_path.absolute() and source_path.resolve() != error_path.resolve():
                atomic_json(error_path, error)
            message_path = directory / "error.txt"
            if source_path.absolute() != message_path.absolute() and source_path.resolve() != message_path.resolve():
                message_path.write_text(str(exc), encoding="utf-16")
        print(str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
