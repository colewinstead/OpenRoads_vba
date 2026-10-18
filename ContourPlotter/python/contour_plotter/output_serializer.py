# SPDX-License-Identifier: GPL-2.0-or-later
"""CP2: bounded ASCII records for streaming in VBA. JSON stores rich provenance."""
import hashlib
import json
import math
import os
from pathlib import Path
import uuid
import numpy as np
from .tin_reader import UNIT_METRES


def atomic_json(path, data):
    path = Path(path)
    temp = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        with temp.open("w", encoding="utf-8") as stream:
            json.dump(data, stream, ensure_ascii=True, allow_nan=False, indent=2)
            stream.write("\n")
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


def chunks(points, size=1024):
    for first in range(0, len(points) - 1, size - 1):
        yield points[first:first + size]


def write_geometry(path, run_id, contours, metadata, height, dimension=2, cancel=lambda: False, *, minor_height=None):
    path = Path(path)
    uuid.UUID(run_id)
    lines = sum(sum(math.ceil((len(p) - 1) / 1023) for p in c["pieces"]) for c in contours)
    labels = sum(len(c["labels"]) for c in contours)
    vertices = sum(sum(sum(len(q) for q in chunks(p)) for p in c["pieces"]) for c in contours)
    zfactor = UNIT_METRES[metadata["vertical_unit"]] / UNIT_METRES[metadata["model_unit"]]
    minor_height = height if minor_height is None else minor_height
    header = ["CP2", run_id, metadata["model_unit"], metadata["vertical_unit"],
              lines, labels, vertices, height, minor_height, *metadata["bounds"], dimension]
    temporary = path.with_name(path.name + ".tmp")
    def emit(stream, fields):
        stream.write("\t".join(format(float(v), ".17g") if isinstance(v, (float, np.floating))
                               else str(v) for v in fields) + "\n")
    try:
        with temporary.open("w", encoding="ascii", newline="\r\n") as stream:
            emit(stream, header)
            for contour in contours:
                if cancel():
                    raise InterruptedError("Serialization cancelled")
                cid, elevation = contour["id"], contour["elevation"]
                kind = "M" if contour["major"] else "m"
                z = elevation * zfactor if dimension == 3 else 0.0
                for piece in contour["pieces"]:
                    for points in chunks(piece):
                        emit(stream, ["C", cid, elevation, kind, len(points)])
                        for x, y in points:
                            emit(stream, ["V", x, y, z])
                for label in contour["labels"]:
                    emit(stream, ["L", cid, elevation, kind, label["x"], label["y"], z,
                                  label["rotation"], label["height"], label["width"], label["text"]])
            emit(stream, ["E", lines, labels, vertices])
        validate_geometry(temporary, run_id)
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)
    with path.open("rb") as stream:
        digest = hashlib.file_digest(stream, "sha256").hexdigest()
    return {"line_elements": lines, "text_elements": labels, "vertices": vertices,
            "geometry_sha256": digest, "geometry_bytes": path.stat().st_size}


def validate_geometry(path, run_id=None):
    """Independent complete pass before publication; VBA repeats this while locked."""
    def numeric(text):
        value = float(text)
        if not math.isfinite(value) or abs(value) > 1e12:
            raise ValueError("Nonfinite or out-of-range CP2 number")
        return value
    with Path(path).open("r", encoding="ascii") as stream:
        def fields():
            line = stream.readline()
            if not line or len(line) > 4096:
                raise ValueError("Truncated or overlong CP2 record")
            return line.rstrip("\n").split("\t")
        h = fields()
        if len(h) != 14 or h[0] != "CP2" or (run_id and h[1] != run_id):
            raise ValueError("Invalid CP2 header or run ID")
        uuid.UUID(h[1])
        if h[2] not in UNIT_METRES or h[3] not in UNIT_METRES or h[13] not in {"2", "3"}:
            raise ValueError("Invalid CP2 units/dimension")
        expected = [int(n) for n in h[4:7]]
        if any(n < 0 or n > 100000000 for n in expected):
            raise ValueError("Invalid CP2 counts")
        height = numeric(h[7])
        minor_height = numeric(h[8])
        xmin, ymin, xmax, ymax = map(numeric, h[9:13])
        if height <= 0 or minor_height <= 0 or xmax < xmin or ymax < ymin:
            raise ValueError("Invalid CP2 bounds/height")
        counts = [0, 0, 0]
        expected_z = 0.0
        def point(v):
            x, y, z = map(numeric, v)
            tol = max(1e-8, max(abs(x), abs(y)) * 1e-14)
            if not xmin - tol <= x <= xmax + tol or not ymin - tol <= y <= ymax + tol:
                raise ValueError("CP2 point outside source bounds")
            if abs(z - expected_z) > max(1e-10, abs(expected_z)*1e-14):
                raise ValueError("CP2 Z differs from contour elevation")
            if h[13] == "2" and z != 0:
                raise ValueError("Nonzero Z in 2D output")
        while True:
            f = fields()
            if f[0] == "E":
                if len(f) != 4 or list(map(int, f[1:])) != counts or counts != expected or stream.read(1):
                    raise ValueError("CP2 footer/count/trailing-data mismatch")
                break
            if f[0] not in {"C", "L"} or len(f) != (5 if f[0] == "C" else 11):
                raise ValueError("Invalid CP2 record")
            if int(f[1]) < 1 or f[3] not in {"M", "m"}:
                raise ValueError("Invalid CP2 contour identity")
            elevation = numeric(f[2])
            expected_z = elevation * UNIT_METRES[h[3]] / UNIT_METRES[h[2]] if h[13] == "3" else 0.0
            if f[0] == "C":
                count = int(f[4])
                if not 2 <= count <= 1024:
                    raise ValueError("Invalid CP2 vertex count")
                previous = None
                for _ in range(count):
                    v = fields()
                    if len(v) != 4 or v[0] != "V":
                        raise ValueError("Invalid CP2 vertex")
                    point(v[1:])
                    if previous == v[1:]:
                        raise ValueError("Duplicate consecutive CP2 vertex")
                    previous = v[1:]
                counts[0] += 1
                counts[2] += count
            else:
                point(f[4:7])
                rotation, text_height, width = map(numeric, f[7:10])
                if abs(rotation) > math.pi / 2 + 1e-12 or text_height != (height if f[3] == "M" else minor_height) or width <= 0:
                    raise ValueError("Invalid CP2 text dimensions/rotation")
                if not f[10] or len(f[10]) > 64 or numeric(f[10]) != numeric(f[2]):
                    raise ValueError("CP2 label does not equal contour elevation")
                counts[1] += 1
    return counts
