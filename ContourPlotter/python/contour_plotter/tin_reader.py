# SPDX-License-Identifier: GPL-2.0-or-later
"""Standalone adapter around the attributed QGIS TIN reader. No GDAL/QGIS."""
import hashlib
from pathlib import Path
import numpy as np
from .vendor.tin import read_tin as _read_tin
from .vendor.landxml.parser import load_document

UNIT_METRES = {"m": 1.0, "ft": 0.3048, "us-ft": 1200 / 3937}
UNIT_NAMES = {"meter": "m", "metre": "m", "foot": "ft", "internationalFoot": "ft",
              "USSurveyFoot": "us-ft", **{key: key for key in UNIT_METRES}}


def resolve_unit(declared, supplied, label):
    actual = UNIT_NAMES.get(declared)
    if declared and actual is None:
        raise ValueError(f"Unsupported declared {label} unit: {declared!r}")
    if supplied and supplied not in UNIT_METRES:
        raise ValueError(f"Unsupported {label} unit: {supplied!r}")
    if actual and supplied and actual != supplied:
        raise ValueError(f"{label} unit conflicts with LandXML declaration")
    if not (actual or supplied):
        raise ValueError(f"Undetermined {label} unit; explicitly supply m, ft or us-ft")
    return actual or supplied


def inspect(path):
    doc = load_document(path)
    report = doc.inspect()
    report["source_path"] = str(Path(path).resolve())
    return report


def read_tin(path, *, surface_index=None, surface_name=None, axis_order=None,
             horizontal_unit=None, vertical_unit=None, model_unit=None,
             crs=None, vertical_datum=None):
    if axis_order not in {"xy", "ne"}:
        raise ValueError("Explicit axis_order required: xy (E N Z) or ne (N E Z)")
    if model_unit not in UNIT_METRES:
        raise ValueError("Explicit model_unit required: m, ft or us-ft")
    if not isinstance(crs, str) or not crs.strip():
        raise ValueError("Explicit verified CRS/local grid description required")
    if not isinstance(vertical_datum, str) or not vertical_datum.strip():
        raise ValueError("Explicit verified vertical datum description required")
    with Path(path).open("rb") as source_stream:
        before_digest = hashlib.file_digest(source_stream, "sha256").hexdigest()
    points, faces, metadata = _read_tin(path, surface_name=surface_name,
                                      surface_index=surface_index)
    if surface_name is not None and metadata["surface_name"] != surface_name:
        raise ValueError("Surface name/index no longer match; inspect the source again")
    if any("conflicting unit declarations" in warning for warning in metadata["warnings"]):
        raise ValueError("LandXML has conflicting unit declarations")
    hu = resolve_unit(metadata["horizontal_unit"], horizontal_unit, "horizontal")
    vu = resolve_unit(metadata["vertical_unit"], vertical_unit, "vertical")
    # Reject malformed geometry before passing it into the native contour library.
    used = np.unique(faces)
    xy = points[used, :2]
    if len(np.unique(xy, axis=0)) != len(xy):
        raise ValueError("TIN contains duplicate used XY vertices; repair the source TIN")
    if len(np.unique(np.sort(faces, axis=1), axis=0)) != len(faces):
        raise ValueError("TIN contains duplicate triangles")
    a, b, c = (points[faces[:, i], :2] for i in range(3))
    ab, ac = b - a, c - a
    cross = ab[:, 0] * ac[:, 1] - ab[:, 1] * ac[:, 0]
    if np.any(cross == 0):
        raise ValueError("TIN contains a degenerate triangle in XY")
    edges = np.sort(np.concatenate((faces[:, [0, 1]], faces[:, [1, 2]],
                                   faces[:, [2, 0]])), axis=1)
    if np.any(np.unique(edges, axis=0, return_counts=True)[1] > 2):
        raise ValueError("TIN contains nonmanifold edges")
    # Unused points must not enlarge the contour elevation range.
    remap = np.full(len(points), -1, dtype=np.int32)
    remap[used] = np.arange(len(used))
    points, faces = points[used].copy(), remap[faces]
    if axis_order == "ne":
        points[:, [0, 1]] = points[:, [1, 0]]
    factor = UNIT_METRES[hu] / UNIT_METRES[model_unit]
    points[:, :2] *= factor
    with Path(path).open("rb") as source_stream:
        digest = hashlib.file_digest(source_stream, "sha256").hexdigest()
    if digest != before_digest:
        raise ValueError("LandXML changed while reading; use a stable export")
    metadata.update(horizontal_unit=hu, vertical_unit=vu, model_unit=model_unit,
                    axis_order=axis_order, confirmed_crs=crs.strip(),
                    vertical_datum=vertical_datum.strip(),
                    coordinate_conversion=factor,
                    source_path=str(Path(path).resolve()),
                    source_sha256=digest)
    metadata.update(bounds=[*points[:, :2].min(axis=0), *points[:, :2].max(axis=0)],
                    zmin=float(points[:, 2].min()), zmax=float(points[:, 2].max()),
                    point_count=len(points), face_count=len(faces))
    return points, faces, metadata
