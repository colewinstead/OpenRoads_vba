# SPDX-License-Identifier: GPL-2.0-or-later
# Adapted from LandXML Road & Terrain GIS Tools core.py, copyright (C) 2026 Edmond Akello.
from __future__ import annotations

import math

import numpy as np

from .landxml.common import child, descendants, first_descendant
from .landxml.parser import load_document


def list_surfaces(path):
    surfaces = load_document(path).surfaces()
    return [
        (s.attrib.get("name", f"Surface {i + 1}"), s) for i, s in enumerate(surfaces)
    ]


def declared_epsg(path):
    document = load_document(path)
    root = document.root
    cs = child(root, "CoordinateSystem")
    if cs is None:
        return None
    value = cs.attrib.get("epsgCode")
    if value and str(value).isdigit():
        return int(value)
    return None


def read_tin(path, surface_name=None, progress=None, cancel=None, surface_index=None):
    if progress:
        progress(2, "Parsing LandXML…")
    document = load_document(path)
    root = document.root
    surfaces = list(descendants(root, "Surface"))
    if not surfaces:
        raise ValueError("No <Surface> elements were found in the LandXML file.")
    if surface_index is not None:
        if not 0 <= surface_index < len(surfaces):
            raise ValueError(f"Surface index {surface_index} is out of range.")
        surface = surfaces[surface_index]
        surface_name = surface.attrib.get("name", f"Surface {surface_index + 1}")
    elif surface_name is None:
        if len(surfaces) != 1:
            raise ValueError("Multiple surfaces: choose a name or zero-based surface index explicitly.")
        surface = surfaces[0]
        surface_name = surface.attrib.get("name", "Surface 1")
    else:
        matches = [
            surface
            for index, surface in enumerate(surfaces)
            if surface.attrib.get("name", f"Surface {index + 1}") == surface_name
        ]
        if not matches:
            raise ValueError(f"Surface not found: {surface_name}")
        if len(matches) > 1:
            raise ValueError(
                f"Surface name '{surface_name}' is ambiguous; {len(matches)} surfaces have that name."
            )
        surface = matches[0]

    definitions = list(descendants(surface, "Definition"))
    if len(definitions) != 1 or definitions[0].attrib.get("surfType") != "TIN":
        raise ValueError("Select a surface with one explicitly declared TIN Definition")
    pnts = child(definitions[0], "Pnts")
    faces = child(definitions[0], "Faces")
    if pnts is None or faces is None:
        raise ValueError(
            f"Surface '{surface_name}' does not contain a TIN Definition with Pnts and Faces."
        )

    point_nodes = list(descendants(pnts, "P"))
    if not point_nodes:
        raise ValueError(f"Surface '{surface_name}' contains no TIN points.")

    xyz = np.empty((len(point_nodes), 3), dtype=np.float64)
    id_to_idx = {}
    for i, p in enumerate(point_nodes):
        try:
            pid = int(p.attrib["id"])
            vals = [float(v) for v in (p.text or "").split()]
            if len(vals) != 3 or not all(math.isfinite(v) for v in vals):
                raise ValueError
        except Exception as exc:
            raise ValueError(f"Invalid LandXML point at position {i + 1}.") from exc
        if pid in id_to_idx:
            raise ValueError(f"Surface '{surface_name}' has duplicate point ID {pid}.")
        xyz[i] = vals
        id_to_idx[pid] = i

    face_nodes = list(descendants(faces, "F"))
    if not face_nodes:
        raise ValueError(f"Surface '{surface_name}' contains no TIN faces.")
    face_list = []
    invisible_count = 0
    for i, f in enumerate(face_nodes):
        vals = (f.text or "").split()
        if len(vals) != 3:
            raise ValueError(
                f"Surface '{surface_name}' face {i + 1} must reference exactly three point IDs."
            )
        try:
            refs = [int(v) for v in vals[:3]]
            triangle = [id_to_idx[r] for r in refs]
        except KeyError as exc:
            raise ValueError(
                f"Face {i + 1} references missing point ID {exc.args[0]}."
            ) from exc
        except ValueError as exc:
            raise ValueError(
                f"Surface '{surface_name}' face {i + 1} contains a noninteger point ID."
            ) from exc
        invisible = f.attrib.get("i", "0")
        if invisible not in {"0", "1"}:
            raise ValueError(f"Face {i + 1} has unsupported invisible flag {invisible!r}.")
        if invisible == "1":
            invisible_count += 1
        else:
            face_list.append(triangle)
    if not face_list:
        raise ValueError("No valid triangular faces were found.")

    farr = np.asarray(face_list, dtype=np.int32)
    b = xyz[:, :2]
    cs = child(root, "CoordinateSystem")
    epsg = cs.attrib.get("epsgCode") if cs is not None else None
    meta = {
        "surface_name": surface_name,
        "point_count": len(xyz),
        "face_count": len(farr),
        "epsg": epsg,
        "declared_crs": document.declared_crs,
        "horizontal_unit": document.horizontal_unit,
        "vertical_unit": document.vertical_unit,
        "warnings": document.warnings,
        "invisible_faces": invisible_count,
        "surface_index": surfaces.index(surface),
        "bounds": (
            float(b[:, 0].min()),
            float(b[:, 1].min()),
            float(b[:, 0].max()),
            float(b[:, 1].max()),
        ),
        "zmin": float(xyz[:, 2].min()),
        "zmax": float(xyz[:, 2].max()),
    }
    if progress:
        progress(12, f"Loaded {len(xyz):,} vertices and {len(farr):,} triangles.")
    return xyz, farr, meta

