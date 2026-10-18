# SPDX-License-Identifier: GPL-2.0-or-later
"""Extract piecewise-linear contours on the ORIGINAL triangles with Matplotlib."""
from decimal import Decimal, ROUND_CEILING, ROUND_FLOOR
import math
import numpy as np
from matplotlib.figure import Figure
from matplotlib.tri import Triangulation


def level_values(zmin, zmax, minor, major, base=0):
    values = [minor, major, base, zmin, zmax]
    if not all(math.isfinite(v) for v in values) or minor <= 0 or major < minor:
        raise ValueError("Intervals must be finite and major >= minor > 0")
    step, big, origin = map(lambda v: Decimal(str(v)), (minor, major, base))
    if big % step:
        raise ValueError("Major interval must be an integer multiple of minor interval")
    low = int(((Decimal(str(zmin)) - origin) / step).to_integral_value(rounding=ROUND_CEILING))
    high = int(((Decimal(str(zmax)) - origin) / step).to_integral_value(rounding=ROUND_FLOOR))
    if high - low > 10000:
        raise ValueError("More than 10,000 contour elevations; increase interval")
    result = []
    for i in range(low, high + 1):
        exact = origin + i * step
        value = float(exact)
        text = format(exact, "f")
        if "." in text:
            text = text.rstrip("0").rstrip(".")
        result.append((value, (exact - origin) % big == 0, "0" if value == 0 else text))
    if len({v for v, _, _ in result}) != len(result):
        raise ValueError("Interval is below floating-point elevation resolution")
    return result


def extract(points, faces, minor=1, major=5, base=0, cancel=lambda: False):
    levels = level_values(float(points[:, 2].min()), float(points[:, 2].max()),
                          minor, major, base)
    if not levels or points[:, 2].min() == points[:, 2].max():
        return []
    # Subtract a local origin internally to protect State Plane precision.
    origin = points[:, :2].min(axis=0)
    xy = points[:, :2] - origin
    mesh = Triangulation(xy[:, 0], xy[:, 1], triangles=faces)
    mesh.get_trifinder()  # Native validation rejects overlapping/invalid triangulations.
    figure = Figure()
    ax = figure.add_subplot()
    result = []
    try:
        # Work in bounded batches; cancellation is checked between native calls.
        for offset in range(0, len(levels), 32):
            if cancel():
                raise InterruptedError("Contour extraction cancelled")
            batch = levels[offset:offset + 32]
            cs = ax.tricontour(mesh, points[:, 2], levels=[v for v, _, _ in batch])
            for (elevation, major_contour, text), paths in zip(batch, cs.allsegs):
                for vertices in paths:
                    if len(vertices) < 2:
                        continue
                    vertices = np.asarray(vertices) + origin
                    vertices = vertices[np.r_[True, np.any(np.diff(vertices, axis=0) != 0, axis=1)]]
                    if len(vertices) < 2:
                        continue
                    result.append({"id": len(result) + 1, "elevation": elevation,
                                   "major": major_contour, "text": text, "points": vertices})
            cs.remove()
    finally:
        figure.clear()
    return result
