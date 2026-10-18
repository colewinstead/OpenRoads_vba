# SPDX-License-Identifier: GPL-2.0-or-later
"""Scale-aware CAD labels and physical gaps; source contours are never smoothed."""
import math
import numpy as np


def point_at(points, distances, station):
    i = min(len(points) - 2, max(0, int(np.searchsorted(distances, station, side="right") - 1)))
    t = (station - distances[i]) / (distances[i + 1] - distances[i])
    return points[i] + t * (points[i + 1] - points[i])


def slice_path(points, distances, start, end):
    inside = (distances > start) & (distances < end)
    return np.vstack((point_at(points, distances, start), points[inside],
                      point_at(points, distances, end)))


class CollisionIndex:
    """Conservative axis-aligned envelopes of rotated text, indexed in model space."""
    def __init__(self, cell):
        self.cell = cell
        self.cells = {}

    def keys(self, box):
        x0, y0, x1, y1 = box
        for x in range(math.floor(x0 / self.cell), math.floor(x1 / self.cell) + 1):
            for y in range(math.floor(y0 / self.cell), math.floor(y1 / self.cell) + 1):
                yield x, y

    def add_if_clear(self, box):
        keys = list(self.keys(box))
        for key in keys:
            for other in self.cells.get(key, ()):
                if box[0] <= other[2] and box[2] >= other[0] and box[1] <= other[3] and box[3] >= other[1]:
                    return False
        for key in keys:
            self.cells.setdefault(key, []).append(box)
        return True


def place_labels(contours, *, mode="major", spacing=250, height=5, clearance=1,
                 break_lines=True, widths=None, max_labels=100, box_height=None, minor_height=None,
                 minor_box_height=None, cancel=lambda: False):
    if mode not in {"none", "major", "both"}:
        raise ValueError("Label mode must be none, major or both")
    if not all(math.isfinite(v) and v > 0 for v in (spacing, height, clearance)):
        raise ValueError("Label spacing, height and clearance must be finite and positive")
    if not isinstance(max_labels, int) or not 1 <= max_labels <= 10000:
        raise ValueError("max_labels must be an integer in 1..10000")
    box_height = height if box_height is None else box_height
    if not math.isfinite(box_height) or box_height < height:
        raise ValueError("Text box height must be finite and >= nominal height")
    minor_height = height if minor_height is None else minor_height
    minor_box_height = (box_height if minor_height == height else minor_height) if minor_box_height is None else minor_box_height
    if not math.isfinite(minor_height) or minor_height <= 0 or not math.isfinite(minor_box_height) or minor_box_height < minor_height:
        raise ValueError("Minor text height must be positive and its box height >= nominal height")
    widths = widths or {}
    index = CollisionIndex(max(max(height, minor_height) * 10, spacing / 2))
    output = {}
    # Major labels have collision priority; preserve original contour ordering in output.
    for contour in sorted(contours, key=lambda c: (not c["major"], c["id"])):
        if cancel():
            raise InterruptedError("Label placement cancelled")
        text_height = height if contour["major"] else minor_height
        text_box_height = box_height if contour["major"] else minor_box_height
        points = contour["points"]
        distances = np.r_[0, np.cumsum(np.linalg.norm(np.diff(points, axis=0), axis=1))]
        length = distances[-1]
        labels, gaps = [], []
        enabled = mode == "both" or (mode == "major" and contour["major"])
        # shortcut: CLI without ORD measurements reserves one height per character;
        # ORD supplies measured numeric widths before generating production geometry.
        width = float(widths.get(contour["text"], len(contour["text"]) * text_height))
        if not math.isfinite(width) or width <= 0:
            raise ValueError("Text width must be finite and positive")
        half_gap = (width + 2 * clearance) / 2
        if enabled and length >= max(4 * half_gap, 4 * text_height):
            count = min(max_labels, max(1, int(length / spacing)))
            for station in (np.arange(count) + 0.5) * length / count:
                if station < half_gap or length - station < half_gap:
                    continue
                sample = slice_path(points, distances, station - half_gap, station + half_gap)
                vector = sample[-1] - sample[0]
                chord = np.linalg.norm(vector)
                if chord <= 0:
                    continue
                unit = vector / chord
                legs = np.diff(sample, axis=0)
                norms = np.linalg.norm(legs, axis=1)
                # A label needs a nearly straight stretch across its entire footprint.
                if np.any((legs @ unit) / norms < math.cos(math.radians(15))):
                    continue
                center = point_at(points, distances, station)
                theta = math.atan2(unit[1], unit[0])
                theta = (theta + math.pi / 2) % math.pi - math.pi / 2
                rx = (abs(math.cos(theta)) * width + abs(math.sin(theta)) * text_box_height) / 2 + clearance
                ry = (abs(math.sin(theta)) * width + abs(math.cos(theta)) * text_box_height) / 2 + clearance
                box = (center[0] - rx, center[1] - ry, center[0] + rx, center[1] + ry)
                if not index.add_if_clear(box):
                    continue
                labels.append({"x": float(center[0]), "y": float(center[1]), "rotation": theta,
                               "height": text_height, "width": width, "text": contour["text"],
                               "box": box})
                if break_lines:
                    gaps.append((station - half_gap, station + half_gap))
        pieces, start = [], 0.0
        for a, b in gaps:
            if a > start:
                pieces.append(slice_path(points, distances, start, a))
            start = max(start, b)
        if start < length:
            pieces.append(slice_path(points, distances, start, length))
        output[contour["id"]] = {**contour, "pieces": pieces, "labels": labels,
                                "short": length < 4 * half_gap}
    return [output[c["id"]] for c in contours]
