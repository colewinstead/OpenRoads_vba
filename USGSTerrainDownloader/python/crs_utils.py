"""Explicit model-master -> CRS-native XY -> geographic boundary contract."""
import math

from pyproj import CRS, Geod, Transformer, network

network.set_network_enabled(False)
UNIT_METERS = {"m": 1.0, "ft": 0.3048, "us-ft": 1200 / 3937}
MAX_VERTICES = 20000


def require(condition, message):
    if not condition:
        raise ValueError(message)


def linear_factor(source, target):
    require(source in UNIT_METERS and target in UNIT_METERS,
            "Units must be m, ft (international), or us-ft (US survey).")
    return UNIT_METERS[source] / UNIT_METERS[target]


def horizontal_crs(value):
    crs = CRS.from_user_input(value)
    require(crs.is_projected and len(crs.axis_info) == 2,
            "Provide a two-dimensional projected horizontal CRS; no compound, vertical, or geographic DGN CRS.")
    require(all(a.direction in ("east", "north") for a in crs.axis_info),
            "Only east/north projected coordinate axes are supported.")
    factors = [a.unit_conversion_factor for a in crs.axis_info]
    require(abs(factors[0] - factors[1]) < 1e-14, "Mixed horizontal axis units are unsupported.")
    return crs


def transformer(source, target):
    for value in (source, target):
        datum = CRS.from_user_input(value).datum
        require(datum is None or "Dynamic" not in datum.type_name,
                "Dynamic CRS needs coordinate epochs; this version supports static CRSs only.")
    # Fail if PROJ's best operation needs missing grids; never allow ballpark substitution.
    return Transformer.from_crs(source, target, always_xy=True,
                               allow_ballpark=False, only_best=True)


def ring_valid(points):
    require(isinstance(points, list) and 3 <= len(points) <= MAX_VERTICES,
            "Boundary needs 3 to 20,000 XY vertices.")
    ring = []
    for point in points:
        require(len(point) == 2, "Boundary vertices must contain XY only.")
        p = [float(v) for v in point]
        require(all(math.isfinite(v) for v in p), "Nonfinite boundary coordinate.")
        ring.append(p)
    if ring[0] != ring[-1]:
        ring.append(ring[0][:])
    require(len(ring) >= 4, "Boundary has fewer than three distinct corners.")
    require(all(a != b for a, b in zip(ring, ring[1:])), "Duplicate adjacent boundary vertices.")
    # Translate before shoelace to avoid cancellation at large State Plane coordinates.
    x0, y0 = ring[0]
    area = sum((a[0]-x0)*(b[1]-y0)-(b[0]-x0)*(a[1]-y0)
               for a, b in zip(ring, ring[1:])) / 2
    require(abs(area) > 1e-10, "Boundary has zero XY area.")
    return ring


def densify(ring, native_to_meters, segment_m=100):
    result = []
    for a, b in zip(ring, ring[1:]):
        n = max(1, math.ceil(math.dist(a, b) * native_to_meters / segment_m))
        require(len(result) + n < MAX_VERTICES, "Boundary too large/complex; reduce the area.")
        result.extend([[a[0]+(b[0]-a[0])*i/n, a[1]+(b[1]-a[1])*i/n]
                       for i in range(n)])
    result.append(result[0][:])
    return result


def prepare_boundary(config):
    require(config.get("crs_confirmed") is True,
            "Confirm source CRS, master-unit type, and absence of local/ground transforms.")
    require(config.get("coordinates_are_grid") is True,
            "Local offsets, rotation, ground scale, or custom Bentley GCS modifiers require a validated adapter.")
    crs = horizontal_crs(config["source_crs"])
    unit = config.get("master_unit")
    require(unit in UNIT_METERS, "Explicit model master unit required: m, ft, or us-ft.")
    native_factor = crs.axis_info[0].unit_conversion_factor
    scale = UNIT_METERS[unit] / native_factor
    ring = ring_valid(config["boundary"])
    ring = [[x*scale, y*scale] for x, y in ring]
    dense = densify(ring, native_factor)
    operation = transformer(crs, "EPSG:4326")
    xs, ys = operation.transform(*zip(*dense), errcheck=True)
    operation = operation.get_last_used_operation()
    require(all(math.isfinite(x) and math.isfinite(y) and -180 <= x <= 180 and -90 < y < 90
                for x, y in zip(xs, ys)), "Transformed boundary is not valid longitude/latitude.")
    bbox = [min(xs)-1e-7, min(ys)-1e-7, max(xs)+1e-7, max(ys)+1e-7]
    require(bbox[2]-bbox[0] <= 2 and bbox[3]-bbox[1] <= 2,
            "Extent exceeds the 2-degree safety limit (antimeridian areas unsupported).")
    area, _ = Geod(ellps="WGS84").polygon_area_perimeter(xs, ys)
    require(0 < abs(area) <= 1e10, "Geographic area must be positive and <= 10,000 square km.")
    bounds = crs.area_of_use
    if bounds:
        require(all(bounds.west-0.01 <= x <= bounds.east+0.01 and
                    bounds.south-0.01 <= y <= bounds.north+0.01 for x, y in zip(xs, ys)),
                "Boundary lies outside the CRS area of use. Verify CRS and model coordinates.")
    return {"crs_wkt": crs.to_wkt(), "crs_name": crs.name,
            "horizontal_unit": crs.axis_info[0].unit_name,
            "model_to_native_scale": scale, "ring": dense, "bbox": bbox,
            "area_m2": abs(area), "search_operation": operation.description,
            "search_operation_accuracy_m": operation.accuracy,
            "search_pipeline": operation.definition}
