# Fixtures

Raster fixtures are generated at test time by `ProcessingTests.raster` into
TemporaryDirectory: north-up 2x2/40x40 Float64 grids, explicit EPSG:26916, 10-metre
pixels, constant/ramp/NoData values and a separately declared synthetic NAVD88
vertical contract. This declaration is test data, not actual survey evidence.

`test_crs.request` generates independent projected input coordinates through
pyproj's EPSG definitions. `test_usgs.item` and Response provide bounded synthetic
TNM metadata and HTTP bodies. No real terrain tiles or private DGNs are included.
