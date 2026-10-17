# Repeatable test plan and evidence record

## Evidence at implementation (2026-10-17)

| Check | Environment | Result |
| --- | --- | --- |
| Python offline suite | Mac QGIS 4.2.2 bundled Python 3.12, GDAL 3.13.3, pyproj 3.7.2 | 32 tests passed; no skips |
| Python offline suite | Mac system Python 3.14 + pyproj, no GDAL | 15 passed, 16 skipped; skips are not passes |
| Python offline suite | Windows test PC, Python 3.14 + pyproj, no GDAL | 15 passed, 16 skipped |
| Windows launcher | Windows PowerShell, real Python child, temporary Unicode/space/ampersand paths | Passed missing interpreter, worker error, status/log encoding |
| PowerShell syntax | PowerShell parser on Mac | Passed |
| Live official API smoke | Python client with verified TLS, metadata only | 1-m: 2, 1/3-arc-second: 4, LiDAR: 10 candidates for the test box |
| Real USGS terrain-file transfer/processing | Mac QGIS GIS worker, two adjacent 1-m tiles, 50 MiB cap | Passed search, actual transfer/hash checks, mosaic, clip, grid reprojection, metre/survey-foot outputs and zero-gap verification; [evidence](LIVE_TEST_RESULTS.md) |
| VBA compilation, picking, form, GCS, DGN interaction | ORD not available | **Not run** |
| Raster attachment / native terrain creation | Not implemented | Deferred |
| PDAL point-cloud inspection | PDAL not run | Unverified |

Counts above are a dated snapshot, not a promise of current coverage. Windows
and system-Python suite counts are the earlier 31-test checkpoint; the latest
32-test suite ran with the Mac GIS runtime. The extra regression is GIS-dependent. The test
box is WGS84 `[-88.82,33.42,-88.80,33.44]`. The metadata smoke test fetched JSON only. The separate live DEM pilot downloaded
35,755,526 terrain bytes (about 34.10 MiB) under its explicit 50 MiB cap.
Mac system Python needed the trusted system CA bundle (`SSL_CERT_FILE=/etc/ssl/cert.pem`);
certificate verification was never disabled. Tests do not contact USGS by default.

## Offline Python tests

In a configured GIS shell, from the repository root:

```text
python -m unittest discover -s USGSTerrainDownloader/tests -v
python -m compileall -q USGSTerrainDownloader/python USGSTerrainDownloader/tests
```

Fixtures are generated in temporary directories, normally 2x2 cells and at most
40x40 cells. No external DEM, private DGN or binary fixture is required by the offline suite.
The opt-in live pilot is separate; originals and generated rasters stay in ignored `.runs/`.

- CRS: explicit confirmations, wrong/missing CRS/unit, finite/nondegenerate bounds,
  exact foot definitions, NAD83/NAD83(2011) Mississippi zones, master-to-native
  conversions, edge densification and independent denser bounding-box probes.
- API: actual parameter names/URL encoding, nested live catalog contract, raw-offset
  pagination, deduplication, empty coverage, excess results, API errors/repeated pages,
  direct product format and HTTPS validation.
- Downloads: bounded mocked byte streams, incomplete content, unknown length,
  cancellation, atomic source naming, completed-source SHA receipt reuse and corruption.
- Jobs/bridge: manifest and stale-boundary gate, source-size confirmation, UTF-16
  status/products, successful small-raster job, incremental error report/recovery.
- GDAL: constant planes, two-tile mosaic, polygon validity, horizontal reprojection
  preserving Z, explicit metres-to-international-feet and survey-feet values,
  **compound vertical-unit read-back**, fully covered reprojected clip-edge regression, NoData gap rejection/explicit acceptance,
  missing CRS, contradictory Z units, mixed vertical datums, resolution conflicts,
  overlap rejection/explicit last-row priority, triangular polygon clipping, mid-process cancellation and output-cell caps. These are software checks, not
  geodetic control-point or survey accuracy certification.

If GDAL is missing, raster-dependent tests are **skipped**. Do not present a green
suite with skips as DEM-processing evidence. Use the same GIS Python/DLL/data
combination intended for ORD. Unittest is in the standard library; pytest is not
required.

On the Windows PC, in PowerShell:

```powershell
powershell -NoProfile -File .\USGSTerrainDownloader\tests\test_windows_bridge.ps1 -Python 'C:\full\path\python.exe'
```

The bridge test uses its own new temporary directory, deliberately invalid
requests, and the actual launcher/worker. It tests errors without downloads or
DGN access. A restrictive execution policy remains an environment failure to
resolve through approved IT policy, not a reason to bypass it.

## Optional live metadata-only API check

```text
python USGSTerrainDownloader/tests/smoke_search.py
```

This opt-in command queries the three official datasets with a small geographic
box, prints a count and two metadata summaries, and **never downloads terrain**.
Record dates and response failures. API outages/catalog changes should fail
clearly; do not substitute fixture data and call that a live pass.

## ORD 2026 acceptance gates — all pending

Use a disposable DGN and separate test output folder. Record full ORD build,
Windows build, Python/GDAL/PROJ versions and grid availability, actual model unit
definition, UORs per master unit, GCS name/EPSG/realization/modifiers, global origin,
known control coordinates and model element IDs/count. Source code review is not
an executed result. Do not change production DGNs to run these tests.

### 1. Assemble, compile, launch

Follow the exact module/class/UserForm assembly instructions in README. Use
Debug > Compile and record missing-reference/compiler errors verbatim. Inspect
F2/F1 signatures against the installed CHM, particularly `GetGCS(True)`, GCS `Name`,
shape `AsVertexList.GetVertices`, `ActiveDesignFile.Path`, command callbacks and
modeless behavior. A GCS-probe failure must lead to explicit manual confirmation,
not a made-up CRS. Save the `.mvba` only through Bentley's editor.

Run the documented VBA LOAD and VBA RUN key-ins. Test a second invocation,
close/reopen, Escape/X, read-only DGN, sheet/drawing models and invalid settings.
Expected: design models accepted even if read-only (this tool does not write
geometry); sheet/drawing models rejected; one session; no autoload/workspace edits.

### 2. Rectangle dynamics and existing shape

In plan, pick two independently recorded known grid coordinates. Test reverse
corner order, snapping, rotated views and a 3D design model. Expected: rectangle
aligned to **model XY**, temporary dynamic preview using supplied DrawMode, no
persistent elements, no source coordinate rounding, accepted XY identical to
recorded master coordinates. Zero width/height rejected. Reset before first/after
first corner, start another command during selection, pan/zoom and switch model.
Confirm safe form restoration and no acceptance of a moving cursor on cancellation.

Locate and accept a simple closed triangular/polygonal shape with a concavity;
verify every original XY vertex in request.json and its eventual cutline. Try cells,
complex/curved shapes, holes, reference elements, nonhorizontal shapes and
self-intersections. Unsupported types must be rejected; an invalid polygon must
never produce a released terrain raster. Compare element IDs/count and active
level/color/style before/after every selection/cancel/search/process.

Fence is not implemented: no enabled fence action or claims of fence acceptance.

### 3. CRS / known-point / working-unit gate

Use independently verified project control, not merely a forward/inverse test
within the same library. Test Mississippi East and West in NAD83 and NAD83(2011),
metre and US survey foot definitions, plus international-foot master display over
an explicitly defined CRS. Confirm the exact EPSG/code realization against model
properties. The form must show detected name or unavailable, never silently map a
Bentley alias. Missing/ambiguous GCS must require explicit source confirmation.

Compare request XY with ORD coordinate readout and known control in models with
nondefault global origin. Confirm no extra origin/UOR adjustment is needed. Test
working units differing from native GCS units; the reported model-to-native scale
must equal the exact physical-unit ratio. Wrong foot type/wrong zone must not be
accepted based solely on plausible map location. Local GCS modifiers/ground scales
must remain unsupported unless a future validated adapter is implemented.

Verify geographic extent against an independent GIS/control workflow. Stop if
unknown modifiers or residuals exceed project criteria; do not relax guards to
force a pass. Repeat with required PROJ grids unavailable and ensure fail-closed
operation selection. Record expected transformation accuracy separately from
source DEM resolution/vertical accuracy.

### 4. Form / async / configuration

At 100%, 150% and 200% Windows display scaling, inspect labels, controls, scrollbar,
keyboard/tab order and status/error wrapping. In the Immediate window check `TerrainNumber(0.5)` and `TerrainNumber(-0.5)` produce valid JSON numbers with leading zeroes. Verify all fields/actions are
accessible. Test decimal-comma Windows input and output JSON period syntax.
Select multiple rows and confirm counts, known-size sums, unknown counts, dates,
metadata links/raw descriptions and no automatic newest-product selection.

Check runtime path persistence, invalid Python/JSON/DLL/GDAL paths, signed-script
policies and writable project/default fallback destinations. Paths with spaces,
Unicode, ampersands and apostrophes must remain data. Check ORD environment and
ProjectWise/VBA settings before/after. Search, Download and Processing states must
remain distinct; controls must disable during work while Cancel and ordinary ORD
UI operations remain responsive. Switch model during work: no DGN action may use
stale context; further requests must require a new session in the intended model.

Cancel during API, download, warp and unit conversion; X/Escape must signal cancel
without unloading the live worker's form. Check terminal status and source
retention. Some atomic GDAL/read-back steps can finish before cancellation is
observed; no half-written `terrain.tif` may appear. Check a killed worker, stale
source lock and restart recovery manually.

### 5. Real USGS DEM pilot (explicitly budgeted manual test)

Search a deliberately small known project boundary. Before downloading, inspect
actual product metadata, select the smallest necessary compatible collection,
confirm vertical CRS/unit with source-specific evidence, and set a suitable hard
byte cap. This manual pilot can involve hundreds of MB despite a small boundary;
never add it to automated tests. Record metadata URLs and product versions.

Verify originals with GDAL/QGIS after transfer, source hash receipts, expected
coverage, horizontal CRS, acquisition/publication distinction and vertical
interpretation. Mosaicking incompatible grids/CRSs/resolutions/datums must stop.
For overlap, choose two identifiable products and confirm the explicitly accepted
last-row priority; no assumption of newest=best is permitted. Test incomplete HTTP
transfer and reconnect using a controlled test server/mock rather than corrupting
real source files.

Review report and raster metadata in independent GIS tools. Compare selected
source/output cells numerically, allowing for the documented bilinear resampling.
For a same-grid unit-only run, 100 m must become approximately 328.0833333333333
US survey ft or 328.0839895013123 international ft; horizontal reprojection alone
must leave numeric Z in its source unit. Verify both band units and vertical CRS
units on reopening. NoData remains NoData, polygon exterior is clipped, and gaps
inside have the recorded gate/result. Compare known project controls to applicable
source accuracy and project-use criteria; a gross overlay is not survey validation.

### 6. Manual ORD raster/terrain import

Open output folder and metadata. With a scratch DGN, use the installed ORD raster
attachment/terrain-import workflows and record the exact UI/options supported by
that build. Validate horizontal alignment, vertical datum/unit, sample elevations,
NoData exclusion, holes and cell resolution after import. Do not treat successful
raster attachment as a successfully created native terrain model. Keep datum/unit
interpretation explicit even if ORD ignores compound GeoTIFF vertical tags.

The VBA performs no attachment or native terrain writes. Do not mark either
feature implemented from manual import success. Future automation requires
separate documented API evidence, a disposable-DGN write/rollback test and native
Civil SDK validation if applicable.

### 7. LiDAR phase boundary

Search/file-count/size and original preservation can be piloted separately. If
PDAL is explicitly configured in a backend request, inspect its actual LAS/LAZ CRS,
vertical metadata, point count and classification schema. Missing PDAL must not
be represented as successful inspection. The form says **download only**; no
unclassified/all-return point cloud may be presented as bare-earth ground. Ground
classification filtering/rasterization remains deferred and must have its own
classification/NoData/vertical test plan before promotion.

## ORD evidence record

| Gate | Full ORD build / model / units | Actual evidence | Result |
| --- | --- | --- | --- |
| Compile/load/form | Pending | Not run | Pending |
| Rectangle/dynamics/reset | Pending | Not run | Pending |
| Shape/concavity/context | Pending | Not run | Pending |
| GCS name + known-point/master-unit/origin | Pending | Not run | Pending |
| Form async/cancel/config | Pending | Not run | Pending |
| Real source DEM/report verification | No ORD; Mac GIS runtime | Two-tile pilot and numeric/unit checks passed; no survey-control comparison | Python evidence passed; ORD/control validation pending |
| Manual raster/terrain import/control | Pending | Not run | Pending |
| LiDAR real source/PDAL | Pending | Not run | Pending |
