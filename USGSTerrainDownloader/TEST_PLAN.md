# Repeatable test plan and evidence record

## Evidence at implementation (2026-10-17)

| Check | Environment | Result |
| --- | --- | --- |
| Python offline suite | Mac QGIS 4.2.2 bundled Python 3.12, GDAL 3.13.3, pyproj 3.7.2 | 32 tests passed; no skips |
| Python offline suite | Mac system Python 3.14 + pyproj, no GDAL | 15 passed, 16 skipped; skips are not passes |
| Python offline suite | Windows test PC, Python 3.14 + pyproj, no GDAL | 15 passed, 16 skipped |
| Python offline suite | ORD Windows computer, OSGeo4W Python 3.12, GDAL 3.11.3, pyproj 3.7.0, numpy 1.26.4 | **32 passed, no skips** on 2026-10-17 |
| Metadata-autofill revision | Same Windows GIS runtime | **40 passed, no skips**; explicit body/provenance, unknown/negated/conflicting metadata, TSV propagation and pre-transfer contradiction checks included |
| Float32 delivery revision | Same Windows GIS runtime | **41 passed, no skips**; Float32 type/finite NoData, exact rounded samples, reported rounding, overflow/rounded-sentinel rejection and working-file cleanup; sandboxed attempt lacked TEMP permissions, rerun with TEMP access passed |
| Cell-size default revision | Same Windows GIS runtime | **42 passed, no skips**; search suggestions tested in metres, international feet and US survey feet; no assumed spacing for angular DEM/LiDAR datasets |
| Native cell-size defaults | ORD 26.00.00.154 / Default / 2D; saved metadata replay without network | **Passed** survey-foot/metric suggestions, manual override, blank refill, replacement/clearing of prior automatic values, independence from elevation unit, unchecked grid confirmation and unchanged DGN IDs/settings; empty Unicode reader regression fixed |
| Actual Windows DEM processing / storage import diagnosis | Three Desktop job reports; cached original | All three dem_complete/zero gaps; user reported Float64 imports failed and original Float32 imports; separate Float32/-999999 diagnostic imported successfully per user, same compound CRS/units/grid |
| Corrected worker derivatives | Desktop/USGS_ORD_fixed, same three original requests/sources | Regenerated without network; GDAL read-back passed; user confirmed small worker output imports successfully; every sample/grid/CRS/unit matches the importing diagnostic; larger copies and known control pending |
| Source EPSG suggestion | ORD 26.00.00.154 / Default / 2D | **Passed** native parser/default assertions and user confirmation of EPSG:6576; editable suggestion, verification checkbox initially unchecked |
| Selected-product vertical autofill | Same model; actual saved two-row live-search TSV loaded through native form code | **Passed** m/EPSG:5703, both full citations, protected detected fields, deselection clearing/manual unlocking, restoration and unchanged IDs/settings; temporary loader/test code removed |
| Updated suite including standalone launcher | Same Windows GIS runtime | **37 passed, no skips**; five new real-EXE tests cover Unicode/metacharacter paths, invalid inputs, environment isolation, large logs, exit codes and cancellation |
| Native standalone bridge | ORD 26.00.00.154 / Default / 2D | **Passed** actual VBA -> EXE -> configured Python -> expected invalid-schema exit/status; no network or DGN writes |
| Live ORD dialog metadata search | Same model; standalone EXE + Windows GIS runtime | **Passed**; user saw products and actual job read-back confirmed search_complete, 2 dem1m candidates, 8-vertex boundary, CRS 6576/us-ft, explicit grid flags, no vertical declaration or data transfer |
| Installed Bentley API inspection | ORD 26.00.00.154; delivered CHM and interop DLL 26.00.00.974 | Signatures inspected; not compilation or graphical execution |
| Windows launcher | Windows PowerShell, real Python child, temporary Unicode/space/ampersand paths | Passed missing interpreter, worker error, status/log encoding |
| PowerShell syntax | PowerShell parser on Mac | Passed |
| Live official API smoke | Python client with verified TLS, metadata only | 1-m: 2, 1/3-arc-second: 4, LiDAR: 10 candidates for the test box |
| Real USGS terrain-file transfer/processing | Mac QGIS GIS worker, two adjacent 1-m tiles, 50 MiB cap | Passed search, actual transfer/hash checks, mosaic, clip, grid reprojection, metre/survey-foot outputs and zero-gap verification; [evidence](LIVE_TEST_RESULTS.md) |
| VBA assembly / Debug Compile / targeted Save | ORD 26.00.00.154, user's disposable test.dgn | **Passed**; real 199,680-byte local MVBA saved; execution gates remain separate |
| Native 2D self-check | ORD 26.00.00.154 / Default / GCS name TN83/2011F | Passed after initialization fix; form/GCS/default guards/shape/boundary and unchanged IDs/settings; not graphical picking/control validation |
| Normal dialog launch | Same host; user-entered documented key-ins | User confirmed dialog stays open; duplicate loaded test copies noted below |
| Graphical rectangle capture / preview / reset | Same 2D test model; user clicks and screenshot | User confirmed capture, visible preview, completion, reset/dialog return, outline cleanup and prior-boundary preservation; earlier invisible-preview cause unproven |
| Graphical existing-shape selection | Same 2D test model; user clicks | User confirmed acceptance returns to dialog with Boundary captured and vertex count; numeric/control/rejection cases remain untested |
| Runtime/master-unit autofill | ORD 26.00.00.154, single delivered project | **Passed** actual compile and manual-F5 native regression; exact us-ft and three runtime paths, manual confirmations retained, IDs/settings unchanged |
| Additional GCS metadata / grid confirmation | Same 2D test model | Actual COM Source=EPSG:6576, datum NAD83/2011, PaperScale=1; local EPSG catalog matches; user verified unmodified State Plane grid; no numeric control comparison |
| Managed full-GCS diagnostic | Installed .NET Framework 4.8 / ORD API DLLs | C# compile passed; two explicit full-path MDL loads failed, no report; WKT/JSON execution unproven |
| Raster attachment / native terrain creation | Not implemented | Deferred |
| PDAL point-cloud inspection | PDAL not run | Unverified |

## Windows ORD computer preflight (2026-10-17)

Current build after the cell-size default: MVBA **538,112 bytes**, SHA-256
`b60be587c0826fafeb7105ff9dace5b2bab4281b674dd4b90987675d15b905f4`;
companion EXE **8,192 bytes**, SHA-256
`d4ed77b988692b40b6ba1405abecdbe1c03270ffd3414b9dec0c1749fb015468`.
Actual compile/targeted save and the native form/launcher checks passed; five
production components remain. Final compile was completed by the user and confirmed
through the disabled terrain-specific compile command; targeted Save returned
Saved=True. Evidence: `.runs/ord-20261017/cell-size-build.json`.
The older PowerShell restriction and build snapshots below are historical results;
current VBA starts the EXE directly. Live dialog Search also passed; actual job
read-back confirmed two candidates and empty stdout/stderr logs. Evidence:
`.runs/ord-20261017/live-standalone-search.json`. Subsequent Desktop job reports
prove Windows processing completed; after the storage fix, the user confirmed
the small regenerated worker output imports. Larger copies and known control
remain untested in ORD; see the evidence rows above.

- ORD reports **Version 26.00.00.154** through actual COM activation. Windows
  build is **26100.9457**, display version 24H2. No production DGN was opened.
- `CreateDesignFile` opened the disposable
  `.runs/ord-20261017/acceptance.dgn` from the delivered `3D Imperial Design.dgn`
  seed. COM reported model `3D Imperial Design` and **3048 UOR/master unit**.
  This resolution alone does not establish a foot definition or grid mapping.
  `GetGCS(True)` returned Nothing; no CRS was guessed or assigned.
- CHM and interop inspection confirmed the signatures detailed in
  [API_NOTES.md](API_NOTES.md). This did not exercise the application's form,
  `CallByName` probe, graphical pickers or model-context guards.
- Windows GIS offline suite: **32 tests passed, no skips**. Command:
  `cmd.exe /d /c "call C:\Users\cwinstead\AppData\Local\Programs\OSGeo4W\bin\o4w_env.bat && python -m unittest discover -s USGSTerrainDownloader/tests -v"`.
  The first sandboxed run failed with temporary-directory access errors through
  OSGeo4W's short user path; the approved run outside that sandbox passed.
  GDAL emitted `transformer options does not support option STRIP_VERT_CS` warnings;
  the horizontal-only Z and vertical-unit read-back tests nevertheless passed.
  This does not replace the Windows real-source/control tests.
- Environment capture and `compileall` succeeded. The machine-specific Python
  path and child environment JSON are in the ignored `.runs/ord-20261017/`
  folder. They are not portable configuration or committed defaults.
- **Launcher gate blocked:** the actual `powershell.exe -NoProfile
  -NonInteractive -File USGSTerrainDownloader/tests/test_windows_bridge.ps1
  -Python <OSGeo4W python.exe>` invocation failed before test execution with
  `running scripts is disabled on this system` / `UnauthorizedAccess`.
  All reported execution-policy scopes were Undefined. No policy/settings were
  changed and no bypass flag was used. Resolve this through the organization's
  approved policy before claiming a form-to-worker pass.
- The desktop automation helper failed with `Computer Use native pipe is
  unavailable ... (os error 2)` after retry and session reset. The user chose
  API/COM checks only. Graphical acceptance remains pending.
- In the automation-only ORD session, `VBE.VBProjects.Add(1)` failed with
  `Invalid index` (`0x8002000B`), documented `VBA CREATE` exposed no project,
  and opening a copied project through `VBProjects.Open` failed with
  `0x800A01B8`. The copied file was renamed `container-template.mvba` and is
  **only an unchanged copy of PatternLineGenerator**, not this application's build.
  Normal-host assembly is a separate pending check. No compiled terrain MVBA is
  supplied or claimed from these attempts. The normal host loaded the existing
  `SquareYard.mvba` from its current configuration; that project was not edited.
- Normal-host `OpenDesignFile` did not return an active DGN; the user reported
  ORD crashed/reopened. The waiting COM call was cancelled and further host
  commands stopped. The latest Bentley `ExceptionHistory-1/Exception.log`
  records **0xc0000005 at 15:30:02**, with `acceptance.dgn` active and no active
  model name. The terrain project had not been assembled or run. This records
  an actual host failure but does not identify its cause.
  Application events at 15:30:06/15:30:08 record a separate
  Bentley `ArchiveDgnFiles.exe` failure (`System.ArgumentException` while creating
  a ZIP entry, exception code `0xe0434352`). That event does **not** establish the
  cause of the reported ORD crash. Resume only after the user opens a disposable
  DGN normally in the intended existing workspace/workset; verify its full path
  before assembly. Do not repeat automation-host restart attempts or change
  workspace/autoload configuration to force this gate.
- The user subsequently opened `test.dgn` normally in the intended managed
  workspace. Read-only COM reported ORD 26.00.00.154, model `Default`, model type
  0 (normal design), **10000 UOR/master**, and raw `GetGCS(True).Name` =
  **TN83/2011F**. This is not an EPSG/unit/modifier/control validation. The four
  existing projects (ViewSet, SnappableToggle, Survey, WorkSpaceVersion) were
  left untouched. In this host `VBA CREATE` again created no project and Message
  Center reported `Unknown key-in or command`; `VBProjects.Open` again returned
  `0x800A01B8`. The user was asked to create only a blank standalone project via
  the editor's New Project command before further COM import/compile attempts.
  No ProjectWise, workspace, VBA search-path or autoload settings were changed.
- The user created the blank terrain project using the editor. COM then imported
  the three standard modules (types 1), picker class (type 2) and blank UserForm
  (type 3), inserting the supplied form code. VBA, MicroStationDGN, stdole and
  MSForms references all reported `IsBroken=False`.
- **Actual compile/save passed:** with terrain as the active project, the editor's
  `Compi&le USGSTerrainDownloader` action (CommandBar ID 578) returned, changed
  from enabled to disabled, and left project mode 2 (design). The targeted File
  Save action named the terrain file and returned `Saved=True`; the file was
  **199,680 bytes**. `VBProject.SaveAs` is unsupported for this host project;
  SAVEALL was not used. Hashing while ORD holds the project open failed with a
  sharing violation. Local evidence: `.runs/ord-20261017/compile-result.json`.
- A temporary `modTerrainORDSmoke` module was added after that saved snapshot
  for native form/GCS/in-memory-shape/boundary/element-ID/settings checks. It
  compiled, but the automated Run command produced no report. Manual F5 actually
  started it and wrote `status=running`, then raised runtime **91** during form
  initialization. Debug highlighted `lblTotals.caption` in `RefreshTotals`:
  ListBox Change fired before the totals label existed. A two-line initialization
  guard was added to the shared routine, all callers were checked, and the
  corrected source compiled. The same native regression was selected for another
  manual F5; completion is **not passed until its report confirms success**.
  Initial-failure evidence: `.runs/ord-20261017/native-smoke-before-fix.json`.
  Remove this diagnostic module, recompile and save the terrain project before
  treating the final MVBA as ready for load/run testing or distribution.
- **Native check subsequently passed.** The model is 2D (`Is3D=False`); the first
  diagnostic fixture incorrectly used nonzero Z. It was corrected to use Z=0
  in 2D and Z=7 in 3D, keeping exact XYZ comparisons. The final report confirms
  the actual form/GCS/default-confirmation/shape/boundary checks and identical
  before/after element IDs, active settings and four VBA configuration values.
  A 3D file is **not a prerequisite** for selection. No 3D test was run.
  Local report: `.runs/ord-20261017/native-smoke.json` (`status=passed`).
- The temporary module was removed. The source fix remains in the UserForm;
  only the three production modules, picker class and form remain in the saved
  project. The final compiled state was up to date (Compile disabled) and targeted
  Save returned Saved=True. The included `USGSTerrainDownloader.mvba` is a
  **429,568-byte ORD-saved test build**, with SHA-256
  `283f2449d0351f934e8132070604aec5c3e0f479aba56fffbd97928ab79bafc2`.
  The copy matched the saved source hash; it was not fabricated or renamed from
  another application's project. Evidence: `.runs/ord-20261017/final-build.json`.
- The user confirmed the normal dialog stays open after the documented key-ins.
  Read-only COM found both the `.runs` project and delivered project loaded under
  the same name; both have five components. Unload the `.runs` row in Project
  Manager before repeating a clean single-project load/run gate. Leave the four
  pre-existing projects and all configuration unchanged.
- The user unloaded the older `.runs` copy; COM verified only the delivered terrain
  copy plus the four untouched existing projects remain. The autofill update then
  passed actual Debug Compile and manual-F5 `modTerrainORDTests.RunTerrainORDSmoke`.
  Report `.runs/ord-20261017/native-autofill.json` confirms exact model `us-ft`,
  detected source folder/Python/environment file, unit classification regressions,
  unconfirmed CRS/grid and vertical declarations, shape/boundary checks and unchanged
  IDs/settings. Only this application's local runtime files were captured into
  `%LOCALAPPDATA%\USGSTerrainDownloaderConfig`; no ORD configuration changed.
  The temporary test module was removed, compile was up to date, and targeted Save
  returned Saved=True. Updated artifact: **420,352 bytes**, SHA-256
  `3b425457febd2a55fa6f86ddcc1cd37295903f4419178d4230f2ccade9feafe6`.
  Evidence: `.runs/ord-20261017/autofill-build.json`. Earlier hashes/sizes above
  describe historical snapshots. The user subsequently confirmed the moving
  preview is visible; no preview-code change was made and the earlier cause is
  unproven. The user then confirmed second-click completion, right-click cancellation
  after the next first corner, dialog return, temporary-outline cleanup and retention
  of the prior boundary. Existing-shape acceptance also returned to the dialog with
  Boundary captured and a vertex count. These reports are recorded in
  `.runs/ord-20261017/graphical-rectangle.json`; they do not establish numeric
  coordinate agreement, rotated/3D behavior or unsupported-shape rejection.
- A later typed C# COM `ModelReference.Scan(null)` succeeded and showed all prior
  element IDs plus only ID 3375. Read-only inspection found an eight-vertex Shape
  on Default; the user confirmed creating a new test shape. The delta is accounted
  for, rather than claiming identical IDs across user drawing. Nothing was deleted.
  Evidence: `.runs/ord-20261017/after-graphical-ids.json` and `new-element-3375.txt`.

### Read-only managed GCS diagnostic (compile only)

`tests/OrdGcsProbe.cs` compiles with the installed Framework64 `csc.exe`, target
library/x64, `System.Web.Extensions.dll`, and these ORD 2026 references:
root `ustation.dll`, `Bentley.DgnPlatformNET.dll`, `Bentley.GeometryNET.dll`,
`Bentley.GeometryNET.Structs.dll`, `Bentley.GeoCoord2.dll`, `Bentley.DgnGeoCoord2.dll`;
`Assemblies\Bentley.MicroStation.Interfaces.1.0.dll`;
`Assemblies\ECFramework\Bentley.General.1.0.dll`.
The output stays in ignored `.runs/ord-20261017/OrdGcsProbe.dll`; it is not part
of the application. Two explicit full-path `MDL LOAD` attempts, including an
explicit USGSGCSPROBE task ID, returned `MDL Loader: Could not load application`.
No `%TEMP%\USGSTerrainDownloader-GcsProbe.json` exists. Compilation passed;
managed WKT/JSON API execution did not. No search path, autoload, installation,
ProjectWise/workspace or execution-policy setting was changed. See
`.runs/ord-20261017/gcs-probe-load.json`. The working MVBA remains independent.

### Repeat the native regression

In a fresh disposable session with no active downloader job, import
`tests/modTerrainORDTests.bas` into a test copy of the terrain project and run
`modTerrainORDTests.RunTerrainORDSmoke` with F5. It briefly opens/closes the form
and writes `%TEMP%\USGSTerrainDownloader-ORDTests.json`. It asserts the native
startup/GCS/default-confirmation/shape/boundary behavior and unchanged model
element IDs/settings. Remove the test module before saving a deliverable build.
The executed `.runs` version used an absolute report path. The checked-in test
uses TEMP and adds exact master-unit classification and runtime/form defaults;
that extended test also passed in the single delivered project. The final application
contains no test module.

Counts above are a dated snapshot, not a promise of current coverage. The older
no-GDAL Windows and system-Python counts are the earlier 31-test checkpoint;
the latest 32-test suite passed with both the Mac GIS and Windows OSGeo4W runtimes.
The extra regression is GIS-dependent. The test
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

On Windows, from the configured GIS shell, build and test the standalone launcher:

```bat
python USGSTerrainDownloader\python\build_launcher.py
python -m unittest discover -s USGSTerrainDownloader\tests -p test_windows_launcher.py -v
```

The five real-EXE tests use disposable Unicode/space/apostrophe/ampersand paths,
the actual worker's invalid-schema path and isolated fixture workers. They verify
child-only environment settings/allowlisting, forced PROJ_NETWORK=OFF, UTF-8
stdout/stderr beyond pipe-buffer size, UTF-16 status, missing interpreter, exact
exit propagation and cooperative cancel.flag/status preservation. They perform
no downloads or DGN access. The legacy PS1 bridge test remains for the old optional
PowerShell launcher; it is not the application's current launch path.

For the actual native bridge gate, import `modTerrainORDTests.bas`, compile, and
run `modTerrainORDTests.RunTerrainLauncherSmoke`. It first repeats the native form
check, then calls production `TerrainStartWorker` with a deliberately invalid
schema. Require both TEMP reports to pass: `USGSTerrainDownloader-ORDTests.json`
and `USGSTerrainDownloader-LauncherTests.json`. Expected Python exit 2 and
"Unsupported request schema" prove the process boundary, not a live search.
This check passed via an explicit `VBA RUN` key-in sent through CadInputQueue;
evidence is `native-standalone-form.json` and `native-standalone-launcher.json` in
the ignored run folder. The temporary module was removed before the final save.

## Optional live metadata-only API check

```text
python USGSTerrainDownloader/tests/smoke_search.py
```

This opt-in command queries the three official datasets with a small geographic
box, prints a count and two metadata summaries, and **never downloads terrain**.
Record dates and response failures. API outages/catalog changes should fail
clearly; do not substitute fixture data and call that a live pass.

## ORD 2026 acceptance gates

Use a disposable DGN and separate test output folder. Selection supports normal
2D and 3D design models; keep the ordinary 2D workflow as an acceptance case.
Record full ORD build,
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
| Compile/save | 26.00.00.154 / Default / 10000 UOR per master | Imported components/references, Debug Compile disabled after execution, targeted Save returned Saved=True | Passed; local test MVBA only |
| Native form/GCS/shape/boundary regression | 26.00.00.154 / Default / 2D / 10000 UOR/master | Manual F5 report passed after startup fix; unchanged model IDs/settings/configuration | Passed for these programmatic checks |
| Load/run/form | Same host | User confirmed normal dialog stays open; duplicate build copy subsequently unloaded; native single-copy autofill regression passed | Basic launch/defaults confirmed; layout/async gates pending |
| Rectangle/dynamics/reset | Same 2D test model | User confirmed capture, visible preview, completion, reset/dialog return, cleanup and retained prior boundary | User-observed checks passed; numeric/rotated/3D cases pending |
| Shape/concavity/context | Same 2D test model | User confirmed Select Existing Shape accepts and returns to dialog with captured boundary | Basic acceptance confirmed; exact geometry, concavity, rejection/context cases pending |
| GCS name + known-point/master-unit/origin | Same 2D model | Actual COM name/Source, exact numeric master-unit classification, user-confirmed grid and live search passed | Independent numeric control/global-origin and full GCS export pending |
| Form async/cancel/config | Pending | Not run | Pending |
| Standalone bridge / live Search | 26.00.00.154 / Default / 2D / us-ft | Native bridge report passed; user saw 2 products; actual request/status/manifest verified search_complete | Passed for startup and live metadata search; active-job UI cancellation/processing separate |
| Real source DEM/report verification | No ORD; Mac GIS runtime | Two-tile pilot and numeric/unit checks passed; no survey-control comparison | Python evidence passed; ORD/control validation pending |
| Manual raster/terrain import/control | Pending | Not run | Pending |
| LiDAR real source/PDAL | Pending | Not run | Pending |
