# Continue on the ORD computer

Saved 2026-10-17. Repository: https://github.com/colewinstead/OpenRoads_vba

Start here, then read [README.md](README.md), [API_NOTES.md](API_NOTES.md),
[TEST_PLAN.md](TEST_PLAN.md), and [LIVE_TEST_RESULTS.md](LIVE_TEST_RESULTS.md).
The [saved conversation](codex-session-01a1485e-079a-78b3-8e11-c54721d64042.md)
contains the original task, implementation activity, and the latest continuation.

## ORD-computer continuation (2026-10-17)

Installed ORD **26.00.00.154** CHM/type-library signatures have been inspected;
see [API_NOTES.md](API_NOTES.md). The actual `GetGCS(True)` call on an unassigned
seed returned Nothing; the actual application CallByName/form probe subsequently
passed with `TN83/2011F` in the user's 2D test model. Full GCS extraction and
known-point agreement remain unproven. Explicit CRS/unit/grid/vertical confirmations
remain required.

Additional actual read-only COM GCS properties are in `gcs-properties.json`:
Source=`EPSG:6576`, Description=`NSRS11(NAD83/2011) Tennessee State Plane Zone, US Foot`,
DatumName=`NAD83/2011`, Units=`Foot`, VerticalDatumName=`NAVD88`, FalseEasting=1968500,
FalseNorthing=0, PaperScale=1. The installed pyproj EPSG catalog independently lists
6576 as NAD83(2011) / Tennessee (ftUS), and 6575 as the metre version. This does not
prove the model has no local transforms, and the DGN's vertical datum does not
establish any selected USGS tile's vertical datum. The user confirmed
"Unmodified State Plane grid—verified". Independent known-point comparison is
still untested; user confirmation and a numeric control test are separate evidence.

The Windows OSGeo4W runtime (Python 3.12, GDAL 3.11.3, pyproj 3.7.0, numpy 1.26.4)
passed the original **32 offline tests with no skips**, environment capture and compileall.
After the standalone launcher change, the complete Windows suite passed **37 tests,
no skips**, including five real-EXE bridge tests.
Machine-specific runtime files are in ignored `.runs/ord-20261017/`.
The earlier PowerShell bridge test could not start: PowerShell reports `running scripts is
disabled on this system`. No execution-policy or workspace/autoload settings
were changed. The user explicitly authorized a standalone launcher; current VBA
uses `python\launch_worker.exe` and no longer invokes PowerShell. The old PS1/test
remain for existing command-line workflows.

**The terrain project is now assembled, compiled and saved**, with the standalone
test build at `USGSTerrainDownloader.mvba` (**395,264 bytes**, SHA-256
`133da49831ff026344f6d2efd9256601556f3e7cd6cfa205109237707d2ccd37`).
Keep the companion `python\launch_worker.exe` (**8,192 bytes**, SHA-256
`d4ed77b988692b40b6ba1405abecdbe1c03270ffd3414b9dec0c1749fb015468`) beside main.py.
The `.runs/ord-20261017/USGSTerrainDownloader.mvba` project is an older build copy,
now unloaded. The delivered project is the current compile/save target.
Actual Debug Compile returned with the command disabled; targeted Save returned
`Saved=True`. Its four references resolve. A temporary `modTerrainORDSmoke` was added for native
tests; it compiled but automated Run produced no report. Manual F5 exposed runtime
91 in form initialization (`RefreshTotals` used lblTotals before it was created).
The shared routine now guards early initialization events; the fix compiled and
the native check then passed after correcting its nonzero-Z fixture for the
actual **2D** model. `native-smoke.json` confirms form/default-confirmation/GCS/
shape/boundary checks and unchanged DGN IDs/settings/configuration. The diagnostic
was removed and the final five-component project saved before the verified copy
above was made. A portable version is `tests/modTerrainORDTests.bas`; it is not
included in the application binary. The user confirmed normal dialog launch.
Graphical picking, 3D behavior, known control, worker launch and terrain import
remain separate gates. The user confirmed rectangle capture; their screenshot
shows "Boundary captured: 5 vertices". The moving preview was initially not
visible; after unloading the duplicate and running the updated single project,
the user confirmed "The preview is visible now". Preview code was not changed;
the earlier cause is unproven. The user also confirmed second-click completion,
right-click cancellation after the next first corner, dialog return, temporary-outline
cleanup and preservation of the prior boundary. Select Existing Shape subsequently
returned to the dialog with Boundary captured and a vertex count, per the user.
These are user-observed graphical checks, not numeric/control validation.
An independent typed COM scan after the graphical tests retained all prior IDs
and found only new ID 3375, an eight-vertex Shape on Default. The user confirmed
they drew a new test shape. This accounts for the scan delta; no element was
deleted. Evidence: `after-graphical-ids.json`, `new-element-3375.txt`. Typed C#
interop `ModelReference.Scan(null)` works where PowerShell late binding failed;
in C# `AsVertexList()` is a method, while VBA allows the no-parentheses spelling.

The user requested autofill of the highlighted fields. Source changes now discover
the source directory from the single loaded terrain project, read the documented
local runtime capture, and suggest master units from the numeric metre definition.
The confirmation checkbox and all vertical declarations remain manual. The extended
native test covers metre/international-foot/survey-foot/custom/unitless/invalid
unit definitions and form defaults. Actual Debug Compile passed, then manual F5
produced `native-autofill.json` with status passed: exact `us-ft` detection, all
three runtime paths present, manual confirmations retained, identical DGN IDs and
checked settings. The temporary test module was removed, the five-component
application recompiled and targeted Save returned Saved=True. The binary/hash above
is the subsequent standalone-launcher build. Historical autofill evidence is
`autofill-build.json`; current evidence is `standalone-build.json`.

The standalone EXE is built from `python/WorkerLauncher.cs` by
`python/build_launcher.py` using the installed Framework64 compiler. It applies
only the existing child-environment allowlist, forces PROJ_NETWORK=OFF and UTF-8
Python IO, reads both pipes asynchronously, preserves status/cancellation and
propagates exit codes. The five Windows tests passed, then actual ORD
`RunTerrainLauncherSmoke` passed: production VBA -> WScript Exec -> GUI-subsystem
EXE -> configured GIS Python -> expected invalid-schema exit 2/status. No network
or DGN writes occurred in that check. It also repeated the native form/unit/GCS/
boundary test with unchanged IDs/settings. Reports: `native-standalone-form.json`
and `native-standalone-launcher.json`. Explicit `VBA RUN` via CadInputQueue now
worked; do not infer native execution from a menu action alone. The temporary
test module was removed, actual compile state was up to date and targeted Save
returned Saved=True. Only five production components remain.

Main was reopened through the command queue. **Live metadata Search passed.**
The user selected their eight-vertex test shape, entered source CRS `6576`, kept
us-ft and explicitly confirmed their verified grid mapping. Products appeared.
Read-back of the actual request, status and manifest confirmed `search_complete`,
dataset dem1m, **2 products**, both confirmation flags true, no vertical declaration,
and empty stdout/stderr logs. Evidence: `live-standalone-search.json`; the actual
job remains beside test.dgn under USGS_Terrain. No Download/Process was requested.
No tile transfer, raster processing or native terrain import is claimed by this gate.

Read-only COM initially found both the `.runs` and delivered project loaded with
the same name. The user unloaded the `.runs` row; COM verified only the delivered
copy and four pre-existing projects remain. Native-test/build evidence stays in
ignored `.runs/ord-20261017/`. Rectangle completion/reset and existing-shape acceptance
are now user-confirmed; `graphical-rectangle.json` records those reports.
No speculative preview change was made: Bentley's installed primitive sample
also calls StartDynamics after the first DataPoint.

Further read-only installed metadata inspection found actual managed GCS signatures
in root `Bentley.GeoCoord2.dll` and `Bentley.DgnGeoCoord2.dll`:
`DgnGCS.FromModel(DgnModelRef, Boolean primaryCoordSys)`,
`BaseGCS.GetWellKnownText(out String, WellKnownTextFlavor, Boolean originalIfPresent)`,
`ToJson(out String)`, `GetEPSGCode(Boolean dontSearch)`, LocalTransformType and
LocalTransformParameters. DgnGCS also exposes PaperScale and UOR/cartesian conversion.
The managed platform API is actually in root `ustation.dll`, with .NET Framework 4.8
runtime dependencies; the initially assumed MstnPlatformNET.dll path was absent.
`tests/OrdGcsProbe.cs` is a standalone read-only diagnostic, not part of the MVBA.
It **compiled successfully** with the installed Framework64 csc and Bentley refs;
actual compilation distinguished C# `out` from reflection's generic ByRef display.
Both explicit full-path MDL load attempts (default and explicit task ID) failed with
`MDL Loader: Could not load application ...OrdGcsProbe.dll`; no TEMP report exists.
Thus WKT/JSON extraction has **not executed**. Evidence: `gcs-probe-load.json`.
No DLL was placed in Program Files, no settings/search paths were changed, and no
autoload was configured. The working MVBA remains saved with the hash above.
Full extraction and modifier/control validation remain pending. The user subsequently
explicitly requested the standalone launcher, authorizing the change recorded above.
No policy, ProjectWise/workspace/autoload or search-path setting was changed.

The user clarified the intended native integration: draw/select in **2D**, then
create or reuse ORD's associated Civil 3D model and add the terrain there. A 3D
file is not a selection prerequisite. Public installed Civil SDK terrain-creation
and importer signatures were inspected; their 2D/3D/GeoTIFF/rollback behavior has
not been executed. See the new API_NOTES integration section. No 3D model or
native terrain was created. Do not substitute an arbitrary 3D model or a generic
point triangulation that bridges NoData for the required Civil workflow.

Earlier automation-only ORD project creation/opening failed. The ignored `container-template.mvba` is an
unchanged copy of PatternLineGenerator used in a failed container-opening attempt;
it is not a terrain downloader build. Existing tools were not modified.
The normal ORD host loaded an existing SquareYard project, which was left alone.
Opening the disposable DGN through COM stalled; the user reported a crash/reopen.
The pending call was cancelled. Do not repeat automatic restarts. The user then
opened the disposable 2D test DGN described below in the intended workspace.
The Computer Use native pipe
is unavailable; the user chose API/COM checks only. Graphical gates remain pending.

The user then opened `test.dgn` normally in the intended managed workspace.
COM confirmed model `Default`, type 0, 10000 UOR/master and raw GCS Name
`TN83/2011F`. No EPSG/unit/modifier mapping was inferred. Its four existing VBA
projects are ViewSet, SnappableToggle, Survey and WorkSpaceVersion; leave them
untouched. Direct CREATE/Open attempts still failed; manual blank-project
creation resolved that step and COM imported the source successfully. Check
actual project/file state before resuming. The user explicitly requires **no changes to ANY
ProjectWise or workspace settings**, and plans explicit `VBA LOAD` / `VBA RUN`
ribbon loading. Do not use SAVEALL or edit the existing projects.

The detailed [Windows preflight record](TEST_PLAN.md#windows-ord-computer-preflight-2026-10-17)
separates these results, failures and limits. The older implementation snapshot
below remains useful history, not additional ORD acceptance evidence.

## Original implementation snapshot (before the ORD continuation)

The independent `USGSTerrainDownloader/` source application is implemented.
Existing VBA tools have not been changed. The repository README adds a link.
There is **no compiled USGSTerrainDownloader.mvba**; only Bentley's VBA environment
may produce that file after actual compilation.

- VBA: modeless programmatic UserForm, two-corner temporary rectangle dynamics,
  graphical simple closed-shape selection, active-model guards and Python bridge.
- CRS: optional Bentley GCS **name** probe only. Explicit projected EPSG/WKT,
  master-unit type and grid-coordinate confirmation remain mandatory. Automatic
  extraction of a complete Bentley GCS/modifier definition is not implemented.
- Python: official TNM searches, selection manifest, bounded HTTPS downloads,
  original retention/SHA-256 receipts, compatible DEM mosaicking, target-grid
  polygon clipping, horizontal reprojection, separate elevation-unit conversion,
  read-back verification and processing reports.
- LiDAR: search/download; optional backend PDAL metadata/schema inspection.
  No bare-earth rasterization or classification is implemented.
- Fence, automatic raster attachment, native Civil terrain creation, vertical
  datum shifts, local/ground transforms and dynamic epochs remain deferred.

## What actually passed

The latest Mac GIS run passed **32 offline tests**, with no skips. Real USGS
search/download/processing also passed using two adjacent 1-metre coastal DEM
files totaling **35,755,526 bytes**, under a 50 MiB cap. The worker produced
122-by-122-cell outputs in metres and US survey feet, in horizontal EPSG:6507,
with **14,884 valid cells and zero interior coverage gaps**. Source hashes,
reopened CRS/unit tags and the complete-array elevation conversion were checked.

A live-test defect was fixed: a source-grid cutline before reprojection caused
artificial edge holes. The processor now resamples first, then masks target cell
centers using the original polygon. A constant-surface regression reproduced the
old failure and passes with the fix. Genuine interior NoData still fails unless
explicitly accepted; the successful live run kept `allow_gaps=false`.

Earlier Windows launcher checks passed for spaces/Unicode/ampersands, missing
Python and worker failure. The Windows test PC had **no ORD or GDAL**. Its earlier
31-test checkpoint was 15 passes and 16 skips; this is not Windows GIS/ORD evidence.

No ORD-dependent test, native terrain import or survey-control validation has run.
The live pilot covers one two-tile 1-metre collection, not every USGS data product.

## Get the files onto the ORD computer

In an existing clean checkout:

```powershell
git pull --ff-only origin main
```

For a new checkout:

```powershell
git clone https://github.com/colewinstead/OpenRoads_vba.git
cd OpenRoads_vba
```

Local downloaded DEMs, generated rasters, caches and `.DS_Store` files were not
committed. The Mac's ignored `.runs/live-dem-20261017T060341Z/` retains the original
and diagnostic artifacts. [LIVE_TEST_RESULTS.md](LIVE_TEST_RESULTS.md) records the
inputs, product IDs, hashes and repeatable worker commands. Re-run the bounded
pilot on the ORD computer when a test raster is needed; do not assume ignored Mac
artifacts arrived through Git. Configure that computer's own GIS Python paths.

In this conversation, separate `cd USGSTerrainDownloader` and `ls` shell requests
still listed the repository root. Run the directory change and work together:

```powershell
cd USGSTerrainDownloader; Get-ChildItem
```

Or have the next agent explicitly set its working directory to that folder.

## Next required work — ORD 2026 validation

1. Installed ORD 26.00.00.154 CHM/type-library signature inspection is complete.
   Use F2/F1 Object Browser for any compiler discrepancy. The optional
   `GetGCS(True)`/`Name` form probe still needs an assigned-GCS runtime test;
   failure must retain the explicit-CRS fallback. Do not invent APIs to force compilation.
2. In a **disposable design DGN**, create a VBA project named
   `USGSTerrainDownloader`. Import the three standard `.bas` modules and the
   `clsTerrainBoundaryPicker.cls` class from `vba/`.
3. Insert a blank UserForm named **frmUSGSTerrainDownloader**. Paste all of
   `vba/frmUSGSTerrainDownloader.code.vba` into its code window. It is source code,
   not an importable `.frm`. Follow README's reference/class-header instructions.
4. Use **Debug > Compile USGSTerrainDownloader**, record the actual result and
   full ORD build, then save a real `.mvba` through Bentley's editor.
5. Test the load/run key-ins below. Validate form layout, rectangle preview,
   shape selection, Reset/Cancel, active-model changes and worker progress.
   Confirm DGN element counts/IDs, active settings and workspace configuration
   stay unchanged.
6. Establish known-point agreement for source coordinates, model master units,
   global-origin behavior, datum realization and GCS modifiers. Do not confirm a
   local/ground-scaled system as ordinary projected grid coordinates.
7. Configure a coherent Windows QGIS/OSGeo4W Python/GDAL/PROJ environment and run
   the offline suite with no GIS skips. Run a capped real download/process pilot.
8. Manually attach/import the audited raster in the disposable DGN using the
   installed ORD build's supported workflow. Verify XY, Z units/datum, NoData and
   sample elevations. Record actual results in TEST_PLAN.md. Manual success does
   not establish automatic native terrain creation.

Example key-ins after building the project, adjusting the saved path:

```text
VBA LOAD "C:\CAD_Custom\VBA\USGSTerrainDownloader.mvba"
VBA RUN [USGSTerrainDownloader]USGSTerrainDownloaderMain.Main
```

## Prompt for the next session

> Continue USGSTerrainDownloader in this repository on the computer with ORD 2026.
> Read USGSTerrainDownloader/HANDOFF.md, README.md, API_NOTES.md, TEST_PLAN.md and
> LIVE_TEST_RESULTS.md, plus the saved conversation as needed. The Python worker
> passed 32 offline tests and a real two-tile download/processing pilot. No ORD
> compile, graphical picking, complete GCS extraction or terrain import is proven.
> Prioritize verifying actual Bentley API signatures, assembling/compiling the
> UserForm project and running the disposable-DGN acceptance gates. Preserve all
> CRS/unit/vertical safeguards and existing tools. Do not modify production DGNs,
> workspace/autoload settings or claim an ORD test passed unless it actually ran.
