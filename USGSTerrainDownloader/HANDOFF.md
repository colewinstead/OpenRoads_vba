# Continue on the ORD computer

Saved 2026-10-17. Repository: https://github.com/colewinstead/OpenRoads_vba

Start here, then read [README.md](README.md), [API_NOTES.md](API_NOTES.md),
[TEST_PLAN.md](TEST_PLAN.md), and [LIVE_TEST_RESULTS.md](LIVE_TEST_RESULTS.md).
The [saved conversation](codex-session-01a1485e-079a-78b3-8e11-c54721d64042.md)
contains the original task, implementation activity, and the latest continuation.

## Current state

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

1. Read the installed ORD 2026 `MicroStationVBA.chm` and F2/F1 Object Browser.
   Verify new application signatures on this actual build. Treat the optional
   `GetGCS(True)`/`Name` probe as unverified until exercised; failure must retain
   the explicit-CRS fallback. Do not invent APIs to force compilation.
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
