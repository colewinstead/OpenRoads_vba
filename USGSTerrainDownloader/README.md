# USGS Terrain Downloader

OpenRoads Designer 2026 VBA application with source and a separate Python GIS
worker. Draw a temporary model-XY rectangle or locate an existing simple closed
shape, search USGS, choose products, and produce an audited GeoTIFF. Existing VBA
tools are independent and unchanged.

**Status:** the Python DEM workflow passes 32 offline tests on both Mac GIS and
Windows OSGeo4W runtimes, plus a bounded
[real USGS two-tile download/processing pilot](LIVE_TEST_RESULTS.md). The VBA project
has now been assembled, compiled and saved in ORD 26.00.00.154. The included
`USGSTerrainDownloader.mvba` is a **test build**. A native 2D self-check passed
form initialization, explicit-confirmation defaults, GCS name display, in-memory
shape vertices, programmatic boundary acceptance and unchanged DGN IDs/settings.
The Windows suite now passes **37 tests**, including five standalone-launcher
tests. An actual ORD bridge test started the EXE and Python and read Python's
expected invalid-schema response. A live search through the ORD dialog then
completed and displayed **2 one-metre DEM candidates** for the user's test shape.
The user confirmed normal launch, rectangle
preview/capture/cancellation and existing-shape acceptance. Numeric coordinate
agreement, 3D behavior and terrain import have separate acceptance gates.
Use the [acceptance plan](TEST_PLAN.md)
before production use. This is not survey-grade data or an automatically created
native ORD Terrain Model.

Installed ORD 26.00.00.154 API signatures have now been inspected. The application
now uses the included standalone `python\launch_worker.exe`; it does not require
PowerShell scripts. The earlier PowerShell launch failure remains recorded as
history. Live download/processing/import acceptance remains outstanding. See the
[Windows preflight evidence](TEST_PLAN.md#windows-ord-computer-preflight-2026-10-17).

Continuing on another computer? Read [HANDOFF.md](HANDOFF.md) and the
[saved conversation](codex-session-01a1485e-079a-78b3-8e11-c54721d64042.md).

## Included behavior

- Modeless, programmatically assembled UserForm; two-point rectangle dynamics
  and active-model simple straight-sided shape selection. Preview elements exist
  only in memory. No rectangle is written, and no active settings are changed.
- Best-effort Bentley GCS **name** probe, model working-unit label, mandatory
  user-confirmed projected EPSG/WKT, master-unit type and grid-coordinate mapping.
  No Bentley-name-to-EPSG guessing. Local/ground transforms are unsupported.
- PROJ horizontal transformations with conventional XY order, no ballpark fallback,
  best-operation requirement and network grids disabled. Densified boundary search;
  Mississippi East/West NAD83 and NAD83(2011), metre and survey-foot EPSGs tested.
  International feet and metres remain distinct from US survey feet.
- Live TNM catalog validation; paginated 1-m DEM, 1/3-arc-second DEM and LAS/LAZ
  searches; product metadata/details, estimates and explicit product selection.
- Bounded HTTPS streaming, temporary downloads, local SHA-256 receipts, safe reuse
  of verified completed originals, cooperative cancellation, diagnostic/status files.
- Compatible DEM mosaicking, optional target-grid polygon clipping after horizontal reprojection to
  the confirmed DGN CRS, **separate numeric elevation-unit conversion**, GeoTIFF 1.1
  compound horizontal/vertical metadata and read-back verification.
- Coverage-gap and overlap gates, source/operation reports, and output-folder opening.
- LiDAR search/download and original preservation. Optional **backend-only** PDAL
  metadata/schema inspection; bare-earth rasterization and classification are deferred.

Fence selection, automatic raster attachment, native Civil terrain creation, vertical
datum shifts, curved/complex boundaries, reference attachments, rotated rasters,
misaligned/mixed-resolution/mixed-CRS mosaics and dynamic coordinate epochs are not
implemented. Unsupported inputs stop with a diagnostic. See [API notes](API_NOTES.md).

## Windows GIS setup (no administrator rights needed by this application)

Use one coherent **64-bit Python + GDAL >= 3.8 + numpy + pyproj >= 3.6** installation.
Python 3.10 or later is recommended. A working QGIS/OSGeo4W environment is suitable;
GDAL's [official distribution page](https://gdal.org/en/stable/download.html) lists
Windows distribution options. Reuse your organization's existing installation or
have a GIS environment supplied in a user-writable location. The application does
not install software, elevate privileges, or change system/workspace settings.
A system-wide GIS installation may separately require your IT administrator.

1. Copy this entire directory to, for example,
   `C:\CAD_Custom\VBA\USGSTerrainDownloader`. Keep its `python` subdirectory intact.
   `python\launch_worker.exe` is included for Windows x64 with .NET Framework 4.8
   (also required by ORD 2026). To rebuild it from `WorkerLauncher.cs`, run this
   from the configured GIS shell, using your application path:

   ```bat
   python "C:\CAD_Custom\VBA\USGSTerrainDownloader\python\build_launcher.py"
   ```

   This uses the installed Windows Framework compiler; no .NET SDK or admin access
   is required. Keep the EXE beside `main.py`.
2. Open the installed **OSGeo4W/QGIS command shell**. Check the actual interpreter:

   ```bat
   where python
   python -c "import sys; from osgeo import gdal; import pyproj,numpy; print(sys.executable); print(gdal.VersionInfo('--version')); print(pyproj.__version__)"
   ```

3. Make a writable configuration folder, then capture that shell's GIS paths:

   ```bat
   mkdir "%LOCALAPPDATA%\USGSTerrainDownloaderConfig"
   python "C:\CAD_Custom\VBA\USGSTerrainDownloader\python\capture_environment.py" "%LOCALAPPDATA%\USGSTerrainDownloaderConfig"
   ```

   This verifies dependencies and creates `runtime-environment.json` and
   `runtime-python.txt`. Do not assume a Python version directory such as
   `Python312` matches your installation. The included environment example is a
   template only. Do not install a second pip GDAL over QGIS's GDAL bindings.
   If pyproj is absent, use the distribution's package manager or a separate,
   coherent user-owned GIS environment; `requirements.txt` intentionally does not
   attempt to build GDAL.
4. Run the offline tests from that same shell:

   ```bat
   python -m unittest discover -s "C:\CAD_Custom\VBA\USGSTerrainDownloader\tests" -v
   ```

5. The dialog discovers **Source code folder** beside the single loaded terrain
   project when its `python\launch_worker.exe` exists. **Python executable** and
   **GIS environment JSON** fill from the files captured in
   `%LOCALAPPDATA%\USGSTerrainDownloaderConfig`. Previously saved choices take
   precedence; missing/ambiguous paths stay editable. If needed, enter the application
   directory, the full Python path from `runtime-python.txt`, and the captured JSON.
   Select `python.exe`, not a
   `.bat`/`.cmd` QGIS launcher or Windows Store alias. These three settings persist
   in this application's HKCU VBA settings. CRS/vertical declarations do not persist.

Model master units fill only when ORD's numeric unit definition matches metres,
international feet or US survey feet; labels alone are insufficient. The CRS/grid
verification checkbox remains unchecked. Source vertical CRS/unit/evidence remain
blank until verified against the selected data; the dataset name is not evidence.

The standalone EXE launches Python in a dedicated child environment. ORD's
environment, ProjectWise/workspace settings, VBA autoload configuration and
PowerShell policy are not modified. The old `launch_worker.ps1` is retained for
existing command-line workflows; the VBA application does not invoke it. TLS certificate
verification stays enabled; configure `SSL_CERT_FILE` in the child environment for
an approved CA bundle when needed. Never disable certificate verification.

## Build the VBA project in ORD 2026

For the included test build, load `USGSTerrainDownloader.mvba` from this directory
and use the key-ins below. The following steps rebuild it from source.

1. Open a **disposable design DGN** (a normal 2D design model is supported), then open **VBA Project Manager** and create a
   new project. Save it as
   `C:\CAD_Custom\VBA\USGSTerrainDownloader.mvba`.
2. In the VBA editor, **Tools > [project] Properties**, set the project name to
   exactly `USGSTerrainDownloader`.
3. **File > Import File** these standard modules:
   `vba\USGSTerrainDownloaderMain.bas`, `vba\modTerrainCoordinates.bas`, and
   `vba\modPythonBridge.bas`.
4. Import `vba\clsTerrainBoundaryPicker.cls`. It must appear under **Class Modules**.
   If it appears under Modules, remove that component, **Insert > Class Module**,
   name it `clsTerrainBoundaryPicker`, and paste from `Option Explicit` through the
   end. Do not paste `VERSION`, `BEGIN/END` or `Attribute` export headers manually.
5. **Insert > UserForm** and set **(Name)** to `frmUSGSTerrainDownloader` using F4.
   Leave the designer empty. Paste the entire contents of
   `vba\frmUSGSTerrainDownloader.code.vba` into its code window. This file is code,
   **not** an importable `.frm`; controls and `WithEvents` handlers are assembled
   at initialization, following the existing Perpendicular Line Generator pattern.
6. In **Tools > References**, check the installed **Bentley MicroStation DGN Object
   Library**, **Visual Basic for Applications**, **OLE Automation**, and
   **Microsoft Forms 2.0 Object Library** (normally added with the UserForm).
   WScript, ADODB, filesystem and Windows Shell objects are late-bound; no manual
   references or programmatic VBA-project access are needed.
7. Use **Debug > Compile USGSTerrainDownloader**. Resolve any missing reference
   against your installed ORD 2026 VBA help/Object Browser; do not substitute
   guessed APIs. Save the project. The included test build passed this step on
   ORD 26.00.00.154; recompilation on another build is a separate check.
8. Run these exact key-ins, adjusting only the load path:

   ```text
   VBA LOAD "C:\CAD_Custom\VBA\USGSTerrainDownloader.mvba"
   VBA RUN [USGSTerrainDownloader]USGSTerrainDownloaderMain.Main
   ```

   After both work separately, a ribbon button can use:

   ```text
   vba load "C:\CAD_Custom\VBA\USGSTerrainDownloader.mvba";vba run [USGSTerrainDownloader]USGSTerrainDownloaderMain.Main
   ```

   No autoload variable changes are necessary. Follow the repository's existing
   ribbon instructions if desired. Only an actual Bentley VBA environment can
   produce the saved `.mvba`.

Boundary selection can start in a **2D model**. The current application does not
create or activate a 3D model. Creating/reusing ORD's associated Civil 3D model
and importing a native terrain remain an integration task, separate from drawing
the boundary and producing the GeoTIFF. No ProjectWise or workspace configuration
changes are required for explicit ribbon load/run.

## First run

1. Set runtime paths as above. The scrollable panel keeps all settings accessible
   on smaller screens; scroll down for the search/process buttons.
2. **Draw Rectangle**: click two opposite corners. The rectangle aligns to model
   XY, even in a rotated view. Reset cancels without accepting the cursor position.
   Alternatively, **Select Existing Shape**, identify and accept a simple closed
   straight-sided horizontal shape in the active design model. Shape vertices,
   not its bounding box, are retained for clipping. References/complex shapes/holes
   are rejected. No boundary is added to the DGN.
3. Inspect the detected Bentley GCS name and master-unit label. Independently
   verify the model GCS, datum realization, global-origin/control-point behavior,
   modifiers and coordinate readout. Enter the equivalent projected EPSG/WKT,
   select `m`, `ft` (international) or `us-ft`, then confirm **grid XY**. A label
   such as `ft` alone is not evidence of the foot definition. No offset, rotation
   or ground-to-grid scale is applied by this version. Do not confirm a local or
   modified coordinate system as ordinary grid coordinates.
4. Select the USGS data source and choose a writable destination. Default is a
   writable `USGS_Terrain` subfolder beside the DGN; read-only/managed locations
   fall back to `%LOCALAPPDATA%\USGSTerrainDownloader`, then `%TEMP%`. Use Browse
   to choose your intended project location.
5. Click **Search USGS**. Review the WGS84 bounding extent, product details and
   metadata links. Only then select products in the list. Publication date is not
   an accuracy ranking. Some geographic bounding-box candidates may fall outside
   an irregular boundary; processing uses the original polygon.
6. For DEM processing, confirm each selected source's actual **vertical** datum
   and unit from source-specific metadata. Enter the source elevation unit,
   vertical CRS (for example `EPSG:5703` only for verified NAVD88), and a metadata
   URL/report citation in **Evidence / metadata**. The form applies this common
   declaration to all selected tiles. Embedded contradictory tags cause failure;
   heterogeneous evidence requires separate runs or per-product backend declarations.
7. Set output elevation units independently of horizontal units. Enter output
   cell size in the **source horizontal CRS's native units**, which can differ
   from the model's displayed master units. A 1-m cell is `1` in a metre CRS;
   represent 1 metre as `3.28083333333333` in a US-survey-foot CRS. Selecting a finer
   grid does not improve source accuracy. Compatible tiles mosaic automatically.
8. Review overlaps, total size and the hard GiB cap (default 5). Select a coherent
   collection. Overlapping coverage fails unless you explicitly accept **last
   selected list row wins**. Gaps inside the polygon fail by default; the optional
   gap checkbox retains/reports NoData. Unknown catalog sizes require explicit
   acceptance and are still bounded by the hard transfer cap.
9. Click **Download / Process**. Inputs disable while a child worker runs; the
   modeless dialog yields to ORD and updates search/download/processing states.
   **Cancel** signals the child; network reads can take up to their 30-second
   timeout, and some GDAL/read-back stages finish before the next cancellation
   check. Closing while busy requests cancellation and keeps the dialog alive.
10. **Open Output Folder** and review `report.json`, `terrain.tif` and
    `ORD_IMPORT.txt`. Compare horizontal location and elevations against known
    project control before using the surface. Attach/import manually in the
    disposable DGN and record the tested ORD build and actual import settings.
    Automatic raster attachment/native terrain creation are not claimed.

## Files and recovery

Every search/process creates a separate `run-*` folder. `request.json` is UTF-8;
`status.json` is machine-readable; `status.txt`, `products.tsv` and `extent.txt`
are UTF-16 for VBA. Writes replace temporary files atomically. Searches save full
TNM product metadata in `search.json`. Processing writes an incremental report
with selected products, verified download sizes, SHA-256 hashes, actual raster
metadata, horizontal pipeline, elevation multiplier, vertical evidence and gaps.

Completed originals and receipts live in `destination\originals`, separately
from processed output, and are always preserved. Interrupted `.part` files are
removed. Retry a failed process in the **same destination** to reuse completed
originals after checksum verification. A corrupt/missing receipt causes a stop;
move the suspect source/receipt aside after review. Concurrent downloads of the
same source are locked. If a worker crashes, remove a stale `.lock` only after
confirming no worker is active. A cancelled/failed raster is never published as
`terrain.tif`. Inspect `error.log`, `stderr.log` and the failure report before retry.

Limits are implementation safeguards, not engineering standards: 500 search
candidates, 2-degree extent span, 10,000 km² area, 20,000 densified vertices,
100-million output cells, configurable download cap up to 100 GiB. Source files
must be direct HTTPS GeoTIFF/LAS/LAZ; archived products require manual handling.
Mosaic tiles must have equal horizontal CRS, pixel spacing, aligned grids and
compatible confirmed vertical datum/units. NoData is preserved; gaps are measured
at pixel centers. No vertical datum/geoid transformation is performed.

## Backend use

The native dialog creates requests automatically. For headless integration, run:

```bat
python python\main.py C:\terrain-work\run-example\request.json
```

A commented schema/example is documented in [REQUEST_FORMAT.md](REQUEST_FORMAT.md).
Use the launcher rather than direct Python when child GIS environment settings
are needed:

```bat
python\launch_worker.exe "C:\terrain-work\run-example\request.json"
```

LiDAR requests can additionally set `pdal_executable` to a full PDAL
CLI path for metadata/schema inspection; this is not a ground-point raster pipeline.

[API evidence and architecture](API_NOTES.md) · [Repeatable tests and results](TEST_PLAN.md) · [Real download/processing evidence](LIVE_TEST_RESULTS.md)
