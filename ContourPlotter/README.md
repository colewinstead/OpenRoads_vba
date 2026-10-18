# Contour Plotter for OpenRoads Designer 2026

A direct TIN contour engine and a native MicroStation VBA writer. Python uses the
original LandXML triangles. VBA inserts connected line strings and ordinary
editable text into the captured active design model on existing DGN levels.
Contours have physical gaps under labels. No raster, DXF, terrain merging or
native terrain modification is involved.

**Status: Windows engine built and ContourPlotter.mvba compiled in ORD
26.00.00.154. Native write/plot acceptance is still required.** Read [test results](TEST_RESULTS.md)
before using a production DGN.

## Run the Windows engine

The locally built executable is dist/ContourPlotter/ContourPlotter.exe. Distribute
the **whole** dist/ContourPlotter folder or its ZIP, including Source and
ThirdPartyLicenses. Coworkers do not need Python, QGIS or GDAL.

From this directory, this complete synthetic example creates native-writer data:

~~~powershell
.\dist\ContourPlotter\ContourPlotter.exe --input .\tests\fixtures\openroads_minimal.xml --output-dir .\jobs\example --axis-order xy --vertical-unit us-ft --model-unit us-ft --crs "Synthetic local grid" --vertical-datum "Synthetic datum" --minor 0.5 --major 1 --label-mode none
~~~

Inspect a source first:

~~~powershell
.\dist\ContourPlotter\ContourPlotter.exe --inspect --input "C:\CAD_Data\MergedTerrain.xml" --output-dir .\jobs\inspect
~~~

Use a dedicated output directory, separate from the source file's directory.
For real data, supply its verified CRS/local grid, datum and axis mapping.
Source horizontal units are read from LandXML; an absent vertical unit must be
supplied explicitly. A supplied unit cannot contradict a declaration. Multiple
surfaces require --surface-index (zero based) or an unambiguous --surface-name.
xy means Easting/Northing/Z; ne means Northing/Easting/Z.

## Load the compiled VBA in ORD

The configured project is vba/ContourPlotter.mvba. The Windows distribution
includes a copy beside ContourPlotter.exe. Load that project, then run
`vba run [ContourPlotter]ContourPlotterMain.Main`. Keep the whole engine folder
together and choose its ContourPlotter.exe in the dialog.

To rebuild the MVBA from source:

1. Open a disposable test DGN design model in ORD 2026. Open VBA Project Manager,
   create a **new** project, save it as ContourPlotter.mvba, and set its VBA project
   name to ContourPlotter.
2. In its Visual Basic Editor import all six vba/*.bas modules. Create one blank
   UserForm and name it **frmContourPlotter**. Paste the contents of
   vba/frmContourPlotter.code.vba into its code pane. Controls are built at runtime.
   Do not import the form code as a standard module.
3. Keep the Bentley MicroStation DGN library and Microsoft Forms references supplied
   by ORD. Other Windows automation is late bound. Choose **Debug > Compile**,
   correct any host-specific errors, and save the project. Record the actual build
   and result in [ORD_TEST_PLAN.md](ORD_TEST_PLAN.md).
4. Use the key-ins below, replacing the path with the MVBA you actually saved.

~~~text
vba load "C:\CAD_Custom\ContourPlotter\ContourPlotter.mvba"
vba run [ContourPlotter]ContourPlotterMain.Main
~~~

The dialog has three tabs:

- **Source and coordinates:** browse to the packaged EXE and manually exported,
  already merged LandXML; Inspect; select the named surface; resolve stored axes,
  horizontal/vertical units, source grid and vertical datum. Browse uses the native
  Windows file picker; pasted Windows Copy-as-path values may include surrounding
  quotes. The current DGN GCS
  is displayed as a reference, including custom MDOT systems.
- **Contours and labels:** minor/major intervals and base use source elevation
  units. Major must be an integer multiple of minor. Label spacing uses model
  master units. Plot scale is **model master units per paper inch**: 20, 50, 100
  ft/in, or a custom value; metric 1:500 is 12.7 m/in. Paper height 0.10 at
  50 ft/in gives a measured five-foot physical text height. Major and minor labels
  have separate paper-height fields; both start at 0.10 in. Each uses its own
  measured text footprint for gaps and collision checks.
- **MDOT levels and output:** select existing levels and text styles from the
  DGN or its configured DGNLIBs, including unused definitions. There are no
  guessed MDOT names or replacement definitions. ORD copies a selected library
  level into the DGN when an element first uses it. A selected library text style
  is copied with its original name into the DGN before measuring labels, allowing
  element-local size overrides; source libraries stay unchanged.
  ByLevel is the default; optional color, weight and line-style overrides require
  turning it off. Text font and proportions come from the selected style; local
  size/justification overrides do not modify its shared definition.

**Analyze / Preview** calculates output and reports counts without inserting
graphics. Measuring a library text style caches that selected definition in the
DGN; it does not modify the DGNLIB or its style.
**Generate** reanalyzes the source, measures its exact label strings in ORD,
calculates gaps/collisions and validates the entire locked geometry stream before
the first write. Confirm the DGN, model, grid, axis, datum, units and bounds.
2D writes Z=0 even in a 3D model; 3D converts source elevations to model units.
Sheet/drawing, locked, read-only and reference targets are rejected.

The worker runs without creating a console window. The dialog remains responsive
to cancellation; failed jobs retain readable error.txt and detailed error.json.

Nothing edits active levels, global active symbology, terrain display, workspace
configuration or VBA autoload settings. If you wish to turn off live terrain
contours, use the approved ORD display workflow separately.

## Regenerate and recover

Every insertion has a UUID, per-element XData and a Unicode elements.cpj journal.
Request/settings, source SHA-256, units, surface and timestamp are retained beside
it in the run folder under %LOCALAPPDATA%\ContourPlotter\jobs.

**Regenerate** creates the entire replacement first. It then validates the prior
journal against the current file/model and checks each recorded ID and run tag.
A separate confirmation covers deletion of previous members, **including manual
edits**. Copies with new IDs and unrelated elements on the same levels are
preserved. Choose No to retain both runs for review.

A failed/cancelled new insertion attempts to remove only its own added elements.
There is no claim of a fully atomic VBA transaction or one-step Undo. Never discard
run folders until the DGN has been saved, reopened and checked. If an application
crash interrupts insertion, preserve the partial journal and result files;
restart the tool against the same DGN/model. The read-only
key-in below inventories tags for manual reconciliation of a partially flushed
batch. Do not delete by level.

~~~text
vba run [ContourPlotter]modContourElements.ListTaggedRuns
~~~

Removal failure during the staged cleanup retains the complete new run and reports
how many old members were removed. A committed old journal can be reviewed and
used again; missing/altered IDs must be reconciled before retrying. Annotation
scale changes after generation, save/reopen tag persistence, manual edits, copying
and plotting all require the ORD acceptance checks.

## Develop and build

Use 64-bit Python 3.14 on Windows (tested). In an approved Python environment:

~~~powershell
python -m pip install -r requirements.txt pyinstaller==6.17.0
$env:PYTHONPATH = Join-Path (Get-Location) "python"
python -m unittest discover -s tests -v
python .\tests\benchmark.py --output .\examples\benchmark.json
python .\packaging\build_windows.py
~~~

The optional PowerShell wrapper is packaging/build_windows.ps1. On systems that
disable scripts, use the Python builder above; no execution-policy changes are
needed. A private dependency directory can be passed to the builder with
--dependency-path. The builder runs tests and both version and actual contour
smoke tests of the packaged EXE, bundles reviewable source/notices and creates
dist/ContourPlotter-Windows-x64.zip. It never includes private LandXML fixtures.

If endpoint security blocks unsigned executables, ask CAD/IT to approve the package
or sign it using the organization's certificate and release process. No bypass is
part of this tool.

## Limits and licensing

No automatic terrain selection/export, optional polygon clipping, smoothing,
reprojection or vertical datum transformation is implemented. The supplied
visible triangles define coverage and holes. Very short fragments are reported,
not removed. Label collisions cover other generated labels; existing DGN text,
roadway linework and labels across separate runs are not obstacles.

The Python workflow holds the parsed XML and calculated contours in memory.
VBA reads bounded records in two passes and retains new element handles for
rollback. Native library calls/XML parsing are not instantly interruptible.
The 10,000-elevation guard and 1,024-vertex element cap are implementation limits,
not agency criteria. No real DGN insertion/save/plot speed is claimed.

The Python engine is **GPL-2.0-or-later**; its independent VBA file client is
**MIT**. Reused parser code retains Edmond Akello's attribution. See
[third-party notices](THIRD_PARTY_NOTICES.md), [architecture](ARCHITECTURE.md),
[format](FORMAT.md) and [API evidence](API_NOTES.md).

For the supplied Merged.xml export, the live terrain bounds establish stored
**N E Z (ne)** order. The previous E/N interpretation transposed the plotted
location. Select axes from source/control evidence for each export; plot scale
only changes label sizes and does not reposition contour coordinates.
