# OpenRoads_vba

Bentley OpenRoads Designer / MicroStation VBA macros. The original
`nodeincell.mvba` and `PerpendicularLineGenerator/PatternLineGenerator.mvba`
ship compiled. The tools also include reviewable **VBA source** where available
(`.bas`, with class/form source where a tool needs them).

**Want it packaged?** [**OpenRoads Field Kit**](https://fingazpdx.gumroad.com/l/openroads-field-kit) — the same macros as a one-click zip with a quick-start guide, printable cheat sheet, and email support. The original macros are MIT and stay free; ContourPlotter's engine is GPL-2.0-or-later and its independent VBA client is MIT; the kit is packaging and a reply when something breaks.

The macros target plain MicroStation VBA so they run in any OpenRoads
Designer release without depending on a specific CivilModel API version.
A few (StationOffset, PointsIO) are written as conservative fallbacks for
the OpenRoads-native annotation tools.

## Importing

1. Clone or download the repo.
2. In OpenRoads Designer: open the **VBA Project Manager**.
3. Right-click an existing project (or create a new one) → **File →
   Import File...** → pick a `.bas` file.
4. Each module appears under your project. Run a sub with:

   ```
   VBA RUN [ProjectName] ModuleName.SubName
   ```

   For example: `VBA RUN MyTools LabelAcres.LabelAcresMain`

5. Bind frequent commands to function keys via **Workspace → Function
   Keys** for one-keystroke access.

## Adding a VBA macro to a ribbon button (ORD / ProjectWise)

You can add a custom ribbon button that **loads a VBA project when clicked and then runs its macro**. This is useful when switching between ProjectWise managed workspaces, where VBA autoload configuration may not persist or may interfere with workspace compilation.

1. Save your compiled `.mvba` file in an accessible, local or synced folder. If using OneDrive, make the file **Always keep on this device**.
2. In ORD's **Key-in** window, test the VBA load and run commands separately:

   ```text
   vba load "C:\CAD_Custom\VBA\RotateElementHorizontal.mvba"
   vba run [RotateElementHorizontal]Main
   ```

3. Test the combined key-in below on a disposable element in a test DGN. Replace the example path, project name, and procedure with those for your macro:

   ```text
   vba load "C:\CAD_Custom\VBA\RotateElementHorizontal.mvba";vba run [RotateElementHorizontal]Main
   ```

   The semicolon separates the two key-ins; depending on your ORD settings, you may need **Parse All Mode** enabled. The quoted path accommodates spaces and punctuation in folder names.

4. Right-click the ORD ribbon and choose **Customize Ribbon**. In a personal ribbon customization, add a group and button (or edit an existing button). Set the button's **Key-in** to the tested combined command, give it a label and icon, and click **Apply**.
5. Restart ORD and test the button in another DGN or ProjectWise project. Back up your personal ribbon `.dgnlib` independently of your `.upf` preferences file.

### ProjectWise managed-workspace warning

**Avoid adding personal, absolute VBA paths to `MS_VBAAUTOLOADPROJECTS` or `MS_VBASEARCHDIRECTORIES` in ProjectWise Personal WorkSpace settings without CAD administration approval.** In a managed workspace, such settings can cause **Error 222: Failed to generate managed workspace configuration file** during workspace compilation. The button-based load-and-run method above does not require these configuration entries.

If this error occurs after adding personal VBA paths:

- Use **ProjectWise Explorer → Tools → Options → Personal WorkSpace** to remove **only the custom entries you added**, leaving the organization's existing configuration intact.
- Do not edit automatically generated ProjectWise workspace `.cfg` files (which can be regenerated), and do not delete `.upf` or ribbon `.dgnlib` files to troubleshoot this issue.
- If you cannot access or safely edit the Personal WorkSpace settings, contact your ProjectWise/CAD administrator.

> **Note:** `RotateElementHorizontal.mvba` is an example project from an external workspace; it is **not included in this repository**. Use a VBA project you have permission to run and distribute.

## Routines

### Parcel / area work

| Sub                          | File                       | What it does                                                                 |
|------------------------------|----------------------------|------------------------------------------------------------------------------|
| `LabelAcresMain`             | `LabelAcres.bas`           | Label every selected closed shape with **area in acres to 0.1 AC** at the bbox center. |
| `TotalAreaMain`              | `TotalArea.bas`            | Sum area of every selected closed shape; report sq ft + acres.                |
| `TotalLengthMain`            | `TotalLength.bas`          | Sum length of every selected linear element; report ft + mi.                  |
| `LotBDMain`                  | `LotBearingDistance.bas`   | Label every straight segment of the first selected shape with bearing+distance, plus area at centroid. |

### Survey / point I/O

| Sub                          | File                       | What it does                                                                 |
|------------------------------|----------------------------|------------------------------------------------------------------------------|
| `ImportPNEZDMain`            | `PointsIO.bas`             | Read a PNEZD CSV and place zero-length lines + 3-line text nodes for each point. |
| `ExportPNEZDMain`            | `PointsIO.bas`             | Write every selected text node out to PNEZD CSV (first line treated as PNO).  |
| `CoordNEMain`                | `CoordLabel.bas`           | Two-line Northing / Easting label at the first selected element's centroid.    |
| `CoordElevMain`              | `CoordLabel.bas`           | One-line `EL ###.##` label at the first selected element's centroid.           |
| `PointGridMain`              | `PointGrid.bas`            | Rectangular grid of zero-length point elements between two corners.            |
| `StationOffsetMain`          | `StationOffset.bas`        | Station & offset (L/R) of a target X,Y relative to a selected alignment polyline. |

### COGO

| Sub                          | File                       | What it does                                                                 |
|------------------------------|----------------------------|------------------------------------------------------------------------------|
| `InverseReportMain`          | `InverseReport.bas`        | Full inverse on the first selected line: bearing, distance, dN/dE/dZ, slope.  |
| `SlopeLabelMain`             | `SlopeLabel.bas`           | Slope label at the midpoint of a selected line (percent + H:V ratio).         |
| `Main`                      | [PerpendicularLineGenerator](PerpendicularLineGenerator/README.md) | ORD 2026 cross-lines at cumulative geometry distances; settings UserForm, forward/reverse, centered/left/right, active level/symbology. |

#### Perpendicular Line Generator (ORD 2026)

Load the included `PerpendicularLineGenerator/PatternLineGenerator.mvba`, a real
project saved in ORD with its UserForm. The [installation and usage guide](PerpendicularLineGenerator/README.md)
also includes complete module/class source and instructions for rebuilding the
UserForm from `frmPerpendicularLines.code.vba`. A `.bas` import alone does not
create this dialog.

```text
VBA LOAD "C:\CAD_Custom\VBA\PatternLineGenerator.mvba"
VBA RUN [PatternLineGenerator]PerpendicularLineGeneratorMain.Main
```

Replace the path with your saved file location. For a ribbon button that loads
and runs the tool, use:

```text
vba load "C:\CAD_Custom\VBA\PatternLineGenerator.mvba";vba run [PatternLineGenerator]PerpendicularLineGeneratorMain.Main
```

Run the tool, select a horizontal alignment, then pick the start and end with
dynamic cross-line previews before the settings dialog opens. Pick Start/Pick End
buttons re-pick limits; Full Alignment restores the complete range. Reset during
a pick opens settings without accepting the moving cursor location.
Defaults are 150 master-unit spacing and 200 master-unit total length, centered,
start 0, geometric end, starting line included, ending line excluded, forward.
The dialog shows active level, working units, selection status and estimated
count. In a feet-master model, a 1,000-foot straight produces seven 200-foot
lines at distances 0 through 900 feet; each extends 100 feet on either side.

Supported exposed VBA geometry includes lines, line strings, arcs, B-splines
and connected complex strings. Inputs are geometry distances, **not ORD
stationing**. Native civil spirals and station equations are not interpreted;
displayed spline/stroke geometry does not certify native alignment accuracy.
See the [API/support notes](PerpendicularLineGenerator/API_NOTES.md),
[troubleshooting](PerpendicularLineGenerator/README.md#troubleshooting), and
[repeatable ORD test plan](PerpendicularLineGenerator/TEST_PLAN.md). This tool
uses actual model master units, including non-foot units. The user reports the
generator and dynamic picking working in ORD; measured acceptance tests remain
pending. It changes no ProjectWise or VBA autoload configuration.

### USGS Terrain Downloader (test build)

[USGS Terrain Downloader](USGSTerrainDownloader/README.md) adds an independent
ORD 2026 VBA dialog for temporary rectangle/simple-shape boundaries, official USGS
DEM/LiDAR searches and a Python GDAL/PROJ terrain-raster workflow. CRS/master units
and source vertical datum/units require explicit verification. The included test
MVBA compiled in ORD 26.00.00.154. Native 2D form/GCS/shape/boundary and standalone
worker-bridge checks passed, along with 42 Windows tests. Basic picking was
user-confirmed; three Windows DEM jobs completed. After Float64 import failures,
the user confirmed a Float32/finite-NoData copy imported; the worker now delivers
that storage format. Known control and complete native terrain validation have
separate acceptance gates. This tool has its own explicit
unit contract; the legacy feet assumptions below do not apply to it.

### Contour Plotter (prototype)

[Contour Plotter](ContourPlotter/README.md) adds a direct LandXML TIN contour engine,
a standalone Windows package and native DGN line/text writer source. It selects
existing MDOT levels and text styles, creates physical label gaps, and journals
exact output IDs for staged regeneration. Python tests and live unsaved API probes
passed. The included MVBA compiled in ORD 26.00.00.154 and completed the form-to-engine
preview using a real LandXML surface; native writes and plotting still need ORD acceptance.
The standalone engine is GPL-2.0-or-later because it reuses the attributed
LandXML parser; the independent VBA file client remains MIT.

### Text utilities

| Sub                          | File                       | What it does                                                                 |
|------------------------------|----------------------------|------------------------------------------------------------------------------|
| `TextUpperMain`              | `TextCase.bas`             | Selected text/text-nodes → UPPER.                                             |
| `TextLowerMain`              | `TextCase.bas`             | Selected text → lower.                                                        |
| `TextTitleCaseMain`          | `TextCase.bas`             | Selected text → Title Case.                                                   |
| `TextHeightMain`             | `TextHeight.bas`           | Apply a new height (master units) to every selected text/node.                |
| `TextRotateHorizMain`        | `TextRotateHoriz.bas`      | Rotate upside-down text 180° so it reads left-to-right.                       |

### Z / elevation

| Sub                          | File                       | What it does                                                                 |
|------------------------------|----------------------------|------------------------------------------------------------------------------|
| `FlattenZMain`               | `ZTools.bas`               | Set Z = 0 on every vertex/origin of every selected element.                   |
| `MoveZMain`                  | `ZTools.bas`               | Add a delta to Z across the selection.                                        |
| `SetZMain`                   | `ZTools.bas`               | Set Z to a single value across the selection.                                 |

### Levels & cells

| Sub                          | File                       | What it does                                                                 |
|------------------------------|----------------------------|------------------------------------------------------------------------------|
| `LevelCountMain`             | `LevelCount.bas`           | Element counts grouped by level (graphical elements only).                    |
| `LevelTranslateMain`         | `LevelTranslate.bas`       | CSV-driven level rename / re-leveler (`OLD,NEW` per row, auto-creates).       |
| `LevelIsolateMain`           | `LevelIsolate.bas`         | Turn display off on every level except the selected element's level.          |
| `LevelShowAllMain`           | `LevelIsolate.bas`         | Turn display on for every level.                                              |
| `CellCountMain`              | `CellCount.bas`            | Tally of placed cells grouped by cell name.                                   |

### Original

| File                | What it does                                                                       |
|---------------------|------------------------------------------------------------------------------------|
| `nodeincell.mvba`   | Copy + increment text inside a cell as a single operation (compiled VBA project).   |

## Notes

- **Working units** are assumed to be survey **feet**. For metric DGNs,
  change `SQFT_PER_ACRE` in `LabelAcres.bas`, `TotalArea.bas`, and
  `LotBearingDistance.bas` to `4046.856` (sq m / acre).
- **`StationOffset`** is a fallback for environments where the OpenRoads
  CivilModel API isn't loaded. For real alignment objects with stationing
  equations and superelevation, prefer the native OpenRoads
  **Civil Annotation** tools — this VBA only understands plain MicroStation
  geometry.
- **`LevelIsolate`** writes to `Level.IsDisplayed`. If your DGN is
  attached to references with their own level override masks, the
  references won't update — toggle Level Display on the reference itself.
- **`ImportPNEZD` / `ExportPNEZD`** treat the second line of a text node
  as the elevation when round-tripping. The import-then-export pair is
  lossless within `Format("0.000")` precision.
- Source modules use `Option Explicit` and `On Error GoTo Cleanup` so
  errors surface with a clear message instead of silently corrupting
  geometry.

## Sister repos

Vanilla AutoCAD / Civil 3D LISP equivalents of many of these macros
(LABELACRES, BD, SLP, NE, ZL, CHZ, FLAT, etc.) live in
[C3D-AutoCAD](https://github.com/mf4633/C3D-AutoCAD). Carlson Civil/Survey
equivalents (LABELAC, LOTBD, PNORENUM, PNOINV, TRAV, FBREP, STALBL, etc.)
live in [Carlson-CAD](https://github.com/mf4633/Carlson-CAD). Together the
three repos cover the LISP / VBA tooling layer of most US civil offices.

## Contributing

One `.bas` per logical unit (one module, one or more closely related Subs).
Header comment names the Sub(s), the key-in syntax, and any assumptions
about working units. Append a row to the right table above.

## See also

Companion calculators for corridor design: [AASHTO superelevation](https://pe-calc.com/tools/superelevation.html) and [vertical curve](https://pe-calc.com/tools/vertical-curve.html) on pe-calc.com.
