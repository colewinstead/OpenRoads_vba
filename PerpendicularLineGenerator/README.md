# Perpendicular Line Generator

An ORD 2026 / MicroStation VBA tool that places independent straight cross-lines
at cumulative **geometry path distances**, on the active level with active
symbology. No source alignment, level definition, workspace variable, ProjectWise
configuration, or VBA autoload setting is changed.

**Support boundary:** this release evaluates standard VBA geometry. It does not
read the native ORD civil schema or certify an underlying civil spiral from its
display graphics. Native spirals that VBA cannot expose reliably require a Civil
SDK add-in. Do not stroke a spiral into a line string to bypass this limitation.
See [geometry and API notes](API_NOTES.md) before using civil alignments.

## Files

| File | Purpose |
| --- | --- |
| `PatternLineGenerator.mvba` | User-built ORD project with the modules, class and assembled UserForm |
| `PerpendicularLineGeneratorMain.bas` | Entry point, alignment/range session, capture of active settings, preflight and writing/rollback |
| `modPerpendicularGeometry.bas` | Geometry validation, component traversal, distances, derivatives and endpoints |
| `clsPerpendicularLocate.cls` | `ILocateCommandEvents` alignment selection and `IPrimitiveCommandEvents` range picking/dynamics |
| `frmPerpendicularLines.code.vba` | Complete UserForm code, including construction of all controls and event handlers |
| `modPerpendicularTests.bas` | Optional, in-memory ORD self-tests; no DGN elements added |
| `TEST_PLAN.md` | Repeatable acceptance tests, including native spiral support gates |
| `API_NOTES.md` | Verified Bentley API signatures, numerical policies and civil SDK alternative |

The included `.mvba` was built and saved by the user in ORD; it is not a
fabricated binary. Its four stored source components match the supplied source
after normalizing VBA editor casing, whitespace and export attributes. The
optional test module is not included in that project. The form code file is
**not an exported `.frm`**. There is no project-editing installer.

## Run the included project

1. Save `PatternLineGenerator.mvba` in an accessible folder, for example
   `C:\CAD_Custom\VBA`. If using OneDrive, choose **Always keep on this device**.
2. Open a disposable design model in ORD 2026 and enter these key-ins, replacing
   the load path with your saved file location:

   ```text
   VBA LOAD "C:\CAD_Custom\VBA\PatternLineGenerator.mvba"
   VBA RUN [PatternLineGenerator]PerpendicularLineGeneratorMain.Main
   ```

3. Select the alignment, pick start and end, review settings, then Generate Lines.

For a ribbon button, test this combined key-in, then use it as the button's
**Key-in** in **Customize Ribbon**:

```text
vba load "C:\CAD_Custom\VBA\PatternLineGenerator.mvba";vba run [PatternLineGenerator]PerpendicularLineGeneratorMain.Main
```

See the repository's [ribbon button instructions](../README.md#adding-a-vba-macro-to-a-ribbon-button-ord--projectwise).
The internal project name of the included file is `PatternLineGenerator`.
Rebuilding from source below is an alternative to loading that file.

## Install, assemble the form, and compile

1. Open a disposable **design model** in ORD 2026. Work in a model containing the
   alignment itself; reference attachment coordinates are not supported.
2. Open **VBA Project Manager** from ORD's Utilities/Macros tools. Create a **new**
   project and save it, for example as
   `C:\CAD_Custom\VBA\PerpendicularLineGenerator.mvba`. Do not import into an
   existing project unless you deliberately choose to do so.
3. Open the VBA editor. In **Tools > [project] Properties**, set the project name
   to `PerpendicularLineGenerator`. Select that project in Project Explorer.
4. Use **File > Import File** to import these three files:
   `PerpendicularLineGeneratorMain.bas`, `modPerpendicularGeometry.bas`, and
   `clsPerpendicularLocate.cls`. Optionally import `modPerpendicularTests.bas`.
   Remove the new project's unused empty standard module if desired.
   **If the `.cls` appears under Modules instead of Class Modules:** remove that
   incorrectly imported component, then choose **Insert > Class Module**. Press
   **F4** and set its **(Name)** to `clsPerpendicularLocate`. Open the supplied
   `.cls` in a text editor and copy only from `Option Explicit` through the end
   into the new class's code window. Do **not** paste the `VERSION`, `BEGIN/END`,
   or `Attribute` export-header lines. The resulting class must appear under
   **Class Modules**. No separate paste-source file is needed.
5. Choose **Insert > UserForm**. In the Properties window, set its **(Name)** to
   exactly `frmPerpendicularLines`. Leave the designer blank: add **no controls**.
6. Right-click this form and choose **View Code**. Paste the entire contents of
   `frmPerpendicularLines.code.vba` into the form's code window, replacing any
   blank default event stub. It begins with `Option Explicit`; paste it only once.
   `UserForm_Initialize` creates the complete dialog and connects its controls
   through `WithEvents` fields. **Importing the `.bas` modules alone will not
   create the form.**
7. **Tools > References** must include ORD's existing **Bentley MicroStation DGN**
   object library and **Microsoft Forms 2.0 Object Library**. Inserting the form
   normally supplies Forms automatically. Do not add Office/Excel or Civil SDK
   references, and do not register arbitrary DLLs.
8. Choose **Debug > Compile PerpendicularLineGenerator** and resolve any errors
   before running. Save the project. The authoring agent has not compiled or
   executed this source inside ORD; the supplied binary was saved by the user.
9. To distribute a genuine exported form later, right-click the assembled form
   and choose **Export File**. Keep its `.frm` and accompanying `.frx` together.
   An exported pair can be imported directly in another ORD VBA project instead
   of repeating steps 5–6. Do not rename or separate that pair.

No automatic access to `VBProject`/`VBComponents` is used or required. No trust
setting needs to be weakened. The source has no platform DLL declarations.

### Update an already working project

1. In the VBA editor, stop the macro with **Run > Reset** if it is active.
   Save/export a backup of your working project before replacing its source.
2. Replace the two standard modules by importing the updated
   `PerpendicularLineGeneratorMain.bas` and `modPerpendicularGeometry.bas`
   after removing the old modules of the same names from **your tool's project**.
3. Keep the existing **Class Module** `clsPerpendicularLocate`. Replace its code
   with the updated `.cls` contents **starting at `Option Explicit`**. Do not
   paste any `VERSION`, `BEGIN/END`, or `Attribute` export-header lines. The new
   body implements both `ILocateCommandEvents` and `IPrimitiveCommandEvents`.
4. Keep the existing blank UserForm `frmPerpendicularLines`. Replace its entire
   code with the updated `frmPerpendicularLines.code.vba`. No designer controls
   need to be added; the three new buttons are created when the form initializes.
5. If using the optional self-tests, replace that standard module too. Compile
   and save the project, then run its existing Main key-in. For a project named
   `PatternLineGenerator`, use:

   ```text
   VBA RUN [PatternLineGenerator]PerpendicularLineGeneratorMain.Main
   ```

Editing these source files does not update an already saved/loaded `.mvba`.
This repository update does not modify your existing `.mvba` automatically.

## Source-build ORD key-ins

With the file saved at the example path and the project name set as above:

```text
VBA LOAD "C:\CAD_Custom\VBA\PerpendicularLineGenerator.mvba"
VBA RUN [PerpendicularLineGenerator]PerpendicularLineGeneratorMain.Main
```

Optional self-test command, after importing the test module and ending the
generator session (including any active range picker):

```text
VBA RUN [PerpendicularLineGenerator]modPerpendicularTests.RunSelfTests
```

Replace the load path with your actual `.mvba` path. The bracketed run name is
the **internal VBA project name**. Load manually or bind these key-ins to a
personal button; this tool does not change autoload/search configuration.

## Dialog and usage

The modal dialog is titled **Perpendicular Line Generator**. A clear heading and
units notice appear above a read-only alignment status, active-level display and
working-unit labels. Aligned setting rows occupy the center, with endpoint
checkboxes, a live line-count/error area, and **Full Alignment / Generate Lines /
Cancel** below. **Pick Start** and **Pick End** buttons sit beside the distance
inputs. While picking, the dialog hides so the DGN receives mouse input.
The form uses Segoe UI, normal Windows controls, keyboard tab order, Enter for
Generate and Escape for Cancel. The selection status includes path length and
component count. No screenshot is supplied because the dialog has not been
rendered in ORD during implementation.

1. **Run the macro, then select the alignment.** The first prompt always starts
   graphical locate, even if elements were preselected. Identify the supported
   horizontal element and accept it with a data point. Reset with no candidate
   highlighted cancels. All components must pass the existing geometry checks.
2. **Pick Start.** Move near the alignment to preview a perpendicular cross-line
   at the closest XY point on the actual supported curve. The status prompt shows
   geometry distance in master units. Click to accept the start location.
3. **Pick End.** The accepted start cross-line stays in the dynamics preview while
   a second cross-line follows the mouse's closest point on the alignment. Click
   to accept an end strictly after the start in the chosen direction. This sets
   the range limit; the Include Ending Line checkbox still controls whether an
   output line is actually placed at that exact end. The settings dialog opens.
4. Review the active level and model master units. All numeric distances use
   **master units**, regardless of the model's master/subunit display format.
5. Adjust settings. **Pick Start** or **Pick End** hides the dialog to re-pick that
   limit, then returns to the dialog. **Full Alignment** sets start to zero and
   end to `End of Alignment`. The live estimate uses the generation distance-list
   builder. Invalid input disables Generate and displays the reason.
6. Click **Generate Lines**. The current active level, color, line style and line
   weight are captured at this point. Every coordinate, tangent and in-memory
   `LineElement` is validated before the first `AddElement` call.
7. Read the completion message and inspect the results. A generation error
   attempts to remove only this run's newly added elements. If cleanup fails,
   the message explicitly calls for Undo and inspection.

**Reset/right-click during a range pick** returns to the settings dialog without
accepting the moving cursor location. Any start already accepted remains set.
Use this to choose Full Alignment, change Direction, adjust preview length or
placement, or cancel. An end before/equal to start is not accepted: choose a later
location or Reset, change Direction, and re-pick. Switching to another ORD command
ends the pick session; run the tool again to select an alignment. Panning/zooming
that merely suspends the command should allow picking to resume (verify in ORD).

Preview elements are only drawn in dynamics; they are never added to the DGN.
No output is created until Generate Lines. The preview uses the last-used total
length, placement and direction, initially 200/centered/forward. Distances saved
from picks retain full numeric precision; the three-decimal status display does
not round the stored values. These are standard VBA dynamics/status prompts,
not the native civil offset tool's floating parameter editor, Alt locks, or civil
offset rules. The selected alignment remains untouched.

Settings survive Cancel, the window close button, and successful runs through
the default form instance for the current loaded VBA session. They reset when
the project unloads, VBA is reset, or ORD closes. Values are not saved to the DGN,
registry, or another project. Reopening the dialog refreshes alignment/level/units.

| Setting | Initial value |
| --- | --- |
| Line Spacing | `150.00` |
| Total Line Length | `200.00` |
| Placement | Centered on Alignment |
| Start Distance Along Alignment | `0.00` |
| End Distance | `End of Alignment` (blank also means the geometric end) |
| Include Starting Line | Yes |
| Include Ending Line | No |
| Direction | Forward |

### Distance and direction rules

- Forward uses the selected element's stored direction. Reverse measures zero
  from the other end and reverses the tangent. Left and right follow the **chosen
  travel direction** looking forward in plan (global XY), not the current view.
- Picked values use the Direction selected **at the time of the pick**. Changing
  Direction later reinterprets the numeric distances from the other end, as with
  typed distances. Re-pick both limits if you want to keep particular physical
  locations after changing direction.
- Start/end inputs are increasing distances from that direction's geometry
  beginning. They are **not** displayed ORD stations, starting station offsets,
  or station-equation values. Enter `150`, not `1+50` or `150 ft`.
- Regular locations are `start + k * spacing`. Excluding the starting line skips
  `k = 0`; it does not shift the spacing grid.
- The ending line is independently controlled. When excluded, a regular-grid
  location coinciding with the end is also excluded. When included, the exact
  requested end is placed **once**, even if not on the grid. That final interval
  can therefore be shorter than the spacing.
- Centered placement puts half the **total** line length on either side. Left
  and right placement put the full length on that side, anchored at the alignment.
- Generated lines keep the horizontal alignment's Z; the tool does not sample a
  terrain or use the active elevation setting.

### Feet example

In a model whose master unit is feet, select an eastbound 1,000-foot line, pick
its beginning and end (or Reset and choose Full Alignment), and accept the other
defaults: 150-foot spacing and 200-foot total line length. Seven
centered cross-lines occur at geometry distances **0, 150, 300, 450, 600, 750,
900 feet**. Each extends 100 feet left and 100 feet right. Including the ending
line adds one at 1,000 feet with a 100-foot final interval.

In a meter model the same numeric inputs mean **150 meters and 200 meters**;
no feet conversion is applied. The exact model unit definition controls whether
feet are international or US survey feet; the short label `ft` alone may not
distinguish them. Inspect ORD's working-unit definition when that distinction
matters.

## Supported geometry and limits

Supported: open horizontal lines, line strings, circular/elliptical arcs, regular
B-spline curves, and connected complex strings recursively containing only those
types. An ORD civil alignment can be used **only to the extent its selected VBA
element exposes this supported geometry**. The selected representation is the
computational definition for this macro.

Curve evaluation uses Bentley B-spline distance integration/inversion and its
analytic local derivative. Arc and line conversions do not use strokes. Spacing
continues through component joins and never restarts per component. Line string
segments are used only for actual selected line string geometry, not as a
replacement for an accessible curved element.

Not supported: native civil schema/station equations, opaque native transition
spirals/custom or application elements, cells, shared cells, shapes, closed paths,
legacy type-11 curves, reference attachments, sheet models, or geometry with changing Z/tilted
planes. Internal spline discontinuities are rejected. A requested placement at a
non-tangent line-string/complex join is rejected rather than assigning an
arbitrary perpendicular. Gaps and ambiguous component ordering are rejected.
The macro does not repair, reverse, or drop the source alignment.

Graphical picking evaluates every supported component with Bentley's closest-point
API, then converts its returned parameter to cumulative curve length. Cursor Z
is ignored; use plan position. Distinct equally close locations on different
components are rejected (including a crossing), while a shared join with one
chain distance is allowed. The API returns one minimum per individual B-spline;
it cannot establish uniqueness for multiple equally close points inside that
single component. Use unambiguous picks or typed distances for such geometry.
Very large component counts may make mouse previews slow because every component
is checked. The native civil-spiral support boundary still applies.

A native spiral displayed as a spline or stroke may already be an approximation.
VBA geometry alone cannot prove equivalence to its civil definition. Do not use
the output for native-spiral engineering work without the comparison gate in
[the test plan](TEST_PLAN.md). This release supplies **no native spiral or
stationing bridge**. Bentley's managed Civil SDK is the documented route for a
future ORD-native evaluator; see [API notes](API_NOTES.md).

Safety limits: at most 10,000 output lines, 10,000 primitive components, 16 nested
complex-string levels, and 1 billion master units for path length/numeric inputs.
Spacing and line length must exceed four computational distance tolerances.
The active model must be writable and the active level must permit modification
and deletion, allowing cleanup after a failed write. No levels are unlocked.

Ordinary `AddElement`/`Rewrite`/`RemoveElement` operations participate in
MicroStation's normal Undo history. A single grouped Undo transaction is **not
promised**. Verify the required number of Undo steps in your ORD build. Runtime
write failures are handled by best-effort removal, not a database transaction;
inspect any reported cleanup failure before retrying.

## Troubleshooting

| Symptom | Action |
| --- | --- |
| `Variable not defined: frmPerpendicularLines` | Create and name the actual blank UserForm, then paste its code; `.bas` import is insufficient. |
| `User-defined type not defined` / `MISSING` reference | Check Bentley MicroStation DGN and Forms references in the ORD VBA editor. Use ORD's supplied libraries. |
| Macro cannot be found | Verify the loaded file and internal project name, then use the full module/procedure key-in above. Compile first. |
| Form is blank in the designer | Expected: controls are constructed when the form runs. Paste the complete form code if it is blank at runtime. |
| Unsupported native civil element/spiral | Use a verified standard geometry representation or an ORD Civil SDK tool; do not drop/stroke the alignment to hide the limitation. |
| Corner, cusp, zero tangent, gap or ordering error | Inspect geometry; change the placement distances to avoid a corner, or select a continuous smooth path. No source repair is attempted. |
| Curve distance/point validation fails | Stop and retain the failing test DGN/build details. Do not relax the tolerances merely to pass; check conversion and evaluator behavior in that ORD version. |
| End distance rejected after reopening | The saved numeric end may exceed this alignment. Enter `End of Alignment` or a valid distance. |
| First prompt asks for the alignment despite a preselection | Expected: the new workflow always locates the alignment, then picks start and end. |
| Need settings before choosing the range | Reset during a range pick to open the dialog. Set Direction/length/placement, then use Pick Start/Pick End, or Full Alignment. |
| End will not accept / preview prompt says end must follow start | Pick farther along in the current Direction, or Reset to change Direction and re-pick. |
| Pick is equally close to different locations | Move away from a crossing/overlap or enter the desired geometry distance explicitly. |
| Preview works but no lines appear in the element count | Expected until Generate Lines: dynamics never calls AddElement. |
| START/END instructions appear at the lower left and native Tool Settings is empty | Expected: ShowPrompt uses the status bar. See API_NOTES.md for the separate modeless VBA panel and native .NET Tool Settings options; neither relocation is implemented yet. |
| Switching commands closes the picker | Expected: this releases the source/session rather than reopening the dialog over the new command. Run the tool again. |
| Too many lines / below model resolution | Increase spacing/line length or shorten the range. Review model working-unit resolution. |
| Level/model is locked or read-only | Choose a writable active model and an unlocked writable active level; this macro does not change access permissions. |
| Failure during write / rollback failure | Follow the message, inspect output, and use normal Undo as needed. Never assume a failed write was atomic. |

## Verification status

Bentley API names/signatures were checked against the **installed ORD 2026.00
`MicroStationVBA.chm` and `Bentley.Interop.MicroStationDGN.dll`**. The source was
checked for structural consistency. These are static checks, **not** a VBA compile
or execution in ORD. The form, in-memory self-tests, DGN writes, symbology,
selection callbacks, native civil representations, and Undo still require
execution of [TEST_PLAN.md](TEST_PLAN.md) in ORD 2026. Record actual build numbers
and results there or in a separate test record.

The included user-built project's stored source was also compared with the four
supplied source components; all matched after normalization. This read-only
comparison does not compile the project or validate its runtime geometry.

The user reported the original generator working after the naming/compile fixes.
The newly added graphical range selection and dynamics have **not** been run
inside ORD by the authoring agent; the new acceptance tests remain pending.
The user subsequently reported graphical picking/preview working and supplied
an ORD screenshot. This is user-reported operational feedback, not completion of
the measured geometry/behavior acceptance tests.
