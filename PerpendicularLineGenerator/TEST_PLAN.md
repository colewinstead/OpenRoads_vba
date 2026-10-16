# ORD 2026 repeatable acceptance tests

**Status: not executed inside ORD during implementation.** Static signature and
source checks do not substitute for these tests. Compile and run in a disposable
DGN/project first. Record ORD's full build, model dimensionality, working-unit
definition, UORs per master unit, tolerance, and measured results. Do not promote
native spiral support merely because output looks perpendicular on screen.

## Preparation and evidence

1. Assemble/import as described in [README.md](README.md), then use **Debug >
   Compile**. Record compiler success/failure. Run the optional
   `modPerpendicularTests.RunSelfTests` with the generator/picker inactive. It creates
   only in-memory fixtures, verifies analytical lines/arcs/B-splines, cumulative
   joins, both directions, placements, counts, and rejection cases. It does not
   exercise AddElement, form layout, selection, or the native civil schema.
2. Use a scratch design model with global XY plan geometry. Choose an existing
   writable level with distinctive color/style/weight; record active settings.
   Repeat with explicit symbology and ByLevel settings. Record original alignment
   IDs, lengths, and control geometry before and after each test.
3. Record the output IDs/count. Confirm each output is an ordinary independent
   two-point line, in the active model and on the captured active level. Inspect
   color, style and weight (including preservation of ByLevel values).
4. For each placement, obtain the actual alignment anchor (centered midpoint or
   one-sided endpoint). Compare XY with an analytical reference or an independent
   native ORD distance/tangent report. Check path distance from the direction's
   origin, not Euclidean distance between anchors.
5. For line direction vector `v = endpoint2 - endpoint1` and reference unit tangent
   `t`, verify `abs(v.x*t.x + v.y*t.y)/|v| <= 1e-6`. Check total length to distance
   tolerance `T = max(2/UORsPerMasterUnit, pathLength*1e-10)` and coordinate/path
   residuals to `T` (or explicitly record stricter project acceptance criteria).
   `T` is the macro's numerical budget, not an agency design standard. Do not
   compare against the macro's tangent as the sole independent reference.
6. Test with no selection, exactly one preselection, multiple preselections, an invalid
   selection, graphical locate/accept, rejection of an unsupported candidate,
   Reset cancellation, and reference attachments. Check Cancel, Escape, and X
   produce no elements; reopening retains settings while refreshing context.
   The new workflow always prompts for alignment selection first, followed by
   start and end picking. Preselection does not bypass the alignment prompt.
7. Record actual form layout at your display scaling: no clipped rows/messages,
   keyboard tab order, Enter, Escape, numeric locale, read-only displays, live
   estimate and invalid input messages. All controls are created at runtime.

## 1. Straight 1,000-foot alignment

Use a feet-master model and a horizontal line `(0,0,0)` to `(1000,0,0)`. With
spacing 150, total length 200, centered, start 0, end-of-alignment, starting line
included and ending line excluded:

- Expect **7** lines at X = `0,150,300,450,600,750,900`.
- Each line has Y endpoints `+100,-100` and Z = 0, and is exactly 200 feet long
  within `T`. Reference tangent is `(1,0)`.
- Include Ending: **8**, with a final line at X = 1000 (last interval 100).
- End = 900: ending excluded gives **6** at 0 through 750; including it gives **7**,
  never two coincident ending lines. Exclude both endpoints: **5**, 150 through 750.
- Start = 50, end = 700: regular distances `50,200,350,500,650`; including the end
  adds 700. Excluding the starting line removes only 50.

## 2. Circular curve, known radius

Create a counterclockwise quarter circle with center `(0,0)`, radius 500 feet,
start `(500,0)`, end `(0,500)`, all at Z = 0. The path length is
`500*pi/2 = 785.398163397448...` feet. Default spacing gives **6** lines at
`0,150,300,450,600,750`; including the end gives **7**.

At distance `s`, independent references are:

```text
theta = s / 500
p = (500*cos(theta), 500*sin(theta))
t = (-sin(theta), cos(theta))
```

At s = 150, theta = 0.3 radians. A centered 200-foot cross-line is radial, with
endpoint radii 400 and 600 feet. Anchor-to-anchor straight distances are shorter
than 150; **do not** treat that as a spacing failure. Repeat clockwise and through
an arc's internal rational B-spline span boundary to exercise knot checks.

## 3. Clothoid spiral transition: native support gate

Create an actual ORD clothoid with origin `(0,0)`, initial heading 0, length 300
feet and curvature increasing from 0 to `1/1000 ft^-1` (left turn). This is a
**native civil test**, not the cubic B-spline fixture in the self-test module.

The mathematical reference is `theta(s) = s^2 / 600000`, with
`x = integral(cos(theta(s)), ds)` and `y = integral(sin(theta(s)), ds)`.
The following independent values were computed from the convergent Fresnel
power series using 55-digit decimal arithmetic (displayed rounded):

| s (ft) | x (ft) | y (ft) | tangent heading (rad) |
| ---: | ---: | ---: | ---: |
| 0 | 0.000000000000 | 0.000000000000 | 0.000000000000 |
| 75 | 74.999340822995 | 0.234373528621 | 0.009375000000 |
| 150 | 149.978907623246 | 1.874811671373 | 0.037500000000 |
| 225 | 224.839872121185 | 6.324907813699 | 0.084375000000 |
| 300 | 299.325702760028 | 14.975910108887 | 0.150000000000 |

Select the native element and record its VBA type/components. An opaque or
unsupported component **must be rejected with no writes**. If ORD exposes only
a B-spline or line-string representation, record that fact explicitly: the tool
evaluates that representation, not the civil clothoid definition. A successful
run alone **does not pass native spiral support**. Compare distance, point and
tangent against the table and native ORD reports before any engineering use.
If the representation disagrees, do not use it and do not lower tolerances; the
native Civil SDK implementation remains required. Native spiral traversal is
not implemented by this release.

For an accepted, independently verified spline representation, use spacing 75
and Include Ending to compare all five points. Check the first and last tangent,
full path length, centered length and left/right orientation. Do not count the
generic B-spline self-test as evidence of native spiral fidelity.

## 4. Tangent–spiral–curve–spiral–tangent

Build a native continuous alignment in feet:

| Part | Length | Curvature / heading |
| --- | ---: | --- |
| Initial tangent | 1000 | Heading 0 |
| Entry clothoid | 300 | Curvature 0 to +1/1000; heading 0 to 0.15 rad |
| Circular arc | 600 | R = 1000; heading 0.15 to 0.75 rad |
| Exit clothoid | 300 | Curvature +1/1000 to 0; heading 0.75 to 0.9 rad |
| Final tangent | 1000 | Heading 0.9 rad |

Distances: TS = 1000, SC = 1300, CS = 1900, ST = 2200, end = 3200. A native
evaluator would give **22** regular lines at `0,150,...,3150`, or **23** with the
end included. Spacing **must not reset** at the four joins. Apply the same native
support gate as test 3; rejection is the expected safe result when native spiral
geometry is inaccessible. Do not claim this native test passed from display
geometry alone.

For independently verified exposed geometry, compare positions/tangents just
before, at, and just after each join (e.g. `join +/- 0.001 ft`). Force exact join
placements using start = 1000, spacing = 300 and include start. For reference:
SC is `(1000 + 299.325702760028, 14.975910108887)`. The circle displacement is
`1000*(sin(0.75)-sin(0.15), cos(0.15)-cos(0.75))`. The exit spiral displacement
is the entry displacement `(299.325702760028, -14.975910108887)` rotated by
0.9 radians. These definitions permit an independent composite reference.

Also run a fully supported ordinary tangent–arc–tangent complex string to check
cumulative transitions without a native spiral. The in-memory self-test checks
the first tangent/arc join at distances 999.999, 1000 and 1000.001, and at 1050.

## 5. Reversed direction

Reverse the stored direction of the straight fixture using ORD on a **copy**.
Forward default anchors should now run X = `1000,850,...,100`. Left should be
south and right north. Test Reverse in the dialog on the original eastbound
fixture: the same positions/sides should result, without changing its stored
direction. Reverse on the reversed copy restores the original direction.
Repeat on the arc and supported complex string; compare `p(L-s)` and `-t(L-s)`.

## 6. Placement, level, symbology, Z and Undo

- Eastbound straight: centered endpoints Y = +100/-100; left-only Y = +200/0;
  right-only Y = 0/-200. All lengths are 200. Reverse swaps physical sides.
- Use a horizontal element at Z = 25 in a 3D model. Output endpoints must keep
  Z = 25. Reject a sloping 3D line; do not replace its 3D path length with XY length.
- Repeat with a different active level/color/style/weight between runs. Check
  Generate uses the current settings. For a precise timing check, break at the
  beginning of `PLGGenerate` in a scratch project, change the active settings in
  the VBA Immediate window, then continue; the captured values must be used.
- No new levels, level edits, source changes, or non-line output elements.
- Use ordinary Undo until the run's additions disappear. Record the number of
  steps; a one-step group is not promised. The original geometry must remain.

## 7. Non-foot master units

In a meter-master model create a 1000-meter line and enter 150/200. Expect seven
lines separated by 150 **meters**, each 200 meters long. Repeat in millimeters
with suitably sized geometry. No automatic foot conversion is expected. Check
the read-only master/subunit display. Repeat international-foot and US survey-foot
models; verify interpretation from the actual unit definition, not the short
unit label. Use local decimal notation and test invalid unit/station suffixes.

## 8. Invalid inputs, unsupported geometry and write errors

Try spacing/length 0, negative, blank, text, out-of-range numbers and values below
four `T`; start < 0, start >= end, end beyond length, and a >10,000-line request.
Try a range shorter than spacing with both endpoints excluded. Expect a clear
error and zero elements, including when errors occur on a later placement.

Test shapes, closed line strings, cells, references, opaque civil components,
legacy type-11 curves, gaps, zero-length components, changing Z, and a line
string with a sharp corner exactly at a requested placement. Reject these
cases as documented. The same corner need not block placements strictly inside
its straight segments. Reject discontinuous/cusped spline knots. Test both
read-only models, sheet models, and locked/read-only levels: no configuration/access changes.

For rollback, use a separate disposable project copied from this source. Insert
a temporary `If i = 2 Then Err.Raise vbObjectError + 2999, , "Test write failure"`
inside the write loop **before** `AddElement`. Generate at least two lines.
Expect the first line removed and the source intact; remove that test injection
and recompile afterward. This tests removal after one successful write. If a
real cleanup failure occurs, verify the message identifies it; record it as a
failure requiring manual inspection/Undo. No filesystem or DGN transaction is
claimed, and disk failure cannot be ruled out by this injected test.

## 9. Exact endpoints and session state

Check zero and exact geometry end on straight, arc and supported spline/complex
fixtures, on/off-grid ends, forward/reverse, with all four checkbox combinations.
No duplicate ending line may occur. Check nonzero starts anchor their own grid.
Verify a new alignment/model refreshes displays while retaining settings, and
that an old numeric end outside the new path gives an error. Unload/reload the
project or reset VBA: defaults should return. Cancel and invalid input must leave
both the original alignment and the element count unchanged.

## 10. Dynamic range workflow (new; pending ORD execution)

1. Run Main with zero, one, and multiple preselected elements. Each run must
   first prompt to identify/accept the alignment, then **START**, then **END**,
   then show settings. No DGN elements should be added during these steps.
2. On the 1000-foot eastbound line, pick near `(100,40)` and `(700,-40)` in plan.
   The anchors must be `(100,0)` and `(700,0)`, and the fields must read 100 and
   700 master units. With spacing 150, Include Start true and Include End false,
   expect **4** output lines at 100,250,400,550 after Generate. Include End true
   should produce **5**, adding 700 exactly once.
3. During START, confirm the moving cross-line matches total length and placement.
   During END, confirm the fixed start and moving end both display. Move through
   arc/line/spline transitions: the status distances must remain cumulative, and
   the moving line must remain perpendicular to the actual local tangent.
4. Test a cursor with a different Z in a 3D model: the closest **XY** position
   should be unchanged. The in-memory tests now check straight, circular,
   tangent/arc composite, reversed and cubic B-spline nearest-point references.
   Check past-end picks clamp to actual geometry endpoints.
5. Reset during the initial START: open settings with prior numeric values
   unchanged. Reset during END: keep the accepted start and leave the prior end
   unchanged. Reset during a re-pick: leave that value unchanged. Cancel or X
   afterward must create no elements. Full Alignment must set 0 / End of Alignment.
6. Change Direction to Reverse in settings, then use Pick Start and Pick End.
   On the same line, physical X=900 to X=300 must store distances 100 to 700.
   Left/right must follow reverse travel. Changing Direction after picking
   reinterprets numeric distances; re-pick to establish particular physical limits.
7. Try an END before/equal to START: retain the command, explain the ordering
   requirement, and add no elements. Reset must still open settings so Direction
   can be changed. Try a corner/cusp, wrong model, unsupported source and invalid
   preview length; verify useful errors and recovery without output.
8. On a complex/line-string crossing, click exactly where distinct path distances
   share the closest XY point: reject the ambiguous pick. A normal shared join
   with one path distance must remain pickable. Multiple equally close minima
   inside a **single** B-spline remain an API limitation; use typed distances.
9. Switch to another ORD command while picking: no modal form should reopen over
   it and no elements should appear. Re-running Main must start a fresh alignment
   selection. Repeat after changing active model. Test view pan/zoom suspension
   and resumption separately; record actual behavior in your ORD build.
10. Confirm only Generate writes. The chosen limits and preview must not create
    levels, alter the alignment, or add transient geometry to the persistent DGN.
    Compare performance on a larger complex string; all components are searched
    each mouse move, so high-complexity paths may be slow. Verify the actual
    Dynamics drawing/erase behavior in ORD 2026, not just the callback signatures.

## Test record

| Test | ORD full build / model / units | Measured results / evidence | Pass, fail, or unsupported |
| --- | --- | --- | --- |
| Compile + in-memory self-tests | Pending | Not run in ORD | Pending |
| 1 straight | Pending | | Pending |
| 2 circle | Pending | | Pending |
| 3 native clothoid gate | Pending | Native evaluator not shipped | Unsupported unless independently established |
| 4 native composite gate | Pending | Native evaluator not shipped | Unsupported unless independently established |
| 5 reversed | Pending | | Pending |
| 6 placements/symbology/Z/Undo | Pending | | Pending |
| 7 units | Pending | | Pending |
| 8 errors/rollback | Pending | | Pending |
| 9 endpoints/session | Pending | | Pending |
| 10 dynamic selection/range | Pending | User reported earlier generator working; new picking/dynamics not run by agent | Pending |
