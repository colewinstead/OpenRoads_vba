# Geometry and Bentley API notes

## Verification sources

The primary VBA reference used during implementation was the help file delivered
with the installed ORD 2026.00:

```text
C:\Program Files\Bentley\OpenRoads Designer 2026.00\OpenRoadsDesigner\MicroStationVBA.chm
```

Signatures and interface membership were also inspected by read-only reflection
of its delivered `Assemblies\Bentley.Interop.MicroStationDGN.dll`. No ORD process
was launched for that inspection. In the VBA Object Browser, use **F2** and search
the Bentley library for the members below. In the help file, search their names.
The HTML topic identifiers below make the reference repeatable after extracting
the help with Windows HTML Help's `-decompile` option.

| Member / topic | Use |
| --- | --- |
| `CommandState.StartLocate` and `ILocateCommandEvents` / `usobjilocatecommandevents.htm`, `ussmpLocate.htm` | Graphical locate, filtering, accept, reset and cleanup signatures |
| `CommandState.StartPrimitive`, `IPrimitiveCommandEvents` / `usobjiprimitivecommandevents.htm`, `ussmpPrimitive.htm` | Start/end data-point callbacks, Reset, key-in and command cleanup signatures |
| `CommandState.StartDynamics`, `EnableAccuSnap`; `Element.Redraw(DrawMode)` / `usmthstartdynamics.htm`, `usevtdynamics.htm` | Cursor-driven temporary cross-line previews; only Redraw inside Dynamics, never AddElement |
| `Element.IsTraversableElement`, `AsTraversableElement` | Test/access traversable geometry |
| `TraversableElement.PointAtDistance(Distance)` / `usmthpointatdistance.htm` | Original geometry endpoints for validation |
| Typed `LineElement.Length`, `ArcElement.Length`, `BsplineCurveElement.Length`, `ComplexStringElement.Length` | Original element length; **`TraversableElement` does not have `Length`** |
| `Element.AsVertexList.GetVertices` / `usmthgetvertices.htm` | Actual line string's zero-based vertices; there is no `AsLineStringElement` cast |
| `ComplexStringElement.GetSubElements` / `usmthgetsubelements.htm` | Enumerate component copies without changing the source |
| `BsplineCurve.FromElement(Element)` / `usmthfromelement.htm` | Convert supported line/arc/spline components to evaluable geometry objects |
| `BsplineCurve.ComputeCurveLength([StartParameter, EndParameter, IgnoreZ])` / `usmthcomputecurvelength.htm` | Full/partial curve length; not chord length |
| `BsplineCurve.EvaluatePointAtDistance(Parameter, DistanceFromStart, [StartParameter, Tolerance])` / `usmthevaluatepointatdistance.htm` | Return point and **output** curve parameter from a distance; explicit relative integration tolerance `1e-12` |
| `BsplineCurve.EvaluatePointTangent(Tangent, Parameter)` / `usmthevaluatepointtangent.htm` | Return point and **output** unnormalized analytic tangent |
| `BsplineCurve.ComputeMinimumDistance(ClosestPoint, ClosestParameter, FromPoint, Rotation, [Tolerance])` / `usmthcomputeminimumdistance.htm` | Closest point and **output** parameter; identity rotation computes XY distance, ignoring cursor Z |
| `BsplineCurve.EvaluatePoint`, `Pole`, `PolesCount`, `Knot`, `KnotsCount`, `Closed` | Independent point check, horizontal plane and knot checks |
| `Application.CreateLineElement2(Template, StartPoint, EndPoint)` / `usmthcreatelineelement2.htm` | Prepare independent straight output elements using active settings with `Template = Nothing` |
| `Element.Level` / `usprolevel.htm` | A real `Level` object can be assigned only **after** insertion; newly created elements get active level automatically |
| `ActiveSettings.Level`, `Color`, `LineStyle`, `LineWeight` | Capture settings on Generate; `LineStyle` is an object assigned with `Set` |
| `ModelReference.AddElement`, `Element.Rewrite`, `ModelReference.RemoveElement` | Write and best-effort rollback |
| `ModelReference.IsReadOnly`, `Level.ElementAccess` | Preflight writable output; `DesignFile.IsReadOnly` is not used |
| `ModelReference.IsActive`, `IsAttachment`, `Type` | Check the captured model is current, reject attachments and exclude sheet models |
| `ModelReference.MasterUnit`, `SubUnit`, `UORsPerMasterUnit`; `MeasurementUnit.Label` | Working units and computational resolution |

The optional self-test also uses documented `CreateLineElement1`,
`CreateArcElement2`, `CreateComplexStringElement1` with a `ChainableElement()`
array, `CreateBsplineCurveElement1`, and `BsplineCurve.SetPoles/Order/SetKnots`.
No deprecated distance-to-parameter method, MDL pointer/DLL declaration, C
expression, guessed `PointAtStation` method, or Civil COM object is used.

The load/run syntax follows Bentley's
[Getting Started with Visual Basic](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0109946)
and [Running a VBA Routine by Keyin](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0039735).

## Geometry policies

These are **implementation safeguards, not roadway design criteria**:

- Distance tolerance `T = max(2 / UORsPerMasterUnit, sourceLength * 1e-10)`, in
  master units. It accounts for two model resolution units plus an accumulated
  numerical length budget. It is used for endpoint matching,
  length residuals, and coincidence with range ends/joins. The code does not
  round placement distances to displayed precision.
- All component poles must share the geometry origin's Z within
  `max(1e-12, abs(originZ)*1e-14)` master units, a floating-point plane-check
  budget. This is deliberately tighter than the distance tolerance so small
  slopes or accumulated component elevation changes are not mistaken for a
  horizontal alignment.
- `EvaluatePointAtDistance` receives a **relative integration error request** of
  `1e-12`. Each result is checked by recomputing partial curve length from zero
  to the returned parameter. This is a cross-check within Bentley's curve engine,
  not an independent proof of the engine. The returned point is also compared
  with `EvaluatePointTangent`'s point.
- B-spline parameter is **not** generally a fraction of length. The code never
  uses `distance / totalLength` as the parameter.
- XY tangents are normalized and rotated left to `(-ty, tx)`. Centered endpoints
  are `p +/- length/2 * normal`. Left-only uses `p + length * normal` to `p`;
  right-only uses `p` to `p - length * normal`. All endpoints retain `p.Z`.
- Components keep their stored order; a component is traversed backwards if its
  end matches the preceding endpoint. No gap is bridged and no source is reversed.
  Ordering that cannot be reconciled with the parent's endpoints is rejected.
- At requested component joins, endpoint tangents must agree within a unit-vector
  difference of `1e-6` (approximately one microradian). A non-smooth corner has
  no unique mathematical tangent, so placement there fails the whole preflight.
- Internal spline knots are checked on both sides using Bentley's derivatives;
  no numerical derivative or nearest-segment tangent is used for a curved
  component. Side probes use `1e-8` of the neighboring parameter span, with a
  further factor of `0.01` for point-limit checks. Very small spans, extreme
  curvature, or near-degenerate splines may be conservatively rejected. These
  probes are discontinuity guards, not a certification of a civil spiral fit.
- Graphical picking calls `ComputeMinimumDistance` for every component with
  `Matrix3dIdentity`, checks the returned point against `EvaluatePoint`, and checks
  its reported distance against explicit XY distance. Cumulative path distance
  is the preceding components' lengths plus `ComputeCurveLength(0, parameter)`,
  with component/global reversal applied. It is never the distance from the
  cursor to an endpoint, a parameter fraction, or a snapped polyline surrogate.
- All component candidates within `T` of the best XY distance are checked for
  conflicting chain distances. Shared endpoints with the same distance are
  allowed; crossings/overlaps with distinct distances are rejected. Bentley's
  method returns only one minimum per individual B-spline. Multiple equally
  close locations **within** one B-spline cannot be resolved by this interface;
  use unambiguous cursor positions or type a distance. Its optional tolerance
  argument is documented as **unused for B-splines**, so none is supplied and no
  tighter closest-point convergence is claimed from that parameter.
- Dynamics draws the moving limit and, during end selection, the fixed start.
  It passes the event's `DrawMode` unchanged to `Redraw`. It creates only temporary
  in-memory lines and does not mutate the data-point argument or source geometry.
  Accepted picks stop the primitive command before resuming the modal form or
  next pick. Cleanup caused by another command releases the session without
  opening a modal dialog over that command.
- Inputs/geometry below these budgets are rejected. Do not weaken safeguards to
  make an unexplained failure pass. Investigate the original geometry and ORD
  build against the analytical fixtures first.

## Status prompts and the native Tool Settings panel

`ShowPrompt` and `ShowCommand` are documented to write to MicroStation's status
area, explaining the lower-left range text. They do not populate Tool Settings.
The installed ORD 2026 VBA `Application`/`CommandState` interfaces checked here
expose no Tool Settings form-attachment API.

Bentley's [Adding Windows Forms to Add-ins](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0012740)
documents native Tool Settings integration through a managed form derived from
`Bentley.MstnPlatformNET.WinForms.Adapter`, using
`AttachToToolSettings(MyAddin.Addin)`. That is a .NET add-in interface, not a
method available on this VBA `UserForm`.

Bentley also documents [modeless VBA forms](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0028347).
A floating modeless VBA panel could show the current step/distance near the
right-hand tools while mouse picking continues. It would remain a separate
window. Native panel integration needs the SDK add-in route. Neither UI change
has been implemented; the choice is pending. Simply changing the existing
modal `.Show` to `vbModeless` would require changing its session callbacks too.

## Native ORD geometry and the documented alternative

Bentley documents the civil schema separately from basic DGN geometry in
[Understanding Geometry model](https://docs.bentley.com/LiveContent/web/OpenRoads%20Designer%20SDK%20Help-v1/en/GUID-E99BC4FE-A0C9-4C20-8756-E8104911C083.html).
The official examples obtain a managed `ConsensusConnection`, a
`GeometricModel`, and its native `Alignment` objects. This supports the decision
to keep native alignment/stationing work in an ORD-specific SDK layer rather
than assuming VBA's ordinary DGN elements expose it.

Bentley's [ORD SDK Developer Guide](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0012549)
documents managed and unmanaged development, building examples, running custom
projects inside ORD, and debugging managed projects. The
[Civil Programming index](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0012546)
links the SDK downloads, installation, examples and geometry-model documentation.

For native spiral/station-equation support beyond the verified exposed VBA
geometry, a **separate Bentley Civil SDK/.NET add-in is required by this
implementation's support contract**. A future evaluator must read the native
alignment's actual horizontal primitive definitions, use a validated native
distance/tangent traversal, and separately translate displayed stationing and
station equations. It must be built and tested against the installed ORD 2026
SDK version. No claim is made here that one particular undocumented managed
method or a generic curve conversion handles all native spirals.

The absence of a spiral-specific evaluator in this VBA release does not prove
that every native alignment is inaccessible. It means the tool cannot establish
the native definition's accuracy through these ordinary VBA interfaces alone.
If the selected element is a B-spline representation, the macro evaluates that
**B-spline**, with no guarantee about its fit to an underlying clothoid. If it
is a line string, it evaluates that **line string**, with no claim that its chord
length equals the native spiral length. The manual support gate must establish
equivalence before use for civil geometry.
