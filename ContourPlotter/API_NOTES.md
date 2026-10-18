# Bentley API evidence

## Installed reference and scope

Inspected the installed ORD 2026.00 assembly:
C:\Program Files\Bentley\OpenRoads Designer 2026.00\OpenRoadsDesigner\Assemblies\Bentley.Interop.MicroStationDGN.dll

The adjacent MicroStationVBA.chm exists. Attempts to decompile it with Windows
HTML Help returned no topics in this environment. Exact signatures/interface
membership were verified by read-only reflection of the installed assembly,
rather than inventing methods. The existing PerpendicularLineGenerator API notes,
geometry module, locate class, form and tests were also inspected. Its AddElement
then Level then Rewrite pattern is retained; its source/project is unchanged.

The live open ORD reported **26.00.00.154**. Read-only COM calls plus the source in
tests/ReadOnlyOrdProbe.cs queried units, GCS, available local levels/styles and
prepared unsaved line/text objects. This did not compile or execute an MVBA and
did not call AddElement, Rewrite, RemoveElement or Save.

## Verified signatures used by the writer

| Member | Exact VBA use / evidence |
| --- | --- |
| CreateLineElement1(Template As Element, Vertices() As Point3d) As LineElement | Original ordered vertices; a 1,024-point unsaved line was prepared, ID=0 |
| CreateTextElement1(Template As Element, Text As String, Origin As Point3d, Rotation As Matrix3d) As TextElement | Ordinary unsaved text prepared, ID=0 |
| Matrix3dFromAxisAndRotationAngle(Axis As Long, Radians As Double) As Matrix3d | Z axis index 2; signature verified |
| TextElement.TextStyle As TextStyle | Setter/getter verified; element-local style override and reapplication probed |
| TextElement.GetTotalTextSize(Width As Double, Height As Double) | Output parameters; actual model-space width/height queried |
| Element.ScaleUniform(Origin As Point3d, Scale As Double) | Prepared text physical size normalized without active-setting changes |
| TextStyle.Height, Width, Justification, UsesColor, BorderAndBackgroundVisible | Verified members; source style height/width unchanged after the live probe |
| ModelReference.AddElement(Element); Element.Level; Element.Rewrite | Signatures verified, insertion order adopted from existing tool; no write test performed |
| ModelReference.RemoveElement(Element) | Signature verified; rollback/removal not executed |
| ByLevelColor(), ByLevelLineWeight(), ByLevelLineStyle() | Function signatures verified; setters applied to unsaved native line |
| DesignFile.Levels; Levels.Find(Name, Nothing); Level.Name, IsLocked, IsFromLevelLibrary, ElementAccess | DGN and attached-library levels enumerated; exact-name lookup retains write-access checks |
| DesignFile.TextStyles; TextStyles.Find(Name, Nothing) | DGN and attached-library styles enumerated; library styles accepted without modifying shared definitions |
| ModelReference.IsReadOnly, IsLocked, IsActive, IsAttachment, Type, Is3D | Verified preflight members |
| ModelReference.MasterUnit, UORsPerMasterUnit; MeasurementUnit.Base, UnitsPerBaseNumerator/Denominator | Ratio used to distinguish survey and international feet |
| ModelReference.GetGCS(Primary As Boolean); GeographicCoordinateSystem.Name, Source, DatumName, VerticalDatumName, Units | Live GCS read using GetGCS(True); no GCS writes |
| Element.SetXData(ApplicationName, Data() As XDatum), HasXData, GetXData | Exact signatures verified; persistent save/copy semantics not yet tested |
| ModelReference.GetElementByID(ID As DLong); DLongFromString; DLongToString; Element.DateLastModified | Verified for exact-ID journaling and edit warnings |
| DesignFile.DateCreated, FullName; ModelReference.Name; ModelReference.Scan | Verified context and recovery inventory members |

C# interop requires get_MasterUnit() for its returned structure, and uses out
parameters for GetTotalTextSize. VBA uses its normal structure-returning property
and ByRef output variables. Windows PowerShell 7 does not expose
Marshal.GetActiveObject; the successful inspection used Windows PowerShell 5
with the installed assembly explicitly loaded. The machine's script execution
policy was left unchanged.

## Live text calibration finding

An existing local style produced a measured height of 3,000 for a nominal local
height of 5 after its style was reapplied in the current model. Normalizing the
prepared element with ScaleUniform produced a measured height of exactly 5 and
a width of about 19.6373 for "250.5". The shared source style height/width remained
unchanged. Both prepared element IDs remained zero, so neither was inserted.

The helper uses this physical measurement and normalization instead of assuming
that a style's Height property is its final model-space size. Style properties
can reflect paper-size conventions under annotation scaling. Ordinary text is
created; civil annotation is never created. The annotation flag itself is not
cleared through an invented API. Future annotation-scale propagation must be
checked in the acceptance tests.

## Source/target metadata observed

The target model's unit ratio was denominator 12,000 / numerator 39,370 metres per
master unit, equal to 1200/3937 metres: US survey feet, despite the "ft" label.
Its GCS name was MDOT-MS83/2011-EF, derived from MS83-E, datum NAD83/2011, vertical
datum NAVD88. This custom GCS was not assigned an assumed EPSG code.
Fifteen writable local levels and four named local text styles were enumerated.
No level or style names are hardcoded in the product.

Source XML omitted CRS and vertical units. The user confirmed E/N/Z and survey
feet; reading the target GCS does not independently prove source/target equivalence.
The final placement confirmation and known-control checks remain mandatory.

## Terrain automation decision

No documented VBA terrain-selection-to-LandXML-export operation was established
from the installed COM interfaces. The installed Civil components are separate
from the plain MicroStation VBA object model. A native Civil SDK integration would
need a separately verified add-in/SDK workflow. Version 1 intentionally requires
manual LandXML export of the already merged terrain and does not claim automated
terrain selection/export.

## Acceptance still required

Compile the actual UserForm/modules in ORD. Then exercise AddElement/Level/Rewrite,
ByLevel persistence, text origins and rotation, XData/journal persistence, exact-ID
deletion/rollback, copying, cancellation, save/reopen, file size and PDF plotting
in a disposable DGN. Interface inspection and unsaved-element tests do not establish
those behaviors.

## Unused library definitions

The active API collections include unused DGNLIB definitions. The original
local-only filters hid those entries; they have been removed from both
dropdowns and exact-name resolvers. On the MDOT workspace the collections
contained 1,300 levels and 184 named text styles, including the user-shown
DTM contour levels and Civil-Proposed Contours style. Selection does not
require existing graphics on those levels. Native placement handles the
selected library level as described in [Bentley's level-library documentation](https://bentleysystems.service-now.com/community?id=kb_article&sysparm_article=KB0108843).

Direct assignment of a locked library text style prevents size overrides.
CPStyle resolves this through TextStyles.Add(existingStyle, existingName),
caching just the selected definition locally before CPText applies element-local
overrides. A live font/emphasis/physical-height check passed and confirmed the
shared source style unchanged. Background DGN placement is insufficient to
validate library level assignment because the operation requires an active file.

## Native Windows browsing and worker launch

CPBrowse uses Unicode GetOpenFileNameW with an ORD-owned modal dialog, existing-
file validation and XML/EXE/CPJ filters. It needs no Office or legacy CommonDialog
COM component. CPPath removes one surrounding Copy-as-path quote pair and rejects
unmatched quotes/control characters before JSON or command-line serialization.

CPRunEngine uses CreateProcessW with CREATE_NO_WINDOW, polls its process handle
without blocking UI events, retains cooperative cancel-file signaling and checks
the actual process exit code. Process/thread handles close on both success and
failure. Native structure sizes are checked for the supported 64-bit ORD host
before invoking the APIs. Error messages come from worker error.txt (UTF-16);
tracebacks remain in error.json. See [Microsoft process creation flags](https://learn.microsoft.com/en-us/windows/win32/procthread/process-creation-flags)
and [GetOpenFileNameW](https://learn.microsoft.com/en-us/windows/win32/api/commdlg/nf-commdlg-getopenfilenamew).

## Class-specific text sizes and placement diagnosis

CP2 stores both nominal physical heights. CPMeasuredWidths measures each exact
elevation string at its class height and tracks each class's box height; CPImport
checks class height and footprint before writing. Legacy CPJ1 membership remains
compatible for regeneration. A read-only comparison of the user's placed CP1
output, Default-3D terrain bounds, global origin and reference transforms showed
that E/N selection transposed this N/E/Z export; no native-writer transform was
introduced. The corrected run uses ne.
