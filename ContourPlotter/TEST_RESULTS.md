# Implementation validation results

## Automated Python tests

35 unittest cases passed on Windows, Python 3.14.0, NumPy 2.3.4 and Matplotlib
3.10.8. They cover the attributed OpenRoads fixture, required spatial metadata,
axis and exact foot conversion, missing/conflicting/unsupported units,
multiple/duplicate surfaces, invalid IDs/nonfinite data, invisible faces,
unused-point elevation bounds, degenerates/duplicates/non-TIN definitions,
unsafe XML, known triangle intersections, shared-edge connection, large State
Plane offsets, closed contours, holes, flat/exact-vertex cases, decimal/base
classification, cancellation, measured label footprints, rotations, sharp bends,
collision priority, physical gaps, long-line splitting, 3D Z conversion, source
hash mismatch, source overwrite protection, CRLF format, corrupt/truncated output and atomic failed jobs.

Commands and build/run instructions are in README.md. Test/runtime logs stay under
ignored jobs/; no private source data is used by the distributable unit suite.

## Performance

See examples/benchmark.json for repeatable raw measurements from synthetic grids
up through 524,288 triangles. Times include standalone source parsing, contouring,
labels and validated file serialization, and exclude ORD insertion, save and print.

A locally supplied OpenRoads export contained 53,692 points and 106,503 triangles.
A raw-coordinate 1/5-elevation extraction produced 1,141 connected contour paths.
The measured parse plus contour time was approximately 1.68 seconds. All 199,767
segment midpoints were located in original triangles. Independent barycentric
elevation checks had maximum residual approximately 1.31e-7 source elevation
units. This is a numerical check, not survey-control certification.

The private source omits CRS and vertical units. The user confirmed E/N/Z and US
survey feet, and requested the open DGN metadata be read through Bentley's API.
The target's custom GCS/datum and unit ratio were read live; this does not
independently establish the source's grid/datum equivalence. Private inputs and
generated coordinate data remain ignored and excluded from packaging.

## Live ORD API probes

Read-only calls to the open ORD 26.00.00.154 session confirmed GCS and exact master
unit ratio, 15 writable local levels, and four named local text styles. An
uninserted 1,024-vertex line string was prepared and given ByLevel symbology.
Uninserted text using an existing style was measured and normalized to exactly
five model units physical height. The shared style's height/width did not change.
Both element IDs were zero. See API_NOTES.md and tests/ReadOnlyOrdProbe.cs.

The user-created vba/ContourPlotter.mvba was assembled through the running ORD
VBA interface, with all six modules and the runtime UserForm. It compiled in
ORD 26.00.00.154 with no broken references. Compilation corrected the design
model enum to msdModelTypeNormal. The journal's file-creation date comparison
was also corrected to use the same invariant numeric format as its writer.

A temporary integration probe initialized the actual form, inspected the private
Merged surface, measured its elevation strings using the local Civil_Labeler
text style, and completed Analyze / Preview through the packaged EXE. The
preview produced 1,616 native line-string pieces and 410 labels. Those counts
differ from the standalone estimate because the VBA run supplied measured
DGN text widths. The target's grid/datum were used for this preview; source
equivalence still requires control verification. The temporary probe was removed
from the saved MVBA.

**No graphics were added, rewritten or deleted in the active project DGN.** Native insertion,
ByLevel persistence, rollback, tag/journal save/reopen/copy behavior, regeneration,
undo, visual UI inspection and PDF printing remain untested.
ORD_TEST_PLAN.md is the required acceptance path.

## Packaging

The builder runs the unit suite, freezes the engine, checks the packaged version
and generates actual contour output from its packaged synthetic fixture.
The Windows x64 EXE and ZIP built successfully with PyInstaller 6.17.0.
Both packaged smoke tests passed. The packaged executable also processed the
private 106,503-triangle terrain: 1,141 contours, 1,602 line-string pieces,
395 labels and 199,758 output vertices in about 3.3 seconds (worker time).
Label widths in that standalone check used the documented conservative estimate;
the VBA workflow supplies actual measured widths. No DGN graphics were inserted.
A successful EXE build does not establish native DGN writer acceptance.

## Unused DGNLIB selection regression

The original selectors and resolvers excluded definitions marked as coming from
a library. The corrected compiled MVBA lists all 1,300 available levels and 184
named text styles. A live regression check resolved each of the four user-shown
major/minor contour and label levels without requiring graphics on them.

The selected Civil-Proposed Contours library style is copied by the native
TextStyles.Add API, with the same name, into the DGN when label measurement
needs editable element-local overrides. No DGNLIB definition is modified. A
live uninserted label retained the selected font/emphasis, measured exactly
five model units high, and had element ID zero. Shared source height/width were
unchanged. Dropdown population does not copy definitions; selecting/measurement
may cache the chosen text style.

A background disposable-file probe could not validate library-level placement:
Bentley requires an active design file for that operation. Full placement and
save/reopen acceptance remain on ORD_TEST_PLAN.md; no successful native writer
run is claimed from that probe. Temporary probe modules were removed from the
final MVBA. Private disposable files remain under ignored jobs/.

## File paths and console-free launcher regression

Windows Copy-as-path values now lose one enclosing quote pair at the VBA path
boundary. A live VBA check passed for ordinary and quoted paths, spaces,
apostrophes and Unicode, JSON request normalization, and rejection of unmatched
quotes. Source, engine and prior-journal paths share that normalization.

A temporary console-application probe launched from the compiled MVBA reported
GetConsoleWindow() = 0. A nonzero-exit probe returned its expected readable error
message. CREATE_NO_WINDOW prevents creation of the console, while process handles
retain completion monitoring and cooperative cancellation. Worker failures now
write protected UTF-16 error.txt alongside error.json. The Python regression
checks readable error output and protects a source named error.txt from overwrite.

Browse now uses the native Unicode Windows file picker, with existing-file and
XML/EXE/CPJ filters, instead of the unavailable legacy CommonDialog component.
Temporary process/path test modules were removed from the final saved MVBA.

The user confirmed the native picker opened and selected Merged.xml; the returned
path matched the fixture and contained no surrounding quotes. The final compiled
form then inspected and previewed the real surface using deliberately quoted
source and engine paths through the console-free launcher. It found one surface
and produced 1,616 line-string pieces and 410 labels; the source field was
normalized to an unquoted path. No contour graphics were inserted.

## Placement diagnosis and separate label heights

Read-only SDK comparison established that the supplied Merged export stores
Northing/Easting/Z. The run selected xy (E/N/Z), and its first inserted native
vertex exactly matched the generated CP1 vertex; the writer introduced no
translation or scale change. Swapping the source axes matches the Default-3D
terrain/reference bounds within the DGN's 0.001-foot resolution. Global origin
was zero and both relevant reference transforms were identity. This resolves
the axis setting for this export; it is not independent survey-control certification.

Major and minor paper text heights are independently configurable. The engine,
measured ORD text widths, collision envelopes, gaps and CP2 writer use each
class's physical height. CP2 declares both heights and rejects a label whose
height differs from its class. New regressions cover differing gaps/footprints,
minor text larger than major text, end-to-end request/analysis propagation,
invalid sizes, and serialized class-height corruption. CPJ1 journals and old run
tags are unchanged for staged replacement of prior output.

The updated compiled ORD form completed a real-source N/E/Z preview with major
0.12-inch and minor 0.08-inch paper heights at 50 ft/in. CP2 recorded exactly
6-foot major and 4-foot minor label heights after native font measurement. The
preview produced 1,141 contours, 2,648 line pieces and 1,489 labels at the corrected
terrain bounds. No replacement graphics were inserted or prior output removed.
Both UI defaults remain 0.10 inch; these distinct values were a verification case.

The frozen engine CP2 smoke check passed in a fresh directory. A prior smoke
retry was blocked by OneDrive read-only progress.json; the builder now allocates
a unique smoke directory per run, avoiding stale outputs and file replacement.
