# Architecture and scope

## Data flow

Manual ORD export of the already merged terrain -> standalone EXE -> original TIN
triangles -> Matplotlib triangular contours -> label layout and gaps -> validated
CP2 stream -> VBA preflight/confirmation -> ordinary active-model DGN elements.

The native DGN writer is the primary output path. No GDAL, QGIS, raster, DXF,
external reference, new terrain model or terrain merge is involved.

## Reuse and contour-library decision

The standalone vendor subtree retains the existing LandXML plugin's safe XML tree
builder, namespace helpers, metadata parser and creator detection. vendor/tin.py
is an extraction of core.py's TIN reader, with the GDAL import removed. Changes
are identified in THIRD_PARTY_NOTICES.md. complete_export.py was inspected but its
raster-based/GDAL processing path is not used. The plugin checkout is unchanged.

The inspected plugin reader did NOT filter invisible faces, despite the original
task's description. This extraction adds explicit i="1" exclusion and rejects
unknown flags. It validates IDs even on excluded faces. A surface must contain
one TIN Definition. Multiple surfaces require an explicit choice; duplicate names
require index selection. Only points used by visible triangles contribute to
bounds and elevation levels.

[Matplotlib triangular contouring](https://matplotlib.org/stable/gallery/images_contours_and_fields/tricontour_demo.html)
accepts supplied triangle connections. The implementation supplies those original
connections, subtracts a local XY origin internally, validates triangulation with
the native triangle finder, and reads TriContourSet.allsegs through public APIs.
No Delaunay reconstruction, smoothing or interpolation across absent faces occurs.
Contour elevations are batched in groups of 32 to bound native calls and provide
cancellation opportunities. Matplotlib supplies the contour mathematics and
segment connection, including open/closed paths and vertex cases.

[ContourPy](https://contourpy.readthedocs.io/en/latest/user_guide/calculate/x_and_y.html)
takes quadrilateral grids; quad_as_tri subdivides those quads and does not accept
the original arbitrary LandXML triangle connectivity. It is therefore unsuitable
for this topology-preserving path. [PyVista/VTK](https://docs.pyvista.org/api/utilities/_autosummary/pyvista.make_tri_mesh.html)
can represent explicit triangular meshes, but adds a substantially larger runtime.
It was reviewed as an alternative, not benchmarked or implemented. Matplotlib's
analytical tests and measured runtime support the initial choice.

## Engineering contract

Every run explicitly resolves axis order, horizontal/vertical units, target
master units, CRS/local grid and vertical datum. Horizontal and optional 3D Z unit
conversion use independent exact definitions: m=1, ft=0.3048 m, US survey
foot=1200/3937 m. The DGN ratio, not its potentially ambiguous "ft" label, identifies
the model's unit. No datum or coordinate-system transformation is performed.
Source SHA-256 is checked across parsing and against the analysis used to measure
label strings. Unknown or contradictory units fail before insertion.

Custom Bentley GCS names are displayed verbatim; no guessed EPSG mapping is made.
A matching GCS name alone does not certify grid/ground scale, local modifiers,
vertical datum or source provenance. The final explicit placement confirmation
must be checked against control. A source CRS declaration and the confirmed grid
are both shown. Source/target grids that do not actually agree require an external
validated transformation, not a confirmation that waves away a mismatch.

Degenerate, duplicate and nonmanifold triangles and duplicate used XY vertices are
rejected. Native triangulation validation rejects invalid overlapping topology.
Flat surfaces produce no arbitrary plateau contours. Internal absent faces and
excluded triangles define holes. Explicit engineering boundary semantics are not
reconstructed; export the correctly clipped visible triangulation from ORD.
Very short paths are counted as flags, not removed. Coordinate-changing cleanup
is absent.

## CAD label layout

Labels are prioritized by major/minor class, distributed at approximately the
requested spacing, and limited per path. Candidates require sufficient length and
a direction change no greater than 15 degrees across the footprint. This angle is
a conservative presentation choice, not an agency criterion. Readable rotations
remain within +/-90 degrees. Rotated label envelopes, enlarged by clearance, are
checked in a spatial hash; conservative rectangular overlap can reject otherwise
usable placements.

The VBA analyzes exact decimal elevation strings and measures them using the
chosen DGN text style. It changes only a prepared element's local style,
reapplies that style, and normalizes measured model-space height with ScaleUniform.
This is necessary because the active annotation scale affected text preparation
in the observed ORD session. The shared style and global active settings remain
unchanged. Width and height measurements are rechecked before insertion.
Annotation scale/plot changes after placement still require acceptance testing;
the VBA COM surface does not expose a verified annotation-flag removal operation.

Python receives measured widths and a conservative height. The CLI alone uses one
nominal height per character and warns that the font width is estimated. Associated
contour gaps are clipped along path distance, preserving all vertices outside the
gaps. There is no guarantee against existing DGN text/linework or labels in other
runs; those are not loaded as collision obstacles. No minimum label count is
forced into an unsafe location.

## Native writer and staged replacement

The form captures the active model object, file path, model name, units,
resolution and GCS. Preflight rejects attachments, locked/read-only and non-design
models, and rechecks identity during processing and insertion. Existing exact-name
local DGN levels must allow full element access; no definitions are created or
modified. Only the output element receives ByLevel or explicit override values.

A deny-write file lock covers both validation and insertion passes. Record sizes,
counts, IDs, classifications, coordinates, elevations, dimensions, text footprints,
footer and EOF are checked. Each line string has at most 1,024 vertices; longer
paths are split with a repeated endpoint. This operational cap was exercised with
a prepared native 1,024-point line; it is not a claim about the API's maximum.
Every output element is primary-class, ungrouped and tagged with a run UUID.
AddElement precedes Level assignment; one Rewrite applies level and symbology.

The new run is journaled before replacement cleanup. A failed new insertion attempts
to remove only its own element handles, in reverse order. An application crash can
leave up to the unflushed journal batch (128 entries) requiring manual tag-based
reconciliation. XData is not a security boundary and can be edited/removed by a
user or stripped by other tools. Journals are essential.

Regeneration checks the old journal's file path, model and file creation timestamp.
Each old ID must still carry the old run tag. Timestamp changes are reported; all
manual edits are covered by a separate explicit deletion confirmation. New copied
IDs are outside the journal and preserved. Missing IDs are tolerated where the
host returns Nothing. No level-based deletion occurs. A cleanup failure keeps the
complete new output and reports partial old removal. Full atomicity, persistent
tag behavior after save/copy, and one-step Undo are not claimed.

## Performance boundaries

Python currently retains the safe XML tree, TIN arrays, contours and label pieces.
VBA streams individual records and retains new handles for rollback. Parsing,
native extraction, labeling and serialization are timed in result.json. The
importer reports separate validation/insertion times; save and PDF timings must
be recorded manually in ORD. Library calls are cancellable between batches, not
mid-call. Benchmarks use synthetic sources and a local private export, not an
unverified claim that resulting DGNs plot faster.
