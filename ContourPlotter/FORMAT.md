# Version 1 request and CP2 streaming contract

## Request (UTF-8 JSON; BOM accepted)

The worker accepts --request request.json. The VBA writes an ASCII-escaped UTF-8
JSON request in a fresh UUID run directory. Example:

~~~json
{
  "version": 1,
  "action": "generate",
  "run_id": "d6f938b3-a3ca-4ae4-bf68-cb01e55ba127",
  "input": "C:\\CAD_Data\\MergedTerrain.xml",
  "output_dir": "C:\\CAD_Data\\contour-job",
  "spatial": {
    "surface_index": 0,
    "axis_order": "xy",
    "horizontal_unit": "us-ft",
    "vertical_unit": "us-ft",
    "model_unit": "us-ft",
    "crs": "Verified project grid",
    "vertical_datum": "Verified project datum"
  },
  "settings": {
    "minor": 1,
    "major": 5,
    "base": 0,
    "label_mode": "major",
    "label_spacing": 250,
    "scale": 50,
    "paper_height_major": 0.1,
    "paper_height_minor": 0.1,
    "clearance_paper": 0.02,
    "break_lines": true,
    "dimension": 2,
    "max_labels": 100,
    "text_widths": {"250": 12.0},
    "text_box_height": 5.0,
    "text_box_height_minor": 5.0
  }
}
~~~

Actions: inspect (source inventory), analyze (strict source contract plus exact
elevation strings), generate (complete output). analyze writes levels.tsv and a
UTF-16 analysis.tsv row for the VBA. generate can include expected_source_sha256
at the root to bind output to the analyzed source. A run-local cancel file requests
cooperative cancellation. Exit 0 means success, exit 2 means an error; error.json
records diagnostics. progress.json is updated between stages.

## Output

result.json stores version/run ID, UTC timestamp, source path/hash/surface/index,
declared and confirmed spatial metadata, settings, warnings and timings/counts.
It is atomically replaced after a validated geometry file is published. Its
statistics contain that geometry file's SHA-256. result.json and geometry.cp1 must
have matching run IDs; do not consume a stale result from a failed job.

geometry.cp1 is **ASCII, tab separated, CRLF terminated, without a BOM**.
CRLF is required for native VBA Line Input. Numbers use invariant decimal points
and optional scientific notation, with 17 significant digits for coordinates.
No free-text source paths/names occur in CP2, avoiding quoting ambiguity.

Header, 14 fields:

~~~text
CP2  run_uuid  model_unit  vertical_unit  line_count  label_count  vertex_count  major_text_height  minor_text_height  xmin  ymin  xmax  ymax  dimension
~~~

Line record, followed immediately by n vertex records:

~~~text
C  contour_id  source_elevation  M_or_m  n
V  x  y  z
...
~~~

A connected source contour can have several C records after gaps or vertex-cap
splitting. IDs preserve its source path identity. M means major; m means minor.
Each C has 2..1024 points. Adjacent chunks repeat their shared endpoint. Coordinates
are target model master units. Elevation values and label strings use source
vertical units. In 2D all Z values are zero; in 3D Z is source elevation converted
to master units. No projection/datum transformation occurs.

Label record, 11 fields:

~~~text
L  contour_id  source_elevation  M_or_m  x  y  z  radians_ccw  physical_height  measured_width  elevation_text
~~~

Insertion is centered, with rotation about global Z. The numeric text must equal
the contour elevation. Physical width/height are model units after font/annotation
calibration; the element's local style Height can reflect its annotation convention.
Each label height must exactly match the major/minor height declared in the header.

Footer:

~~~text
E  line_count  label_count  vertex_count
~~~

Footer must exactly match header and actual records, followed by EOF. Empty runs
are valid with zero counts. Records are bounded at 4096 characters; absolute
numeric values cannot exceed 1e12 and counts cannot exceed 100 million. These
are format safety limits, not performance promises.

The Python validator independently reads the complete temporary file before atomic
publication. VBA holds the file open with a deny-write lock, validates every record,
then seeks to its beginning for insertion. It verifies run ID, units, requested
dimension, physical text sizes, finite numeric grammar, XY bounds, Z/elevation
agreement, classifications, counts and footer. No entire terrain/stream is loaded
into VBA memory.

## Insertion journal (CPJ1)

elements.cpj is UTF-16 with a BOM. It contains CPJ1 + tab + run UUID, then the exact
DGN path, model name and invariant numeric file-creation timestamp on separate
lines. Each member is ID + tab + DLong decimal ID + tab + invariant modified-date
serial. The final record is COMMITTED + tab + member count. The stream is flushed
every 128 members and on normal completion.

Element XData application name is OpenRoadsContourPlotter_v1. Its string is
CP1|UUID. IDs and tags must agree; membership is never inferred from level names.
Copying changes IDs, so copied graphics are outside the original journal.
Tags alone are not proof of membership, and a missing/uncommitted journal does
not authorize automatic broad deletion. See the recovery instructions.

CP2 adds explicit major/minor heights; older CP1 streams are rejected before
insertion. CPJ1 journals and CP1|UUID run tags remain unchanged so staged
regeneration can replace earlier runs by exact recorded membership. analysis.tsv
now has 14 fields: major/minor model heights are indexes 10/11, vertical unit 12
and source SHA-256 13. levels.tsv rows contain elevation text plus M/m class.
Legacy paper_height remains a common CLI/request fallback; explicit
paper_height_major and paper_height_minor override it.
