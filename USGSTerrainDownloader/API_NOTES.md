# API evidence, architecture and support boundary

Reviewed 2026-10-17. No Bentley Civil SDK API or native terrain-model creation is
called. This document distinguishes inspected sources from executed integration.

## Repository review and Bentley API status

Reviewed `PerpendicularLineGeneratorMain.bas`, `clsPerpendicularLocate.cls`,
`frmPerpendicularLines.code.vba`, `modPerpendicularGeometry.bas`, its API notes,
README/installation instructions and test plan, plus the repository README.
The new application reuses the established command-event, temporary dynamics,
UserForm/WithEvents, model-context validation and error-reporting patterns. It
imports none of that application's modules and changes none of its files.

The existing [API notes](../PerpendicularLineGenerator/API_NOTES.md) record
inspection of ORD 2026's `MicroStationVBA.chm` and delivered interop DLL. Those
artifacts were **not available for reinspection in this implementation session**.
The reachable Windows test PC has Python but no ORD installation. Existing
signature evidence plus Bentley's published documentation supports source authoring;
it does not establish this new application's compilation or runtime compatibility.

| Bentley member | Exact use / evidence | Execution status here |
| --- | --- | --- |
| `CommandState.StartPrimitive Object`; `IPrimitiveCommandEvents` | Start, DataPoint, Dynamics, Reset, Cleanup, Keyin signatures match the existing verified picker | ORD pending |
| `CommandState.StartLocate Object`; `ILocateCommandEvents` | Start, LocateFilter, Accept, LocateReset, LocateFailed, Dynamics, Cleanup match the existing picker | ORD pending |
| `CommandState.CreateLocateCriteria(False)`, `SetLocateCriteria` | Same conservative graphical locate pattern | ORD pending |
| `EnableAccuSnap`, `StartDynamics`, `StartDefaultCommand` | Native snapping, cursor callbacks, command cleanup | ORD pending |
| `CreateLineElement1(Nothing, Point3d())`, `Element.Redraw(DrawMode)` | Closed line-string preview, never added to model | ORD pending |
| `Element.AsVertexList.GetVertices` | Actual simple shape vertices, no range/corner substitution | ORD shape gate pending |
| `Element.Type = msdElementTypeShape` | Reject complex/curved shapes, holes, references and cells | ORD pending |
| `ModelReference.IsActive`, `IsAttachment`, `Type`, `Name`, `UORsPerMasterUnit`, `MasterUnit`; `MeasurementUnit.Label` | Active-model guards, unit label and model-resolution checks | ORD pending |
| `Point3dFromXYZ` | XY rectangle at first pick's Z | ORD pending |
| `ActiveDesignFile.Path` | Candidate output directory only; no DGN writes | ORD pending |
| `ModelReference.GetGCS(True)` returning `GeographicCoordinateSystem`; `Name` | **Best-effort late-bound probe** of Bentley COM GCS name; not a full definition/units/modifier extraction | Not verified in ORD 2026 here |

Bentley's [Developing Code in VBA](https://bentleysystems.service-now.com/community?id=kb_article&sysparm_article=KB0110097)
describes command interfaces, supplied Point3d/model context, temporary Redraw versus
persistent AddElement, and modeless event patterns. Bentley's
[VBA documentation guide](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0012836)
explains F2/F1/Object Browser and the delivered CHM.
The `GetGCS(true)` call/`Name` test appears in a
[Bentley-hosted MicroStation 2026 COM example](https://bentleysystemstest.service-now.com/community?id=community_question&sys_id=2ffba6051b69c3d474968730604bcb8e).
That example is corroborating evidence, **not** an independently inspected ORD
2026 VBA signature. Accordingly the call is optional, late-bound, catches failure,
and never supplies an assumed EPSG code. All source coordinates require explicit
user confirmation. Check `GetGCS`/`GeographicCoordinateSystem` in the installed
Object Browser before accepting this probe as supported on a particular build.

No `GetEPSG`, `ExportGCSWKT`, Civil COM terrain method, MDL pointer declaration,
key-in recording or C expression is invented. GCS automatic mapping remains an
integration limitation, rather than a misleading success flag. Bentley's
[unit-system explanation](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0110498)
distinguishes working units from projected GCS units; the application does too.

### Coordinates and selections

VBA Point3d/vertex values are treated as active-model **master-unit** coordinates,
as in the existing generator. The supplied geometric values are retained without
rounding. No extra UOR multiplication or speculative global-origin subtraction is
performed. Python multiplies by `metres_per_master_unit / metres_per_CRS_unit`
once to reach native projected coordinates. Working-unit labels are displayed but
never used to infer physical units. The required ORD known-point test must establish
that these coordinates match the actual model's grid, including global-origin
behavior. Local transforms, rotated/offset/ground-scaled GCS modifiers and dynamic
epochs are unsupported; do not confirm them as grid XY.

The rectangle is model-XY aligned, not screen-aligned. Original polygon geometry
is used for clipping. Simple horizontal straight-sided shape elements are supported;
complex/curved boundaries, reference transforms and holes require a separately
validated extension. Z is irrelevant to horizontal search, but varying shape Z is
rejected to avoid ambiguous projected footprints.

**Fence:** Bentley documents `Fence.IsDefined`/`GetContents`, but fence **contents**
are not the fence boundary. No reliably inspected boundary extraction covering view
projection, clips and void fences was available. Fence selection is deliberately
absent. No fence state is modified.

## USGS interface verified live

Primary references:

- [TNM Access API documentation](https://tnmaccess.nationalmap.gov/api/v1/docs)
- [USGS API FAQ](https://www.usgs.gov/faqs/there-api-accessing-national-map-data)
- [Live dataset catalog](https://tnmaccess.nationalmap.gov/api/v1/datasets)

The Swagger specification is embedded in the documentation HTML. Both its current
parameter definitions and actual `/datasets` and `/products` responses were read.
The client revalidates dataset tags against the live catalog at each search:

| Form choice | `datasets` / `sbDatasetTag` | `prodFormats` |
| --- | --- | --- |
| 1-meter DEM | `Digital Elevation Model (DEM) 1 meter` | `GeoTIFF` |
| 1/3-arc-second DEM | `National Elevation Dataset (NED) 1/3 arc-second` | `GeoTIFF` |
| LiDAR | `Lidar Point Cloud (LPC)` | `LAS,LAZ` |

Requests use `bbox=west,south,east,north` in WGS84 degrees, `outputFormat=JSON`,
`max=100` and `offset` advanced by raw item count, not deduplicated count. The
100-item page size meets the API's multiple-of-five pagination restriction.
Response envelope fields are `total`, `items`, `errors`; full item dictionaries
are preserved, including metadata links, source ID, product title, publication/
update dates, size, format, body and geographic bounding box. Missing acquisition,
horizontal/vertical declarations remain unknown; product descriptions are shown,
not converted into verified engineering assertions. Source rasters supply actual
spatial tags after download. Duplicate source IDs/URLs are collapsed. No newest-
first selection or accuracy ranking is applied. Product boxes can overselect an
irregular polygon; raster clipping and cell-center coverage are separate checks.

Direct TIFF/GeoTIFF URLs are selected from the documented URL metadata, and LAS/LAZ
from point-cloud URL fields; no example URLs are embedded in the client. Unsupported
archives/missing coverage/changed response structures cause an explicit stop. API
HTTP 429/500/502/503/504 and connection/timeouts get three bounded attempts with
cancellable short backoff; other HTTP errors propagate. File transfers restart on
user retry, while completed sources are reused by SHA-256 receipt verification.
TLS remains verified. Download estimates are not checksums: actual HTTP length
(where present), streamed count, local SHA-256 and GDAL read/metadata checks serve
different purposes. The local SHA-256 is not an authoritative USGS checksum.

## Geospatial contract

- [pyproj Transformer](https://pyproj4.github.io/pyproj/stable/api/transformer.html):
  `always_xy=True`, `allow_ballpark=False`, `only_best=True`, network off and
  `errcheck=True`. Missing best-operation grids fail instead of degrading silently.
  These flags do not establish survey accuracy. The executed PROJ pipeline and
  reported accuracy are recorded. Dynamic coordinate epochs are unsupported.
- Boundary edges are densified to <=100 metres in source projected space, capped
  at 20,000 vertices. All samples transform before geographic extrema are taken;
  the bounding box receives a `1e-7`-degree numerical padding. A separate test uses
  7-metre probes. This is bounded sampling for catalog discovery, not a geodetic
  error guarantee for arbitrary custom projections. CRS area-of-use bounds receive
  a 0.01-degree bounding-envelope allowance (their published envelopes are coarse).
- Geographic span is capped at 2 degrees, area at 10,000 km². Antimeridian areas,
  nonprojected DGN CRSs, unusual non-east/north axes, compound DGN CRSs and local
  transformations stop rather than being guessed.
- Metre = 1 m; international foot = 0.3048 m; US survey foot = 1200/3937 m.
  Horizontal native-unit conversion occurs once at the model boundary. Source
  elevation units and vertical datum are separately confirmed with an evidence
  citation. Contradictory embedded units/datums stop processing. Unknown vertical
  information cannot be silently assigned; the user must supply verified evidence.
- [GDAL Warp](https://gdal.org/en/stable/programs/gdalwarp.html): same confirmed
  target horizontal CRS as the DGN; an explicit pyproj-selected **2D** pipeline,
  `-novshift`, exact transform evaluation (`errorThreshold=0`), bilinear resampling,
  optional original-polygon mask applied at target cell centers **after** resampling,
  Float64, tiled/DEFLATE GeoTIFF. No implicit
  vertical datum shift is allowed. Compatible mosaics require equal horizontal
  CRSs, resolution (relative comparison `1e-9`) and aligned grid origins (within
  `1e-6` pixel). These are numerical compatibility budgets, not project criteria.
- Valid warped elevation cells are scaled explicitly by
  `metres_per_source_Z_unit / metres_per_output_Z_unit`. NoData is excluded from
  that multiplication. The chosen output NoData sentinel is `-1e300` in Float64;
  nonfinite valid samples/converted-sentinel collisions stop processing. Source
  scale/offset-packed bands and rotated sources are rejected, not misinterpreted.
- The live coastal pilot exposed false boundary holes when GDAL applied a source-grid
  cutline before reprojection. A fully covered synthetic constant surface reproduced
  the defect. Warping now retains the interpolation neighborhood; the original
  polygon clips the finished target grid. The unchanged gap check still rejects
  true interior NoData. See [live evidence](LIVE_TEST_RESULTS.md).
- Overlaps require explicit list-priority acceptance; the last selected row wins,
  as recorded in the report. Dates/resolutions are never silently used to select
  an alleged better surface. Gaps are independently counted by rasterizing the
  original polygon at output cell centers. NoData gaps fail unless explicitly
  accepted; subpixel voids cannot be certified absent by that check.
- Output vertical CRS retains the **confirmed source datum**, with its unit changed
  only to reflect the explicit numerical conversion. The application sets the
  vertical unit EPSG authority (9001/9002/9003) on GDAL's compound CRS before saving.
  [GDAL GeoTIFF writer source](https://github.com/OSGeo/gdal/blob/master/frmts/gtiff/gt_wkt_srs.cpp)
  uses that authority to encode `VerticalUnitsGeoKey`; synthetic read-back tests
  caught incorrect metre metadata before this was added. Complete output CRS,
  band units and readable pixels must pass read-back before `terrain.tif` is named.
  No NAVD88-to-NGVD29, ellipsoidal-to-orthometric, geoid or epoch operation exists.

## Process boundary and phases

`VBA form -> UTF-8 request.json -> PowerShell child launcher -> Python -> atomic
status/results/report files -> VBA form`. No command interpolation, cmd.exe,
stdin credentials, HTTP service or user-installed web server is needed.
PowerShell `ProcessStartInfo` applies GIS environment paths only to its Python
child and drains stdout/stderr asynchronously to avoid blocked pipes. UI status
files are UTF-16, TSV fields are sanitized, and full metadata remains in JSON.
`DoEvents` plus short Sleep yields the ORD UI while inputs are disabled.
This cooperative model still needs in-ORD reentrancy/pan/close validation.

| Phase | Implemented | Remaining gate |
| --- | --- | --- |
| 1 review/architecture/API validation | Repository + live USGS docs/catalog/products; Python/PowerShell contracts | Delivered ORD CHM/type-library reinspection |
| 2 selection/CRS | Rectangle dynamics, simple shape, GCS name probe, explicit EPSG/master-unit contract | Compile, picking, known-point/global-origin and GCS tests in ORD |
| 3 search/download | Native form, metadata, paging, selection, caps, cancellation, receipts; two real DEM transfers/inspection tested | Native form and field usability |
| 4 DEM | Synthetic suite plus two real adjacent 1-m USGS tiles: mosaic, clip, horizontal warp, explicit Z conversion, GeoTIFF/report | Other product collections and ORD import/control validation |
| 5 LiDAR | Search/download; optional PDAL metadata/schema inspection | Real LAS/LAZ/PDAL execution; ground-only rasterization deferred |
| 6 native integration | Open output folder and import instructions | Raster attachment and Civil SDK terrain creation deferred |

Bentley's [Civil SDK developer guide](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0012549)
is the documented route to investigate native ORD operations. This implementation
has no verified native terrain-creation interface. Basic raster attachment and
civil terrain import are distinct operations; neither is claimed to run. No guessed
key-ins are sent. The first acceptance target is the audited GeoTIFF plus manual
ORD import, rather than a fabricated `.mvba` or untested native terrain object.
