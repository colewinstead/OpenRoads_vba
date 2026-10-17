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

The ORD-computer continuation inspected the installed ORD **26.00.00.154**
`MicroStationVBA.chm` and read-only reflection of
`Assemblies\Bentley.Interop.MicroStationDGN.dll` (file version **26.00.00.974**).
The installed CHM's SHA-256 is
`d239934fad2871108228e87fa1050dd907e228910f404f86b215780c16cb6873`.
An existing extracted copy in `%TEMP%\ORD2026-PLG-VBA-Help\vba.chm` has the same
hash; its `html/` topics were read. Fresh `hh.exe -decompile` attempts produced
no extracted files. This is actual installed-reference evidence, **not a VBA
compile or graphical acceptance pass**. See [current Windows evidence](TEST_PLAN.md#windows-ord-computer-preflight-2026-10-17).

| Bentley member | Exact use / evidence | Execution status here |
| --- | --- | --- |
| `CommandState.StartPrimitive EventHandlers, [WantKeyins]`; `IPrimitiveCommandEvents` | Installed CHM `usmthstartprimitive.htm` / `usobjiprimitivecommandevents.htm` and interop confirm the optional Boolean and all six callback signatures | User confirmed 2D rectangle capture/completion/reset |
| `CommandState.StartLocate EventHandlers`; `ILocateCommandEvents` | Installed CHM `usobjilocatecommandevents.htm` and interop confirm all seven callback signatures | User confirmed existing-shape acceptance; numeric/rejection cases pending |
| `CommandState.CreateLocateCriteria(False)`, `SetLocateCriteria` | Same conservative graphical locate pattern | ORD pending |
| `EnableAccuSnap`, `StartDynamics`, `StartDefaultCommand` | Native snapping, cursor callbacks, command cleanup | User confirmed visible preview and reset/outline cleanup; snapping accuracy untested |
| `CreateLineElement1(Nothing, Point3d())`, `Element.Redraw(DrawMode)` | Closed line-string preview, never added to model | Visible preview confirmed by user; no preview-code change; earlier absence unexplained |
| `Element.AsVertexList.GetVertices` | Installed `usproasvertexlist.htm` / `usmthgetvertices.htm`: returns zero-based `Point3d()`; no range/corner substitution | Native 2D in-memory shape round-trip passed; graphical locate gate pending |
| `Element.Type = msdElementTypeShape` | Reject complex/curved shapes, holes, references and cells | ORD pending |
| `ModelReference.IsActive`, `IsAttachment`, `Type`, `Name`, `UORsPerMasterUnit`, `MasterUnit`; `MeasurementUnit.Label` | Installed interop and CHM confirm model properties. `MeasurementUnit` is a value type with System/Base/unit-ratio/Label fields, matching the existing assignment | Signature verified; full unit/control gate pending |
| `Point3dFromXYZ` | XY rectangle at first pick's Z | ORD pending |
| `ActiveDesignFile.Path` | Candidate output directory only; no DGN writes | ORD pending |
| `ModelReference.GetGCS(Primary As Boolean)` returning `GeographicCoordinateSystem`; `Name` | Installed `usmthgetgcs.htm` says always pass True; interop confirms return type and read-only String Name | Raw COM plus actual application CallByName/form display passed for TN83/2011F; no EPSG/unit/modifier/control validation implied |

Bentley's [Developing Code in VBA](https://bentleysystems.service-now.com/community?id=kb_article&sysparm_article=KB0110097)
describes command interfaces, supplied Point3d/model context, temporary Redraw versus
persistent AddElement, and modeless event patterns. Bentley's
[VBA documentation guide](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0012836)
explains F2/F1/Object Browser and the delivered CHM.
The `GetGCS(true)` call/`Name` test appears in a
[Bentley-hosted MicroStation 2026 COM example](https://bentleysystemstest.service-now.com/community?id=community_question&sys_id=2ffba6051b69c3d474968730604bcb8e).
The installed signature now corroborates that example on build 26.00.00.154.
The call remains optional, late-bound, catches failure, and never supplies an
assumed EPSG code. All source coordinates still require explicit user confirmation.
Reflection shows descriptive GCS properties, coordinate conversions and PaperScale;
it does not establish a complete EPSG/WKT/modifier extraction contract. The full
coordinate safeguard remains necessary. Test an assigned GCS and the actual form
before accepting the displayed name as an executed application-integration result.

Further actual COM reads in the 2D test DGN returned Source=`EPSG:6576`,
DatumName=`NAD83/2011`, Units=`Foot`, VerticalDatumName=`NAVD88` and PaperScale=1.
The local pyproj EPSG catalog identifies 6576 as NAD83(2011) / Tennessee (ftUS).
The form now suggests that plain Source EPSG reference, with strict prefix/digit
validation; unavailable or ambiguous sources remain blank. Native parser/default
assertions passed, including EPSG:6576 and an unchecked grid-verification checkbox.
The field stays editable. This does not extract WKT or certify local transforms.
The user confirmed unmodified State Plane grid coordinates; a numeric known-point
comparison remains untested. Source is descriptive metadata, not a verified full
definition/export contract, and the DGN's vertical datum is not source-tile evidence.

Read-only metadata also established `DgnGCS.FromModel(DgnModelRef, Boolean)` and
`BaseGCS.GetWellKnownText(out String, WellKnownTextFlavor, Boolean)`,
`ToJson(out String)`, `GetEPSGCode(Boolean)`, LocalTransformType/Parameters and
PaperScale in the installed GeoCoord2/DgnGeoCoord2 assemblies. The standalone
`tests/OrdGcsProbe.cs` compiled against the installed DLLs, including `ustation.dll`.
Both explicit full-path MDL load attempts failed; no report was produced, so these
managed export APIs have **not run**. Bentley's
[2026 add-in example](https://bentleysystems.service-now.com/community?id=kb_article&sysparm_article=KB0012768)
documents the platform references and standard add-in search paths. Those paths
were not changed; the exact cause of this load failure is unproven. The diagnostic
is not a dependency of the working VBA application.

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

The master-unit suggestion reads `MeasurementUnit.Base` and the
`UnitsPerBaseDenominator / UnitsPerBaseNumerator` metre factor. Only metre-base
definitions matching 1, 0.3048 or 1200/3937 within 1e-12 metres are recognized;
unknown/custom/unitless definitions stay unselected. This does not confirm grid
mapping or establish the horizontal CRS. Native compile and the extended regression
passed in ORD 26.00.00.154; the actual model definition resolved to `us-ft` while
the CRS/grid checkbox and vertical declarations remained unconfirmed.

The rectangle is model-XY aligned, not screen-aligned. Original polygon geometry
is used for clipping. Simple horizontal straight-sided shape elements are supported;
complex/curved boundaries, reference transforms and holes require a separately
validated extension. Z is irrelevant to horizontal search, but varying shape Z is
rejected to avoid ambiguous projected footprints.

**Fence:** Bentley documents `Fence.IsDefined`/`GetContents`, but fence **contents**
are not the fence boundary. No reliably inspected boundary extraction covering view
projection, clips and void fences was available. Fence selection is deliberately
absent. No fence state is modified.

## 2D selection and requested native 3D terrain workflow

Boundary selection supports a normal **2D** design model; it does not require a
3D file. The native self-check actually ran in 2D. Its initial test fixture used
nonzero Z and was corrected for dimensionality; the application was not changed
to relax its coordinate safeguards. Form initialization exposed a separate real
bug: ListBox Change fired before lblTotals existed. RefreshTotals now returns
until that control is created, and the same native test then passed.

The requested next workflow is to select in 2D, create/reuse the associated Civil
3D model, and import a native terrain there. Bentley's
[OpenRoads SIG FAQ](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0020382)
documents Default-3D creation by operations such as making terrain/profile active.
This is not established by adding an arbitrary 3D ModelReference.

Read-only metadata inspection of the installed assemblies found these public
signatures, without invoking Civil SDK methods or writing any DGN terrain:

- `Bentley.CifNET.GeometryModel.SDK.GeometricModel.CreateTerrainSurface(Bentley.TerrainModelNET.DTM)`.
- `Bentley.CifNET.GeometryModel.SDK.Edit.ConsensusConnectionExtensions.GetOrCreateGeometricModel(ConsensusConnection)`.
- `Bentley.TerrainModelNET.Formats.TerrainImporter.CreateImporter(String)`,
  `ImportTerrain(String)` and the LandXML importer.
- `Bentley.TerrainModelNET.DTM.CreateFromFile(String)`, `CreateFromXyzFile(String)`,
  `AddPoints`, `AddLinearFeature` and `ConvertUnits(Double, Double)`.

These signatures are investigation evidence, **not** verified GeoTIFF support,
transaction/rollback semantics, Civil 2D/3D pairing or active-terrain behavior.
The native integration remains unimplemented. It must preserve original clipping,
NoData holes and separate XY/Z units; simply triangulating every valid raster
point can bridge gaps and is not an acceptable substitute. A supported import
and Civil-model workflow still needs actual disposable-DGN execution. The user
requires no ProjectWise/workspace/search-path/autoload changes.

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
update dates, size, format, body and geographic bounding box. Missing acquisition
and horizontal declarations remain unknown. Vertical autofill recognizes only the
explicit full-body USGS sentence declaring all bare-earth elevations in metres
referenced to NAVD88, with a credential-free HTTPS metadata citation. Conflicting
datum/unit keywords or unsupported wording leave it unresolved. All selected
records must agree before the form fills/protects unit, EPSG:5703 and every source
citation. Selection changes clear stale declarations; unresolved selections allow
manual verification. The backend rejects contradictory detected declarations
before transfer. No vertical datum is inferred from the dataset name or DGN.
Actual ORD tests passed the two saved live-search rows and selection clear/restore.
Source rasters supply actual
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
  Float64 working raster followed by Float32 tiled/DEFLATE GeoTIFF delivery. No implicit
  vertical datum shift is allowed. Compatible mosaics require equal horizontal
  CRSs, resolution (relative comparison `1e-9`) and aligned grid origins (within
  `1e-6` pixel). These are numerical compatibility budgets, not project criteria.
- Valid warped elevation cells are scaled explicitly by
  `metres_per_source_Z_unit / metres_per_output_Z_unit`. NoData is excluded from
  that multiplication. Internal NoData is `-1e300`; delivered Float32 NoData is
  `-999999`. Conversion and coverage checks retain Float64 arithmetic until final
  storage rounding. Float32 overflow and valid elevations rounding to the delivered
  sentinel stop processing. The report records maximum storage rounding in output
  elevation units, and its elevation range describes stored samples. Both working
  and partial files are removed after failure/cancellation. Source
  scale/offset-packed bands and rotated sources are rejected, not misinterpreted.
- Three actual Windows Float64 outputs passed GDAL processing/read-back but the user
  reported ORD "Import failed" while the original Float32 tile imported. A separate
  Float32 copy with finite NoData imported successfully per the user, retaining the
  full compound CRS and US-survey-foot elevations. This isolates the storage pair,
  not which of Float64 versus the original extreme NoData alone caused rejection.
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

`VBA form -> UTF-8 request.json -> standalone launch_worker.exe -> Python -> atomic
status/results/report files -> VBA form`. No command interpolation, cmd.exe,
stdin credentials, HTTP service or user-installed web server is needed.
The C# launcher uses .NET Framework `ProcessStartInfo` with UseShellExecute=False,
CreateNoWindow=True and a quoted script/request argument pair. It accepts one
request path, checks the interpreter and environment allowlist, and applies GIS
paths only to its Python child. PROJ_NETWORK stays OFF; PYTHONIOENCODING is UTF-8.
It drains stdout/stderr asynchronously and propagates the worker exit code,
preserving existing worker status (including cancellation). Launcher failures write
UTF-16 status and UTF-8 diagnostics. The x64 Windows EXE is built from
`python/WorkerLauncher.cs` by the standard-library `build_launcher.py` using the
installed Framework compiler. No PowerShell execution/policy change is involved;
the historical PS1 remains available but is not used by VBA. UI status
files are UTF-16, TSV fields are sanitized, and full metadata remains in JSON.
Search additionally writes `cell-size.txt` for the 1-metre DEM collection using
`1 / horizontal_axis.unit_conversion_factor`; angular-resolution DEMs and LiDAR
leave it empty. The form fills blanks or updates its previous automatic value,
preserving manual choices. It uses the validated horizontal CRS rather than master
unit labels or elevation units. The shared UTF-16 reader returns empty at EOF;
native testing exposed and fixed a BOM-only file ReadAll error. Native defaults,
manual preservation and empty-file behavior passed without DGN/settings changes.
`DoEvents` plus short Sleep yields the ORD UI while inputs are disabled.
This cooperative model still needs in-ORD reentrancy/pan/close validation.

| Phase | Implemented | Remaining gate |
| --- | --- | --- |
| 1 review/architecture/API validation | Repository + live USGS docs/catalog/products; standalone launcher; installed ORD 26.00.00.154 CHM/type-library inspection and actual VBA compile | Full GCS export and native Civil API execution |
| 2 selection/CRS | User-confirmed rectangle/shape picking; native form/unit/default safeguards; explicit EPSG/master-unit contract | Numeric known-point/global-origin, unsupported-geometry, rotated-view/3D and full GCS tests |
| 3 search/download | Native form, metadata, paging, selection, caps, cancellation, receipts; actual ORD dialog search returned 2 candidates; earlier headless real DEM transfers/inspection passed | Live ORD download/processing and active-job UI cancellation |
| 4 DEM | Synthetic suite plus two real adjacent 1-m USGS tiles: mosaic, clip, horizontal warp, explicit Z conversion, GeoTIFF/report | Other product collections and ORD import/control validation |
| 5 LiDAR | Search/download; optional PDAL metadata/schema inspection | Real LAS/LAZ/PDAL execution; ground-only rasterization deferred |
| 6 native integration | Open output folder and import instructions | Raster attachment and Civil SDK terrain creation deferred |

Bentley's [Civil SDK developer guide](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0012549)
is the documented route to investigate native ORD operations. This implementation
has no verified native terrain-creation interface. Basic raster attachment and
civil terrain import are distinct operations; neither is claimed to run. No guessed
key-ins are sent. The first acceptance target is the audited GeoTIFF plus manual
ORD import, rather than a fabricated `.mvba` or untested native terrain object.
