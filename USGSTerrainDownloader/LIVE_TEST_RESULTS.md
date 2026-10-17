# Real USGS download and processing pilot — 2026-10-17

**Passed without ORD.** The actual Python worker searched USGS, downloaded two
real 1-meter DEM tiles, verified their source hashes, mosaicked them, reprojected
them, clipped the original test polygon and wrote two GeoTIFFs with independently
checked horizontal/vertical unit tags. No DGN was read or modified.

This was an opt-in, bounded pilot, separate from the offline test suite. It used
the installed Mac QGIS 4.2.2 Python 3.12 runtime (GDAL 3.13.3, pyproj 3.7.2), with
TLS verification enabled and PROJ grid network access disabled.

## Inputs and evidence

### Additional Windows ORD metadata-only gate

Later on 2026-10-17, the saved ORD 26.00.00.154 VBA project called the new
standalone `launch_worker.exe` from its actual dialog. The user selected an
eight-vertex test shape in the 2D test DGN, entered `6576` (EPSG:6576), retained
US survey feet and explicitly confirmed their unmodified grid mapping. Search
displayed **2 one-metre DEM candidates**. Read-back of request.json, status.json
and search.json confirmed `search_complete`, the two products, both coordinate
confirmation flags true, no vertical declaration and empty stdout/stderr logs.
No Download/Process was requested. This is an actual ORD-to-USGS metadata-search
pass, not a Windows raster-processing or native-terrain-import pass.

Machine-local evidence: `.runs/ord-20261017/live-standalone-search.json` and the
user's test-DGN output folder. The complete Windows suite passed 37 offline tests,
including five real-EXE launcher tests; native form and invalid-schema bridge
checks passed before this search. No ProjectWise/workspace/autoload or PowerShell
policy settings were changed. The original Mac download/processing pilot follows.

### Original pilot inputs

A deliberately specified **headless test grid**, not a detected DGN, was used:
NAD83(2011) / Mississippi East (ftUS), EPSG:6507, master units `us-ft`, no local
coordinate modifiers. A 400-by-400-US-survey-foot square was centered on grid
XY `(939660.4243427969, 258955.18303075252)`. This was obtained from the explicit
WGS84 test location longitude -88.9745, latitude 30.212. The exact four vertices,
confirmed test-grid declarations and search bbox are saved in `search/request.json`
and `search/search.json` below.

Search used the live TNM `Digital Elevation Model (DEM) 1 meter` dataset and actual
returned download URLs. These two selected tiles came from the same collection
and publication; last-selected-row priority was explicitly accepted for their
12-meter overlap collars. Publication date was not used as an accuracy ranking.

| Product | USGS catalog record | Actual download bytes |
| --- | --- | ---: |
| USGS one meter x30y335 LA SoTerrebonne-GI 2015 | [5eacea0182cefae35a249a01](https://www.sciencebase.gov/catalog/item/5eacea0182cefae35a249a01) | 15,226,265 |
| USGS one meter x31y335 LA SoTerrebonne-GI 2015 | [5eaceac282cefae35a249edb](https://www.sciencebase.gov/catalog/item/5eaceac282cefae35a249edb) | 20,529,261 |
| **Total** | **50 MiB hard cap** | **35,755,526 (about 34.10 MiB)** |

The saved TNM descriptions for both products explicitly declare elevation values
in metres referenced to NAVD88. Both actual rasters declare horizontal EPSG:26916,
NAD83 / UTM zone 16N, approximately 1-metre pixels and a 10012-by-10012 grid. Their
band-unit tags are empty; source vertical units/datum were therefore confirmed
from the saved per-product catalog descriptions, not inferred from coordinate
magnitudes or empty raster tags. The source declarations and evidence citations
are retained in each request/report. No vertical datum shift was performed.

Actual sizes differ slightly from catalog estimates. HTTP transfer length,
readability and independent SHA-256 checks passed; catalog size is only an estimate.
Original source hashes:

```text
x30y335: e3fd6621da0971e6cfbd5e2825f64288953ddfa3fabe928d828f39dcb65de952
x31y335: 2377abca300c8d210b3b1d30dcd4fad858e5b7b4bed8f89e27b3787889f8ee8f
```

## Verified outputs

| Check | Result |
| --- | --- |
| Search and actual HTTPS file transfers | Passed; no mocked download responses |
| Original file retention / source SHA-256 recheck | Passed for both files |
| Adjacent tile compatibility and mosaic | Passed; equal horizontal CRS/resolution/aligned grid |
| Horizontal reprojection | EPSG:26916 → EPSG:6507, with the explicit NADCON5 grid pipeline recorded in reports |
| Output pixel spacing | 3.2808333333333333 US survey ft (1 metre) |
| Output grid | 122 × 122 cells; same geotransform in both unit variants |
| Original polygon clipping | Applied at target pixel centers after horizontal resampling |
| Coverage | 14,884 valid boundary cells; **0 gaps**, with `allow_gaps=false` |
| Metre elevation output | Approximately 0.196332 to 4.103634 m, NAVD88 |
| US-survey-foot elevation output | Approximately 0.644131 to 13.463340 ftUS, same datum |
| Entire-array elevation conversion comparison | Maximum difference from metre values × 3937/1200 was 1.78e-15 ftUS (floating-point rounding, not a survey accuracy measure) |
| Reopened GeoTIFF horizontal CRS / band unit / vertical-unit factor | Passed; EPSG:6507, explicit m/us-ft tags and 1200/3937 m vertical foot definition |
| Latest offline regression suite | **32 passed**, no skips with GIS runtime |

The NADCON5 operation's reported 0.2-metre transformation accuracy is retained in
reports. It is an operation metadata value, not a measurement of this DEM's
accuracy or a validation against survey control.

## Defect found and corrected by this live test

The first strict clipped run rejected 17 apparent boundary gaps. A diagnostic run
with clipping disabled had no gaps, and an independently generated fully covered
constant raster reproduced 21 artificial edge holes with the original code.
GDAL's source-grid cutline had removed resampling support at the reprojected
boundary. The processor now warps the covered neighborhood first and applies the
original polygon at target cell centers afterward. The strict zero-gap rule was
left enabled. Existing tests still reject genuine interior NoData and mask the
polygon exterior. The new synthetic regression fails before this change and
passes afterward.

The original failure and diagnostic outputs remain in the ignored run folder as
evidence; **use `verified-meters` or `verified-survey-feet` for the corrected outputs**.
The diagnostic run with `allow_gaps=true` is not the successful acceptance result.

## Local artifacts and repeatability

Local artifact directory (ignored by Git; DEMs are not included in the source):

```text
USGSTerrainDownloader/.runs/live-dem-20261017T060341Z/
    search/request.json
    search/search.json
    originals/                         # two sources and SHA-256 receipts
    process-meters/                    # preserved initial failed run
    diagnostic-gaps/                   # diagnostic only, gaps explicitly allowed
    diagnostic-no-cutline/             # diagnostic only, clipping disabled
    verified-meters/request.json
    verified-meters/terrain.tif
    verified-meters/report.json
    verified-survey-feet/request.json
    verified-survey-feet/terrain.tif
    verified-survey-feet/report.json
    verification.json                  # independent reopened-array/hash checks
```

To repeat on this Mac, create a new run directory, copy the successful requests,
update their `search_manifest` paths, and use the existing worker:

```sh
# In a configured GIS Python environment, from the application directory:
python python/main.py /absolute/new-run/search/request.json
# Review returned products and evidence; retain a 50 MiB cap in the process request.
python python/main.py /absolute/new-run/verified-meters/request.json
python python/main.py /absolute/new-run/verified-survey-feet/request.json
```

The first command obtains fresh products; select current IDs from that search
rather than assuming the IDs or coverage above remain available. Use a new folder
for each stage so successful outputs are never overwritten. The two processing
stages reuse verified originals when they share a destination parent. Refer to
[REQUEST_FORMAT.md](REQUEST_FORMAT.md) for request fields. This Mac's runtime needs
its QGIS Python paths and GDAL/PROJ data configured; see the exact invocation below.
On Windows, use the [captured QGIS/OSGeo4W child environment](README.md#windows-gis-setup-no-administrator-rights-needed-by-this-application).

```sh
SSL_CERT_FILE=/etc/ssl/cert.pem \
PROJ_DATA=/Applications/QGIS-final-4_2_2.app/Contents/Resources/qgis/proj \
GDAL_DATA=/Applications/QGIS-final-4_2_2.app/Contents/Resources/qgis/gdal \
PYTHONPATH=/Applications/QGIS-final-4_2_2.app/Contents/Resources/python3.12:/Applications/QGIS-final-4_2_2.app/Contents/Resources/python3.12/lib-dynload:/Applications/QGIS-final-4_2_2.app/Contents/Resources/python3.12/site-packages \
DYLD_LIBRARY_PATH=/Applications/QGIS-final-4_2_2.app/Contents/Frameworks \
/Applications/QGIS-final-4_2_2.app/Contents/MacOS/python3.12 \
python/main.py /absolute/path/to/new-run/request.json
```

## Limits of this result

This proves the headless download/processing path for this two-tile collection,
not every USGS dataset/format. LiDAR, the 1/3-arc-second processing path, Windows
GDAL runtime, ORD VBA compilation, graphical selection, automatic DGN GCS extraction,
raster attachment and native terrain import were not exercised by this pilot.
No comparison with survey control was performed, and no survey-grade claim is made.
