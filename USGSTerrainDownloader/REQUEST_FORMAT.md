# Worker request format (schema 1)

The VBA form creates these files automatically. Paths below are examples to adapt,
not a source-CRS declaration for your project. Each request belongs in a newly
created, writable run folder. Do not reuse a completed output folder.

## Search

```json
{
  "schema_version": 1,
  "action": "search",
  "source_crs": "EPSG:26916",
  "master_unit": "m",
  "crs_confirmed": false,
  "coordinates_are_grid": false,
  "boundary": [[300000,3700000],[300020,3700000],[300020,3700020],[300000,3700020]],
  "dataset": "dem1m",
  "python_executable": "C:\\OSGeo4W\\bin\\python.exe",
  "environment_file": "C:\\terrain-config\\runtime-environment.json"
}
```

The boundary above is **synthetic**. Both confirmation flags must remain false
until the actual project mapping has been verified. `boundary` is XY in DGN
master units; `source_crs` is the native projected grid definition. Accepted unit
keys are `m`, `ft` (international), `us-ft`. The ring can be open or explicitly
closed; at least three distinct vertices are required. Source horizontal CRS must
be two-dimensional and projected, with conventional east/north axes. Dataset keys:
`dem1m`, `dem13`, `lidar`.

Search writes `search.json`, `products.tsv`, `extent.txt`, `status.json` and
`status.txt` beside the request. Products start unselected. Keep `search.json`
unchanged: its full raw catalog metadata is the source-selection record.

## Process

Copy the same confirmed source CRS, master unit, boundary and dataset into a
new request folder. Change `action` to `process`, and add:

```json
{
  "search_manifest": "C:\\terrain-output\\run-search\\search.json",
  "selected_ids": ["actual-source-id-from-search"],
  "max_download_bytes": 5368709120,
  "unknown_size_confirmed": false,
  "source_vertical": {
    "unit": "m",
    "crs": "EPSG:5703",
    "evidence": "REPLACE with verified source-specific datum/unit metadata citation"
  },
  "resolution": 1,
  "elevation_unit": "m",
  "clip": true,
  "overlap_order_confirmed": false,
  "allow_gaps": false
}
```

This fragment is **not** a stand-alone request. `EPSG:5703` means NAVD88 height,
and must only be used with actual evidence for the selected sources. No default
vertical datum is assigned by the form. Resolution is in target/source projected
CRS native **horizontal** units; elevation units remain independent. The confirmed
source vertical declaration can be replaced per product using
`vertical_confirmations: {"product-id": {"unit": "m", "crs": "EPSG:5703",
"evidence": "source-specific citation"}}`. Each datum must agree; this interface
cannot perform vertical datum transformations. Mixed source elevation units are
rejected even with separate declarations.

Selected products retain their order in the search manifest; the last selected
row has priority when overlaps are explicitly accepted. The ring/CRS fingerprint
and dataset must match the search. Changing them requires a new search. The hard
byte cap covers the total actual sizes of selected originals, including cached
files. File names derive from URL hashes, never product titles.

For `lidar`, processing fields for DEMs are ignored. Optional
`pdal_executable: "C:\\OSGeo4W\\bin\\pdal.exe"` enables `pdal info --metadata
--schema`; otherwise completion is explicitly `lidar_downloaded_uninspected`.
Neither state means a validated terrain raster was created.

## Status and cancellation

The worker updates `status.json` with `state` and `message`. Terminal states are
`search_complete`, `dem_complete`, `lidar_downloaded_inspected`,
`lidar_downloaded_uninspected`, `failed`, `cancelled`. Exit 0 means the requested
stage completed; exit 2 means failure/cancellation. Stage success alone does not
mean a native ORD terrain model or survey-control validation occurred.

Create `cancel.flag` beside the request to request cooperative cancellation.
Network operations have 30-second socket timeouts; some GDAL/PDAL/read-back
operations finish before another checkpoint. PDAL inspection has a 180-second
subprocess timeout. Logs and completed originals are preserved. API calls have
bounded retries; an interrupted file transfer restarts on a subsequent request.
