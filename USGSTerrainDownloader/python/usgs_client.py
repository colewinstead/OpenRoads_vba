"""TNM Access API and bounded, atomic HTTPS downloads; no third-party HTTP client."""
import hashlib
import json
import math
import re
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urlparse
from urllib.request import Request, urlopen

from crs_utils import require

BASE = "https://tnmaccess.nationalmap.gov/api/v1/"
DATASETS = {
    "dem1m": "Digital Elevation Model (DEM) 1 meter",
    "dem13": "National Elevation Dataset (NED) 1/3 arc-second",
    "lidar": "Lidar Point Cloud (LPC)",
}


def https_url(url):
    p = urlparse(url)
    require(p.scheme == "https" and p.hostname and not p.username and not p.password,
            "Only credential-free HTTPS URLs are allowed.")
    return url


def open_https(url):
    response = urlopen(Request(https_url(url), headers={"User-Agent": "ORD-USGSTerrainDownloader/0.1"}), timeout=30)
    try:
        https_url(response.geturl())
    except Exception:
        response.close()
        raise
    return response


class TNMClient:
    def __init__(self, opener=open_https, check=lambda: None, progress=lambda message: None):
        self.opener, self.check, self.progress = opener, check, progress

    def get_json(self, endpoint, params=None):
        url = BASE + endpoint + ("?" + urlencode(params) if params else "")
        for attempt in range(3):
            self.check()
            try:
                with self.opener(url) as response:
                    raw = response.read(16 * 1024 * 1024 + 1)
                    require(len(raw) <= 16 * 1024 * 1024, "API response exceeds 16 MiB.")
                    return json.loads(raw)
            except (HTTPError, URLError, TimeoutError, OSError) as error:
                if isinstance(error, HTTPError) and error.code not in (429, 500, 502, 503, 504):
                    raise
                if attempt == 2:
                    raise
                for _ in range(10 * (attempt + 1)):
                    self.check()
                    time.sleep(0.1)

    def search(self, kind, bbox, limit=500):
        require(kind in DATASETS, "Unknown dataset choice.")
        require(len(bbox) == 4 and -180 <= bbox[0] < bbox[2] <= 180 and
                -90 <= bbox[1] < bbox[3] <= 90, "Invalid WGS84 bounding box.")
        require(1 <= limit <= 2000, "Search limit must be 1 to 2,000 products.")
        catalog = self.get_json("datasets")
        tags = set()
        def collect(rows):
            require(isinstance(rows, list), "Malformed USGS dataset catalog.")
            for row in rows:
                tags.add(row.get("sbDatasetTag"))
                collect(row.get("tags", []))
        collect(catalog)
        require(DATASETS[kind] in tags, "Dataset identifier missing from live TNM catalog; update required.")
        offset, products, seen_ids, seen_urls, seen_pages = 0, [], set(), set(), set()
        while True:
            self.progress(f"Search: requesting results at offset {offset}")
            page = self.get_json("products", {"datasets": DATASETS[kind],
                "bbox": ",".join(format(v, ".10f") for v in bbox),
                "prodFormats": "LAS,LAZ" if kind == "lidar" else "GeoTIFF",
                "max": 100, "offset": offset, "outputFormat": "JSON"})
            require(not page.get("errors"), f"USGS API errors: {page.get('errors')}")
            rows, total = page.get("items"), page.get("total")
            require(isinstance(rows, list) and isinstance(total, int) and total >= 0,
                    "Malformed USGS product response.")
            require(total <= limit, f"USGS returned {total} products; limit {limit}. Draw a smaller boundary.")
            if not rows:
                require(offset >= total, "USGS pagination stopped before all products arrived. Retry search.")
                break
            signature = json.dumps(rows, sort_keys=True)
            require(signature not in seen_pages, "USGS repeated a page; search stopped.")
            seen_pages.add(signature)
            for item in rows:
                product = normalize(item, kind)
                if product["id"] in seen_ids or product["url"] in seen_urls:
                    continue
                seen_ids.add(product["id"])
                seen_urls.add(product["url"])
                b = product["bbox"]
                if b and not (b[2] <= bbox[0] or b[0] >= bbox[2] or b[3] <= bbox[1] or b[1] >= bbox[3]):
                    products.append(product)
            offset += len(rows)
            require(offset <= 2000, "Pagination safety limit exceeded.")
            if offset >= total:
                break
        return products


def vertical_metadata(item, kind):
    if kind == "lidar":
        return None
    text = " ".join(str(item.get("body") or "").split())
    statement = ("All bare earth elevation values are in meters and are referenced to the "
                 "North American Vertical Datum of 1988 (NAVD88).")
    # shortcut: recognize this explicit USGS declaration only; add other formats with verified fixtures.
    match = re.search(r"(?:^|[.!?]\s+)(" + re.escape(statement) + r")", text, re.IGNORECASE)
    if not match or re.search(r"\b(?:NGVD\s*29|EGM\s*(?:96|2008)|ellipsoid\w*|feet|foot)\b", text, re.IGNORECASE):
        return None
    source = str(item.get("metaUrl") or "")
    if not source:
        return None
    try:
        https_url(source)
    except ValueError:
        return None
    return {"unit": "m", "crs": "EPSG:5703",
            "evidence": f"USGS TNM source {item.get('sourceId', '')}; {source}; {match.group(1)}"}


def normalize(item, kind):
    urls = item.get("urls") or {}
    url = (item.get("downloadLazURL") or urls.get("LAZ") or urls.get("LAS") or item.get("downloadURL")) if kind == "lidar" else (
        item.get("downloadURLRaster") or urls.get("TIFF") or urls.get("GeoTIFF") or item.get("downloadURL"))
    https_url(url or "")
    suffix = Path(urlparse(url).path).suffix.lower()
    require(suffix in ((".las", ".laz") if kind == "lidar" else (".tif", ".tiff")),
            "USGS product is not a direct GeoTIFF/LAS/LAZ download; archives require manual handling.")
    b = item.get("boundingBox") or {}
    bbox = [float(b[k]) for k in ("minX", "minY", "maxX", "maxY")] if b else None
    require(bbox is not None and all(math.isfinite(v) for v in bbox)
            and bbox[0] < bbox[2] and bbox[1] < bbox[3], "Product has missing/invalid geographic coverage.")
    size = item.get("sizeInBytes")
    require(size is None or isinstance(size, (int, float)) and size >= 0, "Invalid product size.")
    vertical = vertical_metadata(item, kind)
    return {"id": str(item.get("sourceId") or hashlib.sha256(url.encode()).hexdigest()),
            "title": str(item.get("title") or "Untitled product"), "url": url,
            "bytes": int(size) if size else None, "bbox": bbox, "kind": kind,
            "resolution": {"dem1m": "1 meter", "dem13": "1/3 arc-second", "lidar": "point cloud"}[kind],
            "publication": item.get("publicationDate"), "updated": item.get("lastUpdated"),
            "acquisition": item.get("acquisitionDate"), "metadata_url": item.get("metaUrl"),
            "vertical": "Metadata: metres / NAVD88" if vertical else "Unverified: inspect source and product metadata before processing",
            "vertical_confirmation": vertical,
            "raw": item}


def _download(product, folder, remaining_bytes, check, progress, opener):
    """Restart interrupted transfer safely; only reuse files with a matching local SHA-256 receipt."""
    require(remaining_bytes > 0, "Download budget exhausted.")
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    suffix = Path(urlparse(https_url(product["url"])).path).suffix.lower()
    require(suffix in (".tif", ".tiff", ".las", ".laz"), "Unsupported source file type.")
    path = folder / (hashlib.sha256(product["url"].encode()).hexdigest()[:24] + suffix)
    receipt = path.with_suffix(path.suffix + ".json")
    if path.exists() and receipt.exists():
        saved = json.loads(receipt.read_text(encoding="utf-8"))
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024*1024), b""):
                check()
                digest.update(block)
        if saved.get("url") == product["url"] and saved.get("sha256") == digest.hexdigest():
            require(0 < path.stat().st_size <= remaining_bytes, "Cached source exceeds download budget.")
            return path, saved
    require(not path.exists(), f"Existing source lacks a matching receipt: {path}. Move it aside before retrying.")
    temp = path.with_suffix(path.suffix + ".part")
    digest, received = hashlib.sha256(), 0
    try:
        with opener(product["url"]) as response, temp.open("wb") as stream:
            length = response.headers.get("Content-Length")
            length = int(length) if length is not None else None
            if length is not None:
                require(0 < length <= remaining_bytes, "Server download size exceeds remaining budget.")
            while True:
                check()
                block = response.read(1024*1024)
                if not block:
                    break
                received += len(block)
                require(received <= remaining_bytes, "Actual transfer exceeds download budget.")
                stream.write(block)
                digest.update(block)
                progress(f"Download: {product['title']} — {received / 1048576:.1f} MiB")
        require(received > 0 and (length is None or received == length), "Incomplete/empty HTTP download.")
        # TNM sizes are estimates; HTTP length + GDAL inspection are the integrity gates.
        saved = {"url": product["url"], "bytes": received, "sha256": digest.hexdigest(),
                 "catalog_bytes": product.get("bytes")}
        temp.replace(path)
        receipt.write_text(json.dumps(saved, indent=2), encoding="utf-8")
        return path, saved
    finally:
        temp.unlink(missing_ok=True)


def download(product, folder, remaining_bytes, check=lambda: None, progress=lambda message: None,
             opener=open_https):
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    lock = folder / (hashlib.sha256(product["url"].encode()).hexdigest()[:24] + ".lock")
    try:
        stream = lock.open("x")
    except FileExistsError as exc:
        raise ValueError(f"Source is locked by another job: {lock}. Remove a stale lock only after confirming no worker is running.") from exc
    try:
        with stream:
            return _download(product, folder, remaining_bytes, check, progress, opener)
    finally:
        lock.unlink(missing_ok=True)
