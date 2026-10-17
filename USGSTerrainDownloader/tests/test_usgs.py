import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from urllib.error import URLError
from unittest.mock import patch
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'python'))
from usgs_client import DATASETS, TNMClient, download, normalize


class Response(io.BytesIO):
    def __init__(self, value, length=None):
        raw=value if isinstance(value,bytes) else json.dumps(value).encode()
        super().__init__(raw)
        self.headers={} if length is None else {'Content-Length': str(length)}


def item(n=1):
    return {'sourceId': str(n), 'title': f'Tile {n}', 'sizeInBytes': 4,
            'downloadURL': f'https://example.org/{n}.tif',
            'publicationDate': '2024-01-01', 'boundingBox': {'minX':-89,'minY':33,'maxX':-88,'maxY':34}}


class ClientTests(unittest.TestCase):
    def test_catalog_query_pagination_dedup(self):
        urls=[]
        def open_url(url):
            urls.append(url)
            if '/datasets' in url:
                return Response([{'tags':[{'sbDatasetTag':DATASETS['dem1m']}]}])
            offset=int(parse_qs(urlparse(url).query)['offset'][0])
            return Response({'total':3,'items': [item(1), item(1)] if offset==0 else [item(2)],'errors':[]})
        products=TNMClient(opener=open_url).search('dem1m',[-88.9,33.1,-88.8,33.2])
        self.assertEqual([p['id'] for p in products],['1','2'])
        query=parse_qs(urlparse(urls[1]).query)
        self.assertEqual(query['datasets'],[DATASETS['dem1m']])
        self.assertEqual(query['prodFormats'],['GeoTIFF'])
        self.assertEqual(query['bbox'],['-88.9000000000,33.1000000000,-88.8000000000,33.2000000000'])
        self.assertEqual(parse_qs(urlparse(urls[2]).query)['offset'],['2'])

    def test_total_limit_missing_coverage_and_api_errors(self):
        for page in ({'total':0,'items':[]}, {'total':501,'items':[]},
                     {'total':1,'items':[]}, {'total':0,'items':[], 'errors':['outage']}):
            def opener(url):
                return Response([{'sbDatasetTag':DATASETS['dem13']}]) if '/datasets' in url else Response(page)
            if page == {'total':0,'items':[]}:
                self.assertEqual(TNMClient(opener).search('dem13',[-89,33,-88,34]),[])
            else:
                with self.assertRaises(ValueError): TNMClient(opener).search('dem13',[-89,33,-88,34])

    def test_repeated_page_guard(self):
        def opener(url):
            return Response([{'sbDatasetTag':DATASETS['dem1m']}]) if '/datasets' in url else Response({'total':3,'items':[item()]})
        with self.assertRaisesRegex(ValueError,'repeated'): TNMClient(opener).search('dem1m',[-89,33,-88,34])

    def test_network_retry_is_bounded(self):
        calls=[]
        def offline(url):
            calls.append(url)
            raise URLError('offline')
        with patch('usgs_client.time.sleep'):
            with self.assertRaises(URLError): TNMClient(offline).get_json('datasets')
        self.assertEqual(len(calls),3)

    def test_product_url_and_metadata_validation(self):
        p=item(); p['downloadURL']='http://example.org/1.tif'
        with self.assertRaises(ValueError): normalize(p,'dem1m')
        p=item(); p['boundingBox']=None
        with self.assertRaises(ValueError): normalize(p,'dem1m')
        p=item(); p['downloadURL']='https://example.org/1.zip'
        with self.assertRaises(ValueError): normalize(p,'dem1m')
        p=item(); p['downloadURL']='https://example.org/1.laz'
        self.assertEqual(normalize(p,'lidar')['kind'],'lidar')

    def test_download_receipt_reuse_and_corruption(self):
        with tempfile.TemporaryDirectory() as root:
            p=normalize(item(),'dem1m')
            path, receipt=download(p,root,10,opener=lambda u:Response(b'data',4))
            self.assertEqual(path.read_bytes(),b'data')
            self.assertEqual(receipt['bytes'],4)
            path2,_=download(p,root,10,opener=lambda u:self.fail('cache should be reused'))
            self.assertEqual(path2,path)
            path.write_bytes(b'bad!')
            with self.assertRaises(ValueError): download(p,root,10)

    def test_incomplete_capped_cancelled_transfers_leave_no_partial(self):
        for length,budget in ((20,100),(4,3),(None,3)):
            with tempfile.TemporaryDirectory() as root:
                with self.assertRaises(ValueError):
                    download(normalize(item(),'dem1m'),root,budget,opener=lambda u:Response(b'data',length))
                self.assertEqual(list(Path(root).iterdir()),[])
        with tempfile.TemporaryDirectory() as root:
            def cancel(): raise InterruptedError('cancel')
            with self.assertRaises(InterruptedError):
                download(normalize(item(),'dem1m'),root,10,check=cancel,opener=lambda u:Response(b'data',4))
            self.assertEqual(list(Path(root).iterdir()),[])
