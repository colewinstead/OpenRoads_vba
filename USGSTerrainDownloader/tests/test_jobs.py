import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
from main import Job, run, ui_file
from test_crs import request
from test_usgs import item
from usgs_client import normalize
try:
    from osgeo import gdal
    HAS_GDAL=True
except ImportError:
    HAS_GDAL=False


class JobTests(unittest.TestCase):
    def test_cell_size_default_uses_horizontal_crs_units_and_supported_dataset(self):
        from pyproj import Transformer
        cases = [('EPSG:26916', 'm', 'dem1m', 1),
                 ('EPSG:6576', 'us-ft', 'dem1m', 3937/1200),
                 ('+proj=utm +zone=16 +datum=NAD83 +units=ft +type=crs', 'ft', 'dem1m', 1/0.3048),
                 ('EPSG:6576', 'us-ft', 'dem13', None),
                 ('EPSG:6576', 'us-ft', 'lidar', None)]
        for crs, unit, dataset, expected in cases:
            with self.subTest(crs=crs, dataset=dataset), tempfile.TemporaryDirectory() as root:
                x, y = Transformer.from_crs(4326, crs, always_xy=True).transform(-86.8, 36.1)
                config = dict(request(), action='search', dataset=dataset, source_crs=crs, master_unit=unit,
                              boundary=[[x,y],[x+20,y],[x+20,y+20],[x,y+20]])
                with patch('main.TNMClient') as client:
                    client.return_value.search.return_value = []
                    run(config, Job(root))
                value = (Path(root)/'cell-size.txt').read_text(encoding='utf-16').strip()
                if expected is None:
                    self.assertEqual(value, '')
                else:
                    self.assertAlmostEqual(float(value), expected, places=10)

    def test_search_manifest_ui_and_stale_boundary(self):
        with tempfile.TemporaryDirectory() as root:
            folder=Path(root)
            config=dict(request(),action='search',dataset='dem1m')
            with patch('main.TNMClient') as client:
                product = item()
                product.update(body='All bare earth elevation values are in meters and are referenced to the North American Vertical Datum of 1988 (NAVD88).',
                               metaUrl='https://example.org/product/1')
                client.return_value.search.return_value=[normalize(product,'dem1m')]
                result=run(config,Job(folder))
            self.assertEqual(result['products'][0]['id'],'1')
            self.assertEqual((folder/'products.tsv').read_bytes()[:2],b'\xff\xfe')
            row = (folder/'products.tsv').read_text(encoding='utf-16').strip().split('\t')
            self.assertEqual(len(row), 12)
            self.assertEqual(row[9:11], ['m', 'EPSG:5703'])
            self.assertIn('https://example.org/product/1', row[11])
            self.assertIn('search_complete',(folder/'status.txt').read_text(encoding='utf-16'))
            self.assertNotIn('\r\r\n', (folder/'status.txt').read_bytes().decode('utf-16'))
            config.update(action='process',selected_ids=['1'],search_manifest=str(folder/'search.json'))
            config['boundary'][0][0]+=1
            with self.assertRaisesRegex(ValueError,'changed since search'): run(config,Job(folder))

    @unittest.skipUnless(HAS_GDAL, 'GDAL unavailable')
    def test_catalog_vertical_conflict_stops_before_transfer(self):
        with tempfile.TemporaryDirectory() as root:
            folder = Path(root)
            product = dict(item(), body='All bare earth elevation values are in meters and are referenced to the North American Vertical Datum of 1988 (NAVD88).',
                           metaUrl='https://example.org/product/1')
            config = dict(request(), action='search', dataset='dem1m')
            with patch('main.TNMClient') as client:
                client.return_value.search.return_value = [normalize(product, 'dem1m')]
                run(config, Job(folder))
            config.update(action='process', selected_ids=['1'], search_manifest=str(folder/'search.json'))
            for unit, crs in (('us-ft', 'EPSG:5703'), ('m', 'EPSG:5714')):
                config['source_vertical'] = {'unit': unit, 'crs': crs, 'evidence': 'Manual declaration conflicting with catalog'}
                with self.subTest(unit=unit, crs=crs), patch('main.download') as transfer:
                    with self.assertRaisesRegex(ValueError, 'contradicts selected product metadata'):
                        run(config, Job(folder))
                    transfer.assert_not_called()

    def test_cooperative_cancel_and_utf16(self):
        with tempfile.TemporaryDirectory() as root:
            folder=Path(root); (folder/'cancel.flag').touch()
            with self.assertRaises(InterruptedError): Job(folder).check()
            ui_file(folder/'unicode.tsv',['Δ\t工程\tUS survey feet'])
            self.assertIn('工程',(folder/'unicode.tsv').read_text(encoding='utf-16'))

    def test_unknown_size_needs_confirmation_before_transfer(self):
        with tempfile.TemporaryDirectory() as root:
            folder=Path(root)
            c=dict(request(),action='search',dataset='dem1m')
            p=normalize(item(),'dem1m');p['bytes']=None
            with patch('main.TNMClient') as client:
                client.return_value.search.return_value=[p]; run(c,Job(folder))
            c.update(action='process',selected_ids=['1'],search_manifest=str(folder/'search.json'))
            with patch('main.download') as transfer:
                with self.assertRaisesRegex(ValueError,'unknown size'): run(c,Job(folder))
                transfer.assert_not_called()

    @unittest.skipUnless(HAS_GDAL,'GDAL unavailable')
    def test_capture_environment_whitelist(self):
        from capture_environment import capture
        with tempfile.TemporaryDirectory() as root, patch.dict('os.environ',{'TERRAIN_TEST_SECRET':'must-not-be-captured'}):
            result=capture(root)
            env=json.loads((Path(root)/'runtime-environment.json').read_text())
            self.assertNotIn('TERRAIN_TEST_SECRET',env)
            self.assertIn('PATH_prepend',env)
            self.assertTrue(result['python'])

    @unittest.skipUnless(HAS_GDAL,'GDAL unavailable')
    def test_full_job_and_failure_report_with_tiny_raster(self):
        from test_processing import ProcessingTests
        helper=ProcessingTests(); helper.setUp()
        try:
            import numpy as np
            source=helper.raster(x=299980,y=3700040,values=np.full((6,6),100.))
            c=dict(helper.config,schema_version=1,action='search',dataset='dem1m')
            search=helper.root/'search'; search.mkdir()
            p=normalize(item(),'dem1m')
            with patch('main.TNMClient') as client:
                client.return_value.search.return_value=[p];run(c,Job(search))
            c.update(action='process',selected_ids=['1'],search_manifest=str(search/'search.json'),max_download_bytes=1000000)
            for clip in (True,False):
                with self.subTest(clip=clip):
                    c['clip']=clip
                    dest=helper.root/str(clip);dest.mkdir()
                    with patch('main.download',return_value=(source,{'url':p['url'],'bytes':source.stat().st_size,'sha256':'fixture'})):
                        report=run(c,Job(dest))
                    self.assertEqual(report['state'],'dem_complete')
                    self.assertTrue((dest/'terrain.tif').exists())
                    self.assertTrue((dest/'terrain.xml').exists())
                    self.assertGreater(report['landxml']['faces'],0)
                    self.assertEqual(json.loads((dest/'report.json').read_text())['processing']['gap_cells'],0)
                    guide=(dest/'ORD_IMPORT.txt').read_text(encoding='utf-8')
                    if clip:
                        self.assertIn('original selected shape',guide)
                        self.assertTrue(report['landxml']['clipped'])
                    else:
                        self.assertIn('Polygon clipping was disabled',guide)
                        self.assertFalse(report['landxml']['clipped'])
                    self.assertIn('Build Terrain From = Definition',guide)
                    self.assertIn('terrain.xml',(dest/'status.txt').read_text(encoding='utf-16'))
            mesh_fail=helper.root/'mesh-fail';mesh_fail.mkdir()
            with patch('main.download',return_value=(source,{'url':p['url'],'bytes':source.stat().st_size,'sha256':'fixture'})), \
                 patch('landxml_export.export_landxml',side_effect=ValueError('mesh export failed')):
                with self.assertRaisesRegex(ValueError,'mesh export failed'):run(c,Job(mesh_fail))
            failure=json.loads((mesh_fail/'report.json').read_text())
            self.assertEqual(failure['state'],'failed')
            self.assertIn('processing',failure)
            self.assertTrue((mesh_fail/'terrain.tif').exists())
            self.assertFalse((mesh_fail/'terrain.xml').exists())
            fail=helper.root/'fail';fail.mkdir()
            with patch('main.download',side_effect=OSError('network interrupted')):
                with self.assertRaises(OSError):run(c,Job(fail))
            self.assertEqual(json.loads((fail/'report.json').read_text())['state'],'failed')
            self.assertFalse((fail/'terrain.tif').exists())
        finally:
            helper.tearDown()
