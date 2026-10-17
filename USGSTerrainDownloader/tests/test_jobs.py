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
    def test_search_manifest_ui_and_stale_boundary(self):
        with tempfile.TemporaryDirectory() as root:
            folder=Path(root)
            config=dict(request(),action='search',dataset='dem1m')
            with patch('main.TNMClient') as client:
                client.return_value.search.return_value=[normalize(item(),'dem1m')]
                result=run(config,Job(folder))
            self.assertEqual(result['products'][0]['id'],'1')
            self.assertEqual((folder/'products.tsv').read_bytes()[:2],b'\xff\xfe')
            self.assertIn('search_complete',(folder/'status.txt').read_text(encoding='utf-16'))
            config.update(action='process',selected_ids=['1'],search_manifest=str(folder/'search.json'))
            config['boundary'][0][0]+=1
            with self.assertRaisesRegex(ValueError,'changed since search'): run(config,Job(folder))

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
            source=helper.raster()
            c=dict(helper.config,schema_version=1,action='search',dataset='dem1m')
            search=helper.root/'search'; search.mkdir()
            p=normalize(item(),'dem1m')
            with patch('main.TNMClient') as client:
                client.return_value.search.return_value=[p];run(c,Job(search))
            c.update(action='process',selected_ids=['1'],search_manifest=str(search/'search.json'),max_download_bytes=1000000)
            dest=helper.root/'process';dest.mkdir()
            with patch('main.download',return_value=(source,{'url':p['url'],'bytes':source.stat().st_size,'sha256':'fixture'})):
                report=run(c,Job(dest))
            self.assertEqual(report['state'],'dem_complete')
            self.assertTrue((dest/'terrain.tif').exists())
            self.assertEqual(json.loads((dest/'report.json').read_text())['processing']['gap_cells'],0)
            fail=helper.root/'fail';fail.mkdir()
            with patch('main.download',side_effect=OSError('network interrupted')):
                with self.assertRaises(OSError):run(c,Job(fail))
            self.assertEqual(json.loads((fail/'report.json').read_text())['state'],'failed')
            self.assertFalse((fail/'terrain.tif').exists())
        finally:
            helper.tearDown()
