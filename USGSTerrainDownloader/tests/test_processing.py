import json
import sys
import tempfile
import unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
from crs_utils import prepare_boundary
from terrain_processor import process_dem, inspect_raster
try:
    from osgeo import gdal, osr
    import numpy as np
    gdal.UseExceptions(); osr.UseExceptions()
    HAS_GDAL=True
except ImportError:
    HAS_GDAL=False


@unittest.skipUnless(HAS_GDAL,'GDAL unavailable: run with QGIS/OSGeo4W Python')
class ProcessingTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.root=Path(self.temp.name)
        self.confirm={'unit':'m','crs':'EPSG:5703','evidence':'Synthetic fixture declares NAVD88; not real survey data.'}
        self.config={'crs_confirmed':True,'coordinates_are_grid':True,'source_crs':'EPSG:26916','master_unit':'m',
                     'boundary':[[300000,3700000],[300020,3700000],[300020,3700020],[300000,3700020]],
                     'source_vertical':self.confirm,'resolution':10,'elevation_unit':'us-ft','clip':True}
        self.boundary=prepare_boundary(self.config)
        self.product={'id':'fixture','title':'Synthetic DEM'}

    def tearDown(self): self.temp.cleanup()

    def raster(self,name='a.tif', x=300000, y=3700020, values=None, crs=26916, unit='m', nodata=-9999, pixel=10):
        if values is None: values=np.full((2,2),100,dtype=float)
        path=self.root/name
        ds=gdal.GetDriverByName('GTiff').Create(str(path),values.shape[1],values.shape[0],1,gdal.GDT_Float64)
        ds.SetGeoTransform((x,pixel,0,y,0,-pixel))
        if crs:
            srs=osr.SpatialReference(); srs.ImportFromEPSG(crs); ds.SetProjection(srs.ExportToWkt())
        band=ds.GetRasterBand(1); band.WriteArray(values); band.SetNoDataValue(nodata); band.SetUnitType(unit)
        band=None; ds=None
        return path

    def process(self,paths,config=None,boundary=None):
        folder=self.root/'output'; folder.mkdir(exist_ok=True)
        products=[dict(self.product,id=str(i)) for i in range(len(paths))]
        return process_dem(paths,products,config or self.config,boundary or self.boundary,folder)

    def test_constant_dem_units_crs_and_readback(self):
        source=self.raster()
        report=self.process([source])
        ds=gdal.Open(report['output'])
        self.assertTrue(np.allclose(ds.ReadAsArray(),328.0833333333333,atol=1e-9))
        self.assertEqual(ds.GetRasterBand(1).GetUnitType(),'us-ft')
        self.assertEqual(report['gap_cells'],0)
        self.assertEqual(report['valid_boundary_cells'],4)
        self.assertAlmostEqual(report['elevation_multiplier'],3937/1200,places=14)
        self.assertIn('North American Vertical Datum', report['vertical_datum'])
        original=gdal.Open(str(source))
        self.assertEqual(original.ReadAsArray()[0,0],100)

    def test_international_foot_metadata(self):
        source=self.raster()
        c=dict(self.config,elevation_unit='ft')
        report=self.process([source],c)
        from pyproj import CRS
        ds=gdal.Open(report['output'])
        self.assertAlmostEqual(ds.ReadAsArray()[0,0],100/0.3048)
        vertical=CRS.from_wkt(ds.GetProjection()).sub_crs_list[1]
        self.assertEqual(vertical.axis_info[0].unit_conversion_factor,0.3048)

    def test_horizontal_reprojection_does_not_scale_z(self):
        path=self.raster('utm.tif',x=299900,y=3700100,values=np.full((40,40),123.45))
        from pyproj import Transformer
        t=Transformer.from_crs(26916,2254,always_xy=True)
        c=dict(self.config,source_crs='EPSG:2254',master_unit='us-ft',elevation_unit='m',resolution=20)
        c['boundary']=[list(t.transform(x,y)) for x,y in self.config['boundary']]
        b=prepare_boundary(c)
        result=self.process([path],c,b)
        a=gdal.Open(result['output']).ReadAsArray()
        self.assertTrue(np.allclose(a[a!=-1e300],123.45,atol=1e-8))
        self.assertEqual(result['elevation_multiplier'],1)

    def test_compatible_tiles_mosaic(self):
        a=self.raster(); b=self.raster('b.tif',x=300020)
        c=dict(self.config,boundary=[[300000,3700000],[300040,3700000],[300040,3700020],[300000,3700020]])
        result=self.process([a,b],c,prepare_boundary(c))
        self.assertEqual(result['valid_boundary_cells'],8)

    def test_reprojected_clip_keeps_fully_covered_boundary_cells(self):
        from crs_utils import transformer
        x,y=transformer(4326,26916).transform(-88.9745,30.212,errcheck=True)
        a=self.raster(x=round(x)-128,y=round(y)+128,pixel=1,values=np.full((256,256),100.))
        tx,ty=transformer(26916,2254).transform(x,y,errcheck=True)
        c=dict(self.config,source_crs='EPSG:2254',master_unit='us-ft',elevation_unit='m',resolution=3937/1200,
               boundary=[[tx-200,ty-200],[tx+200,ty-200],[tx+200,ty+200],[tx-200,ty+200]])
        result=self.process([a],c,prepare_boundary(c))
        ds=gdal.Open(result['output']);values=ds.ReadAsArray()
        self.assertEqual(result['gap_cells'],0)
        self.assertTrue(np.allclose(values,100.,atol=1e-10))

    def test_original_polygon_clips_box_exterior(self):
        a=self.raster(y=3700040,values=np.full((4,4),100.))
        c=dict(self.config,boundary=[[300000,3700000],[300040,3700000],[300000,3700040]])
        result=self.process([a],c,prepare_boundary(c))
        ds=gdal.Open(result['output']); values=ds.ReadAsArray()
        self.assertEqual(values[0,-1],ds.GetRasterBand(1).GetNoDataValue())
        self.assertGreater(values[-1,0],0)
        self.assertEqual(result['gap_cells'],0)
        self.assertLess(result['valid_boundary_cells'],16)

    def test_explicit_overlap_priority(self):
        a=self.raster(); b=self.raster('b.tif',values=np.full((2,2),200.))
        result=self.process([a,b],dict(self.config,overlap_order_confirmed=True))
        ds=gdal.Open(result['output'])
        self.assertAlmostEqual(ds.ReadAsArray()[0,0],200*3937/1200,places=10)
        self.assertEqual(result['overlaps'],[[0,1]])

    def test_cancel_during_processing_removes_partial(self):
        a=self.raster(); folder=self.root/'cancel';folder.mkdir()
        calls=0
        def cancel():
            nonlocal calls
            calls+=1
            if calls>3: raise InterruptedError('cancel during processing')
        with self.assertRaises(InterruptedError):
            process_dem([a],[self.product],self.config,self.boundary,folder,check=cancel)
        self.assertFalse((folder/'terrain.tif').exists())
        self.assertFalse((folder/'terrain.partial.tif').exists())

    def test_gap_rejected_and_original_retained(self):
        a=self.raster(values=np.array([[100.,-9999],[100.,100.]]))
        with self.assertRaisesRegex(ValueError,'Coverage gap'): self.process([a])
        self.assertTrue(a.exists()); self.assertFalse((self.root/'output/terrain.tif').exists())
        self.assertFalse((self.root/'output/terrain.partial.tif').exists())

    def test_gap_can_be_explicitly_retained(self):
        a=self.raster(values=np.array([[100.,-9999],[100.,100.]]))
        c=dict(self.config,allow_gaps=True)
        result=self.process([a],c)
        self.assertEqual(result['gap_cells'],1)

    def test_missing_crs_unit_conflict_unknown_vertical(self):
        for kwargs,confirmation in (({'crs':None},self.confirm),({'unit':'us-ft'},self.confirm),({},None)):
            path=self.raster('invalid.tif',**kwargs)
            with self.assertRaises((ValueError,TypeError)): inspect_raster(path,confirmation)

    def test_vertical_datum_mixing_rejected(self):
        a=self.raster(); b=self.raster('b.tif',x=300020)
        c=dict(self.config,vertical_confirmations={'0':self.confirm,'1':dict(self.confirm,crs='EPSG:5714')})
        with self.assertRaisesRegex(ValueError,'Mixed vertical'): self.process([a,b],c)

    def test_overlaps_and_resolution_conflict(self):
        a=self.raster(); b=self.raster('b.tif')
        with self.assertRaisesRegex(ValueError,'overlap'): self.process([a,b])
        b=self.raster('b.tif',pixel=5)
        with self.assertRaisesRegex(ValueError,'resolutions'): self.process([a,b])

    def test_cancel_and_output_cap(self):
        a=self.raster(); folder=self.root/'cancelled'; folder.mkdir()
        def cancel(): raise InterruptedError('cancel')
        with self.assertRaises(InterruptedError): process_dem([a],[self.product],self.config,self.boundary,folder,check=cancel)
        c=dict(self.config,resolution=0.0001)
        with self.assertRaisesRegex(ValueError,'100 million'): self.process([a],c)

    def test_invalid_polygon_rejected(self):
        a=self.raster()
        b=dict(self.boundary,ring=[[300000,3700000],[300020,3700020],[300000,3700020],[300020,3700000],[300000,3700000]])
        with self.assertRaisesRegex(ValueError,'self-intersecting'): self.process([a],boundary=b)
