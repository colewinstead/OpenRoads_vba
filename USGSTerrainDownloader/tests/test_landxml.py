import json
import math
import sys
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'python'))
from crs_utils import UNIT_METERS, prepare_boundary
from landxml_export import NS, export_landxml, triangulate_ring, cross
from terrain_processor import process_dem
from test_processing import HAS_GDAL
import test_processing


@unittest.skipUnless(HAS_GDAL, 'GDAL unavailable: run with QGIS/OSGeo4W Python')
class LandXMLTests(unittest.TestCase):
    def setUp(self):
        from osgeo import ogr
        import numpy as np
        self.ogr, self.np = ogr, np
        self.helper = test_processing.ProcessingTests(); self.helper.setUp()
        self.root = self.helper.root
        x, y = np.meshgrid(299985+np.arange(10)*10, 3700075-np.arange(10)*10)
        self.values = 100+0.2*(x-300000)+0.3*(y-3700000)
        self.config = dict(self.helper.config)
        self.source = self.helper.raster(x=299980,y=3700080,values=self.values)

    def tearDown(self):
        self.helper.tearDown()

    def export(self, config=None, source=None, name='output'):
        config = config or self.config
        source = source or self.source
        folder = self.root/name; folder.mkdir()
        boundary = prepare_boundary(config)
        processing = process_dem([source],[self.helper.product],config,boundary,folder)
        result = export_landxml([source],config,boundary,processing,folder)
        return result, processing, folder

    def read(self, result):
        root = ET.parse(result['output']).getroot()
        ns = {'l':NS}
        points = {}
        for point in root.findall('.//l:P',ns):
            n,e,z = map(float,point.text.split())
            points[int(point.attrib['id'])] = (e,n,z)
        faces = [tuple(map(int,face.text.split())) for face in root.findall('.//l:F',ns)]
        self.assertEqual(len(points),result['points'])
        self.assertEqual(len(faces),result['faces'])
        self.assertEqual(set(points),set(v for face in faces for v in face))
        self.assertTrue(all(math.isfinite(v) for p in points.values() for v in p))
        self.assertEqual(len(faces),len(set(tuple(sorted(f)) for f in faces)))
        return root, points, faces

    def geometry(self, vertices):
        linear = self.ogr.Geometry(self.ogr.wkbLinearRing)
        for x,y in vertices: linear.AddPoint_2D(x,y)
        linear.CloseRings()
        polygon = self.ogr.Geometry(self.ogr.wkbPolygon); polygon.AddGeometry(linear)
        return polygon

    def check_surface(self, result, polygon, plane):
        root, points, faces = self.read(result)
        union_input = self.ogr.Geometry(self.ogr.wkbMultiPolygon)
        edges = {}
        for ids in faces:
            xyz = [points[v] for v in ids]
            self.assertGreater(cross(*xyz),0)
            triangle = self.geometry([(x,y) for x,y,z in xyz])
            self.assertLessEqual(triangle.Difference(polygon).GetArea(),result['area_tolerance'])
            union_input.AddGeometry(triangle)
            for a,b in zip(ids,ids[1:]+ids[:1]):
                edge = tuple(sorted((a,b))); edges[edge] = edges.get(edge,0)+1
        self.assertLessEqual(max(edges.values()),2)
        union = union_input.UnionCascaded()
        self.assertLessEqual(polygon.SymDifference(union).GetArea(),result['area_tolerance']*4)
        self.assertAlmostEqual(union.GetArea(),result['triangle_area'],delta=result['area_tolerance']*4)
        for x,y,z in points.values():
            self.assertAlmostEqual(z,plane(x,y),places=7)
        return root, points, faces

    def test_concave_and_diagonal_boundary_preserves_plane_and_exact_extent(self):
        ring = [[300000,3700000],[300040,3700000],[300040,3700015],
                [300013,3700015],[300013,3700040],[300000,3700033]]
        config = dict(self.config,boundary=ring,elevation_unit='us-ft')
        result,_,_ = self.export(config)
        root,points,_ = self.check_surface(result,self.geometry(ring),
                                           lambda x,y:100+0.2*(x-300000)+0.3*(y-3700000))
        self.assertEqual(root.find(f'{{{NS}}}Units/{{{NS}}}Metric').attrib['linearUnit'],'meter')
        self.assertEqual(result['vertical_unit'],'m')
        self.assertLessEqual(result['gap_area'],result['area_tolerance'])
        for vertex in ring:
            self.assertTrue(any(math.dist(vertex,p[:2])<result['xy_join_tolerance']*2 for p in points.values()))

    def test_metres_international_feet_and_survey_feet_are_not_mixed(self):
        for unit in ('m','ft','us-ft'):
            with self.subTest(unit=unit):
                factor = UNIT_METERS[unit]
                crs = 'EPSG:26916' if unit == 'm' else f'+proj=utm +zone=16 +datum=NAD83 +units={unit} +type=crs'
                config = dict(self.config,source_crs=crs,master_unit=unit,resolution=10/factor,elevation_unit='m',
                              boundary=[[x/factor,y/factor] for x,y in self.config['boundary']])
                result,_,_ = self.export(config,name=unit)
                polygon = self.geometry(config['boundary'])
                root,_,_ = self.check_surface(result,polygon,
                    lambda x,y:(100+0.2*(x*factor-300000)+0.3*(y*factor-3700000))/factor)
                units = list(root.find(f'{{{NS}}}Units'))[0]
                self.assertEqual(units.attrib['linearUnit'],{'m':'meter','ft':'foot','us-ft':'USSurveyFoot'}[unit])
                self.assertEqual(result['horizontal_unit'],unit)
                self.assertEqual(result['vertical_unit'],unit)

    def test_master_display_units_do_not_change_native_landxml_coordinates(self):
        factor = UNIT_METERS['us-ft']
        config = dict(self.config,master_unit='us-ft',boundary=[[x/factor,y/factor] for x,y in self.config['boundary']])
        result,_,_ = self.export(config)
        self.check_surface(result,self.geometry(self.config['boundary']),
                           lambda x,y:100+0.2*(x-300000)+0.3*(y-3700000))

    def test_disabled_clip_keeps_rectangle(self):
        config = dict(self.config,clip=False,boundary=[[300000,3700000],[300020,3700000],[300000,3700020]])
        result,_,_ = self.export(config)
        self.check_surface(result,self.geometry(self.config['boundary']),
                           lambda x,y:100+0.2*(x-300000)+0.3*(y-3700000))
        self.assertFalse(result['clipped'])

    def test_nodata_hole_is_never_bridged(self):
        values = self.values.copy(); values[6,3] = -9999
        source = self.helper.raster('gap.tif',x=299980,y=3700080,values=values)
        config = dict(self.config,allow_gaps=True)
        result,processing,folder = self.export(config,source)
        self.assertGreater(result['gap_area'],result['area_tolerance'])
        _,points,faces = self.read(result)
        missing = self.ogr.Geometry(self.ogr.wkbPoint); missing.AddPoint_2D(300015,3700015)
        for face in faces:
            triangle = self.geometry([points[v][:2] for v in face])
            self.assertFalse(triangle.Intersects(missing))
        other = self.root/'gap-rejected';other.mkdir()
        with self.assertRaisesRegex(ValueError,'LandXML coverage gap'):
            export_landxml([source],dict(config,allow_gaps=False),prepare_boundary(config),processing,other)
        self.assertFalse((other/'terrain.xml').exists())
        self.assertEqual(list(other.iterdir()),[])
        self.assertTrue((folder/'terrain.tif').exists())

    def test_missing_boundary_samples_fail_without_extrapolation(self):
        source = self.helper.raster('edge.tif')
        folder = self.root/'edge';folder.mkdir()
        boundary = prepare_boundary(self.config)
        processing = process_dem([source],[self.helper.product],self.config,boundary,folder)
        with self.assertRaisesRegex(ValueError,'LandXML coverage gap'):
            export_landxml([source],self.config,boundary,processing,folder)
        self.assertTrue((folder/'terrain.tif').exists())
        self.assertFalse((folder/'terrain.xml').exists())
        self.assertFalse(any(folder.glob('*.tmp')))
        self.assertFalse(any(folder.glob('*working*')))

    def test_cancellation_and_size_limit_leave_no_partial_xml(self):
        folder = self.root/'cancel';folder.mkdir()
        boundary = prepare_boundary(self.config)
        processing = process_dem([self.source],[self.helper.product],self.config,boundary,folder)
        cancelled = False
        def progress(message):
            nonlocal cancelled
            if 'clipping triangles' in message: cancelled=True
        def check():
            if cancelled: raise InterruptedError('cancel mesh')
        with self.assertRaisesRegex(InterruptedError,'cancel mesh'):
            export_landxml([self.source],self.config,boundary,processing,folder,check,progress)
        self.assertFalse((folder/'terrain.xml').exists())
        self.assertFalse(any(folder.glob('*.tmp')))
        self.assertFalse(any(folder.glob('*working*')))
        with patch('landxml_export.MAX_GRID_CELLS',1):
            with self.assertRaisesRegex(ValueError,'grid-cell limit'):
                export_landxml([self.source],self.config,boundary,processing,folder)
        self.assertTrue((folder/'terrain.tif').exists())
        with patch('landxml_export.MAX_FACES',1):
            with self.assertRaisesRegex(ValueError,'two million faces'):
                export_landxml([self.source],self.config,boundary,processing,folder)
        self.assertFalse((folder/'terrain.xml').exists())
        self.assertFalse(any(folder.glob('*.tmp')))

    def test_nonfinite_boundary_elevation_is_rejected(self):
        folder=self.root/'nonfinite';folder.mkdir()
        boundary=prepare_boundary(self.config)
        processing=process_dem([self.source],[self.helper.product],self.config,boundary,folder)
        values=self.values.copy();values[5,3]=float('nan')
        source=self.helper.raster('nan.tif',x=299980,y=3700080,values=values)
        with self.assertRaisesRegex(ValueError,'Nonfinite LandXML source'):
            export_landxml([source],self.config,boundary,processing,folder)
        self.assertFalse((folder/'terrain.xml').exists())

    def test_concave_ear_clipping_keeps_area(self):
        ring = [(0,0),(10,0),(10,10),(8,10),(8,2),(2,2),(2,10),(0,10)]
        triangles = list(triangulate_ring(ring))
        self.assertEqual(len(triangles),len(ring)-2)
        self.assertAlmostEqual(sum(abs(cross(*t))/2 for t in triangles),52)
        polygon = self.geometry(ring)
        for triangle in triangles:
            self.assertAlmostEqual(self.geometry(triangle).Difference(polygon).GetArea(),0)