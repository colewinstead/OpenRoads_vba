import math
import sys
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'python'))
from pyproj import CRS, Transformer
from crs_utils import prepare_boundary, linear_factor, densify, ring_valid


def request(epsg=26916):
    x, y = Transformer.from_crs(4326, epsg, always_xy=True).transform(-88.81, 33.43)
    return {'schema_version': 1, 'crs_confirmed': True, 'coordinates_are_grid': True,
            'source_crs': f'EPSG:{epsg}', 'master_unit': 'm',
            'boundary': [[x,y], [x+100,y], [x+100,y+100], [x,y+100]]}


class CRSTests(unittest.TestCase):
    def test_unit_definitions(self):
        self.assertEqual(linear_factor('ft','m'), 0.3048)
        self.assertEqual(linear_factor('us-ft','m'), 1200/3937)
        self.assertNotEqual(linear_factor('ft','m'), linear_factor('us-ft','m'))
        self.assertAlmostEqual(100 * linear_factor('m','us-ft'), 328.0833333333333)
        with self.assertRaises(ValueError): linear_factor('feet','m')

    def test_mississippi_crs_and_master_conversion(self):
        # EPSG database definitions, not a zone inferred from numeric coordinates.
        for epsg in (26994,26995,2254,2255,6506,6507,6508,6509,6510):
            crs = CRS.from_epsg(epsg)
            config = request(epsg)
            if 'West' in crs.name:
                x,y=Transformer.from_crs(4326,epsg,always_xy=True).transform(-90.5,33.43)
                config['boundary']=[[x,y],[x+100,y],[x+100,y+100],[x,y+100]]
            native = crs.axis_info[0].unit_conversion_factor
            config['master_unit'] = 'us-ft' if abs(native-1200/3937) < 1e-12 else 'm'
            b = prepare_boundary(config)['bbox']
            lon=-90.5 if 'West' in crs.name else -88.81
            self.assertLess(b[0], lon); self.assertGreater(b[2], lon)
        config = request(2254)
        config['master_unit'] = 'ft'
        # Feed master international-foot coordinates corresponding to EPSG survey feet.
        ratio = (1200/3937)/0.3048
        config['boundary'] = [[x*ratio,y*ratio] for x,y in config['boundary']]
        self.assertAlmostEqual(prepare_boundary(config)['model_to_native_scale'], 0.999998)

    def test_confirmation_and_reasonableness(self):
        for key in ('crs_confirmed','coordinates_are_grid'):
            c=request(); c[key]=False
            with self.assertRaises(ValueError): prepare_boundary(c)
        for ring in ([[0,0],[1,1],[2,2]], [[0,0],[1,0],[math.nan,1]]):
            c=request(); c['boundary']=ring
            with self.assertRaises(ValueError): prepare_boundary(c)
        c=request(); c['source_crs']='EPSG:4326'
        with self.assertRaises(ValueError): prepare_boundary(c)
        c=request(); c['boundary']=[[0,0],[10,0],[10,10],[0,10]]
        with self.assertRaises(ValueError): prepare_boundary(c)

    def test_densifies_all_edges(self):
        ring=ring_valid([[0,0],[1000,0],[1000,1000],[0,1000]])
        dense=densify(ring,1)
        self.assertEqual(len(dense),41)
        self.assertEqual(dense[0],dense[-1])
        self.assertTrue(all(math.dist(a,b)<=100 for a,b in zip(dense,dense[1:])))

    def test_bbox_contains_independent_dense_probes(self):
        c=request(26916)
        x,y=c['boundary'][0]; c['boundary']=[[x,y],[x+10000,y],[x+10000,y+10000],[x,y+10000]]
        b=prepare_boundary(c)['bbox']
        transform=Transformer.from_crs(c['source_crs'],4326,always_xy=True)
        points=densify(ring_valid(c['boundary']),1,segment_m=7)
        for x,y in points:
            lon,lat=transform.transform(x,y)
            self.assertTrue(b[0]<=lon<=b[2] and b[1]<=lat<=b[3])
