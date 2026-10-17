"""Opt-in live official API smoke test. Metadata only; never downloads elevation files."""
import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
from usgs_client import TNMClient

if __name__ == '__main__':
    for kind in ('dem1m','dem13','lidar'):
        products=TNMClient().search(kind,[-88.82,33.42,-88.80,33.44])
        print(kind,len(products),'candidate products; no data downloaded')
        for p in products[:2]:
            print(' ',p['title'],p['bytes'],p['publication'])
