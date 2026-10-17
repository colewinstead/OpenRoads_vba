"""Run in a working QGIS/OSGeo4W shell; capture only GIS paths for the child worker."""
import json
import os
import sys
from pathlib import Path


def capture(folder):
    folder = Path(folder)
    if not folder.is_dir():
        raise ValueError('Choose an existing configuration folder.')
    # Verify before writing a configuration; this interpreter must support the DEM pipeline.
    from osgeo import gdal
    import numpy
    import pyproj
    if int(gdal.VersionInfo()) < 3080000:
        raise ValueError('GDAL >= 3.8 required')
    values = {k: os.environ[k] for k in ('PYTHONHOME','PYTHONPATH','GDAL_DATA','PROJ_DATA','SSL_CERT_FILE') if os.environ.get(k)}
    values['PATH_prepend'] = os.environ.get('PATH','')
    if 'PROJ_DATA' not in values:
        values['PROJ_DATA'] = os.environ.get('PROJ_LIB') or pyproj.datadir.get_data_dir()
    if 'GDAL_DATA' not in values:
        path = gdal.FindFile('gdal', 'gdalvrt.xsd')
        if path:
            values['GDAL_DATA'] = str(Path(path).parent)
    (folder/'runtime-environment.json').write_text(json.dumps(values,indent=2),encoding='utf-8')
    (folder/'runtime-python.txt').write_text(sys.executable,encoding='utf-8')
    return {'python': sys.executable, 'gdal': gdal.VersionInfo('--version'),
            'pyproj': pyproj.__version__, 'numpy': numpy.__version__}


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit('Usage: python capture_environment.py existing-config-folder')
    print(json.dumps(capture(sys.argv[1]),indent=2))
