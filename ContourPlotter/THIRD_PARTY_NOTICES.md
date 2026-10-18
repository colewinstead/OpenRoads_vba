# Licensing and provenance

The standalone Python engine, its parser, Python tests and build scripts are
GPL-2.0-or-later. LICENSE contains the full GNU GPL version 2 text. The parent
repository's MIT license does not supersede that engine license.

The independent VBA file client and ReadOnlyOrdProbe.cs are MIT licensed; see
licenses/VBA-MIT.txt. They contain no GPL parser code and communicate with the
standalone engine through documented ordinary geometry files. They do not link
the GPL engine to proprietary Bentley runtime libraries. Their source notices
retain the original repository's MIT attribution.

The standalone parser is adapted from LandXML Road & Terrain GIS Tools:
https://github.com/colewinstead/LandXML-Road-and-Terrain-GIS-Tools
Original attribution: Copyright (C) 2026 Edmond Akello.
The original project notice is retained in licenses/LandXML-NOTICE.txt.

Copied with attribution: safe_xml.py; landxml/common.py, parser.py, detection.py;
and the synthetic OpenRoads terrain_minimal.xml fixture. vendor/tin.py extracts
the original core.py TIN reader. Its modifications remove GDAL imports, avoid
repeated parsing, require explicit multi-surface choice and one TIN Definition,
filter explicit invisible faces, reject unknown flags and return source-backed
unit/CRS metadata. All modifications were made in this tool, not the QGIS plugin.
The original safe-XML security validation is retained.

NumPy, Matplotlib and their dependencies retain their own licenses and notices.
The distribution builder copies installed license/notice files, installed
versions, and Python's license into ThirdPartyLicenses. NumPy uses the BSD license;
Matplotlib includes its Matplotlib/PSF-derived license and bundled notices.
PyInstaller 6.17.0 is a build dependency, with its bootloader licensing exception;
see https://pyinstaller.org/en/stable/license.html. The exception does not remove
the GPL obligations of this reused application code.

The Windows distribution includes the complete corresponding Python/VBA source,
tests, synthetic fixture, requirements and build scripts in Source. Redistribute
those sources and notices alongside the executable and any compiled MVBA.
Company survey/project inputs are not part of the program and are not bundled.
Bentley runtime assemblies, MDOT workspace files, fonts and seed files are not
redistributed. Recipients need their licensed ORD installation and approved
workspace independently.
