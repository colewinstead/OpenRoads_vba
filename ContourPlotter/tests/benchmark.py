# SPDX-License-Identifier: GPL-2.0-or-later
"""Reproducible synthetic TIN benchmark; includes parsing and disk serialization."""
import argparse
import json
from pathlib import Path
import platform
import tempfile
import numpy
import matplotlib
from contour_plotter.cli import run


def make_grid(path, size):
    with path.open("w", encoding="utf-8") as stream:
        stream.write('<LandXML><Units><Metric linearUnit="meter" verticalLinearUnit="meter"/></Units><Surfaces><Surface name="Synthetic benchmark"><Definition surfType="TIN"><Pnts>')
        for y in range(size):
            for x in range(size):
                stream.write(f'<P id="{y*size+x+1}">{2_000_000+x*10} {1_000_000+y*10} {(x+y)*.1}</P>')
        stream.write('</Pnts><Faces>')
        for y in range(size-1):
            for x in range(size-1):
                a = y*size+x+1
                stream.write(f'<F>{a} {a+1} {a+size+1}</F><F>{a} {a+size+1} {a+size}</F>')
        stream.write('</Faces></Definition></Surface></Surfaces></LandXML>')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--sizes", nargs="+", type=int, default=[33, 129, 257, 513])
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    rows = []
    for size in args.sizes:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "grid.xml"
            make_grid(source, size)
            result = run(dict(version=1, input=str(source), output_dir=str(Path(tmp)/"output"),
                spatial=dict(axis_order="xy", model_unit="m", crs="Synthetic local grid",
                             vertical_datum="Synthetic datum"),
                settings=dict(minor=1, major=5, scale=50, paper_height=.1)))
            rows.append(dict(grid=size, points=size*size, triangles=2*(size-1)**2,
                             source_bytes=source.stat().st_size, **result["statistics"]))
            print(json.dumps(rows[-1]), flush=True)
    report = dict(python=platform.python_version(), numpy=numpy.__version__,
                  matplotlib=matplotlib.__version__, platform=platform.platform(),
                  notes="One run each, synthetic plane, cold/warm caches differ. Excludes ORD insertion/save/plot.",
                  measurements=rows)
    if args.output:
        args.output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
