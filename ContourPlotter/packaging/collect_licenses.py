# SPDX-License-Identifier: GPL-2.0-or-later
"""Copy the installed packages' license notices into the distributable."""
import ast
import importlib.metadata as md
from pathlib import Path
import shutil
import sys

NAMES = ["numpy", "matplotlib", "contourpy", "cycler", "fonttools", "kiwisolver",
         "packaging", "pillow", "pyparsing", "python-dateutil", "six", "pyinstaller"]


def collect(destination):
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    versions = []
    names = set(NAMES)
    toc = Path(__file__).resolve().parents[1] / "build/ContourPlotter/Analysis-00.toc"
    if toc.exists():
        mapping = md.packages_distributions()
        def visit(value):
            if isinstance(value, (tuple, list)):
                if len(value) == 3 and isinstance(value[0], str) and value[2] in ("PYMODULE", "EXTENSION"):
                    names.update(mapping.get(value[0].split(".")[0], []))
                for child in value:
                    visit(child)
        visit(ast.literal_eval(toc.read_text(encoding="utf-8")))
    for name in sorted(names, key=str.lower):
        dist = md.distribution(name)
        target = destination / name
        target.mkdir(exist_ok=True)
        versions.append(f"{dist.metadata['Name']}=={dist.version}")
        for entry in dist.files or []:
            if any(word in str(entry).lower() for word in ("license", "copying", "notice", "authors")):
                source = Path(dist.locate_file(entry))
                if source.is_file():
                    safe = "_".join(entry.parts)
                    shutil.copy2(source, target / safe)
    for name in ("LICENSE.txt", "LICENSE"):
        source = Path(sys.base_prefix) / name
        if source.is_file():
            shutil.copy2(source, destination / "Python-LICENSE.txt")
            break
    (destination / "installed-versions.txt").write_text("\n".join(versions) + "\n", encoding="utf-8")


if __name__ == "__main__":
    collect(sys.argv[1])
