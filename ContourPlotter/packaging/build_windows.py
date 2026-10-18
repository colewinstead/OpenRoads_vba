# SPDX-License-Identifier: GPL-2.0-or-later
"""Build and verify a complete Windows distribution, without changing system settings."""
import argparse
import os
from pathlib import Path
import shutil
import stat
import subprocess
import sys
import uuid
from collect_licenses import collect

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dependency-path", type=Path, help="Optional private installed package directory")
    args = parser.parse_args()
    paths = [str(ROOT / "python")]
    if args.dependency_path:
        paths.insert(0, str(args.dependency_path.resolve()))
        sys.path.insert(0, str(args.dependency_path.resolve()))
    env = dict(os.environ, PYTHONPATH=os.pathsep.join(paths))
    subprocess.run([sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v"],
                   cwd=ROOT, env=env, check=True)
    if os.name == "nt":
        # OneDrive marked these generated folders read-only; reset only build artifacts.
        for target in (ROOT/"build/ContourPlotter", ROOT/"dist/ContourPlotter"):
            if target.exists():
                if target.is_symlink() or not target.resolve().is_relative_to(ROOT.resolve()):
                    raise ValueError("Generated cleanup target escaped the project")
                for item in [target, *target.rglob("*")]:
                    if not item.is_symlink():
                        item.chmod(stat.S_IREAD | stat.S_IWRITE)
    command = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onedir", "--console",
        "--name", "ContourPlotter", "--distpath", str(ROOT/"dist"), "--workpath", str(ROOT/"build"),
        "--specpath", str(ROOT/"build"), "--paths", str(ROOT/"python"),
        "--hidden-import", "matplotlib.backends.backend_agg"]
    if args.dependency_path:
        command += ["--paths", str(args.dependency_path.resolve())]
    command.append(str(ROOT/"python/run_contour_plotter.py"))
    subprocess.run(command, cwd=ROOT, env=env, check=True)
    package = ROOT/"dist/ContourPlotter"
    source = package/"Source"
    if source.exists():
        if not source.resolve().is_relative_to((ROOT/"dist").resolve()):
            raise ValueError("Source cleanup target escaped the generated distribution")
        shutil.rmtree(source)
    source.mkdir(parents=True)
    for directory in ("python", "vba", "packaging"):
        for file in (ROOT/directory).rglob("*"):
            if file.is_file() and file.suffix in {".py", ".bas", ".vba", ".ps1", ".cs"} and "__pycache__" not in file.parts:
                destination = source/file.relative_to(ROOT)
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(file, destination)
    shutil.copytree(ROOT/"licenses", source/"licenses")
    tests = source/"tests"
    tests.mkdir()
    for file in (ROOT/"tests").glob("*"):
        if file.suffix in {".py", ".cs"}:
            shutil.copy2(file, tests/file.name)
    (tests/"fixtures").mkdir()
    # Whitelist the synthetic fixture: never distribute a coworker's project XML.
    shutil.copy2(ROOT/"tests/fixtures/openroads_minimal.xml", tests/"fixtures/openroads_minimal.xml")
    for name in ("README.md", "ARCHITECTURE.md", "API_NOTES.md", "ORD_TEST_PLAN.md", "FORMAT.md",
                 "TEST_RESULTS.md", "LICENSE", "THIRD_PARTY_NOTICES.md", "requirements.txt"):
        shutil.copy2(ROOT/name, source/name)
    for name in ("LICENSE", "README.md"):
        shutil.copy2(ROOT/name, package/name)
    collect(package/"ThirdPartyLicenses")
    if (ROOT/"vba/ContourPlotter.mvba").is_file():
        shutil.copy2(ROOT/"vba/ContourPlotter.mvba", package/"ContourPlotter.mvba")
    executable = package/"ContourPlotter.exe"
    subprocess.run([str(executable), "--version"], cwd=ROOT, check=True)
    smoke = ROOT/"jobs"/("packaged_smoke_" + uuid.uuid4().hex)
    subprocess.run([str(executable), "--input", str(tests/"fixtures/openroads_minimal.xml"),
        "--output-dir", str(smoke), "--axis-order", "xy",
        "--vertical-unit", "us-ft", "--model-unit", "us-ft", "--crs", "Synthetic local grid",
        "--vertical-datum", "Synthetic datum", "--minor", ".5", "--major", "1",
        "--label-mode", "none"], cwd=ROOT, check=True)
    archive = shutil.make_archive(str(ROOT/"dist/ContourPlotter-Windows-x64"), "zip", ROOT/"dist", "ContourPlotter")
    print(f"Built {package}\nDistribution: {archive}")


if __name__ == "__main__":
    main()
