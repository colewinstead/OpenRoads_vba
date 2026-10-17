"""Build the standalone Windows launcher with the installed .NET Framework compiler."""
import os
from pathlib import Path
import subprocess


if __name__ == "__main__":
    if os.name != "nt":
        raise SystemExit("Build on Windows with .NET Framework 4.8 installed.")
    source = Path(__file__).resolve().parent
    compiler = Path(os.environ["WINDIR"]) / "Microsoft.NET/Framework64/v4.0.30319/csc.exe"
    subprocess.run([str(compiler), "/nologo", "/target:winexe", "/platform:x64",
                    "/reference:System.Web.Extensions.dll",
                    "/out:" + str(source / "launch_worker.exe"),
                    str(source / "WorkerLauncher.cs")], check=True)
    print(source / "launch_worker.exe")
