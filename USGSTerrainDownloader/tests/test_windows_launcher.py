"""Exercise the real standalone EXE; no ORD access or network requests."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import unittest

SOURCE = Path(__file__).resolve().parents[1] / "python"


@unittest.skipUnless(os.name == "nt", "Windows standalone launcher")
class WindowsLauncherTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="terrain Δ & ' launcher-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.request = self.root / "request Δ & '.json"
        self.launcher = SOURCE / "launch_worker.exe"
        self.assertTrue(self.launcher.is_file(), "Run python/build_launcher.py first.")

    def configure(self, **extra):
        config = {"schema_version": 0, "action": "search", "python_executable": sys.executable}
        config.update(extra)
        self.request.write_text(json.dumps(config), encoding="utf-8-sig")

    def run_launcher(self):
        return subprocess.run([str(self.launcher), str(self.request)], timeout=30).returncode

    def fake_worker(self, code):
        self.launcher = self.root / "launch_worker.exe"
        shutil.copy2(SOURCE / "launch_worker.exe", self.launcher)
        (self.root / "main.py").write_text(code, encoding="utf-8")

    def status(self):
        return (self.root / "status.txt").read_text(encoding="utf-16")

    def test_real_worker_failure_unicode_paths(self):
        self.configure()
        self.assertNotEqual(self.run_launcher(), 0)
        self.assertTrue(self.status().startswith("failed"))
        self.assertIn("Unsupported request schema", (self.root / "error.log").read_text(encoding="utf-8"))
        self.assertTrue((self.root / "stdout.log").is_file())
        self.assertTrue((self.root / "stderr.log").is_file())

    def test_missing_interpreter(self):
        self.configure(python_executable=str(self.root / "missing.exe"))
        self.assertEqual(self.run_launcher(), 2)
        self.assertIn("Python executable not found", self.status())

    def test_child_environment_logs_exit_and_parent_isolation(self):
        self.fake_worker("""import json, os, pathlib, sys
pathlib.Path('child.json').write_text(json.dumps({key: os.environ.get(key) for key in ('GDAL_DATA','PATH','PROJ_NETWORK','PYTHONIOENCODING')}), encoding='utf-8')
sys.stdout.write('Δ' * 100000)
sys.stderr.write('é' * 100000)
sys.exit(7)
""")
        environment = self.root / "environment Δ & '.json"
        environment.write_text(json.dumps({"GDAL_DATA": "child Δ", "PATH_prepend": str(self.root)}), encoding="utf-8")
        before = dict(os.environ)
        self.configure(environment_file=str(environment))
        self.assertEqual(self.run_launcher(), 7)
        child = json.loads((self.root / "child.json").read_text(encoding="utf-8"))
        self.assertEqual(child["GDAL_DATA"], "child Δ")
        self.assertTrue(child["PATH"].startswith(str(self.root) + ";"))
        self.assertEqual(child["PROJ_NETWORK"], "OFF")
        self.assertEqual(child["PYTHONIOENCODING"], "utf-8")
        self.assertEqual(dict(os.environ), before)
        self.assertEqual((self.root / "stdout.log").read_text(encoding="utf-8"), "Δ" * 100000)
        self.assertEqual((self.root / "stderr.log").read_text(encoding="utf-8"), "é" * 100000)
        self.assertTrue(self.status().startswith("failed"))

    def test_environment_allowlist_rejects_before_start(self):
        self.fake_worker("from pathlib import Path; Path('started').touch()")
        environment = self.root / "environment.json"
        environment.write_text(json.dumps({"UNSUPPORTED_KEY": "do not apply"}), encoding="utf-8")
        self.configure(environment_file=str(environment))
        self.assertEqual(self.run_launcher(), 2)
        self.assertIn("Unsupported child environment key", self.status())
        self.assertFalse((self.root / "started").exists())

    def test_cooperative_cancellation_preserves_worker_status(self):
        self.fake_worker("""from pathlib import Path
import time, sys
Path('ready').touch()
deadline = time.monotonic() + 15
while not Path('cancel.flag').exists():
    if time.monotonic() > deadline: raise RuntimeError('No cancellation received')
    time.sleep(0.05)
Path('status.txt').write_text('cancelled\\r\\nOriginals retained', encoding='utf-16')
sys.exit(130)
""")
        self.configure()
        process = subprocess.Popen([str(self.launcher), str(self.request)])
        try:
            deadline = time.monotonic() + 10
            while not (self.root / "ready").exists() and process.poll() is None and time.monotonic() < deadline:
                time.sleep(0.05)
            self.assertTrue((self.root / "ready").exists(), "Child failed to start.")
            (self.root / "cancel.flag").write_text("cancel", encoding="utf-8")
            self.assertEqual(process.wait(timeout=20), 130)
            self.assertTrue(self.status().startswith("cancelled"))
        finally:
            if process.poll() is None:
                (self.root / "cancel.flag").touch()
                process.wait(timeout=20)


if __name__ == "__main__":
    unittest.main()
