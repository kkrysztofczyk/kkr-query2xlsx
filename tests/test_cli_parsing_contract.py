import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class CliParsingContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.repo_root = Path(__file__).resolve().parents[1]
        cls.main_path = cls.repo_root / "main.pyw"

    def _run_cli(self, args):
        with tempfile.TemporaryDirectory() as td:
            app_home = Path(td) / "app"
            app_home.mkdir(parents=True, exist_ok=True)
            run_main = app_home / "main.pyw"
            shutil.copy2(self.main_path, run_main)

            env = os.environ.copy()
            env["XDG_DATA_HOME"] = str(Path(td) / "xdg-data")
            env["KKR_LANG"] = "en"
            env["LC_ALL"] = "C"
            env["LANG"] = "C"
            env["PYTHONUTF8"] = "1"

            cli_args = [*args, "--lang", "en"]
            return subprocess.run(
                [sys.executable, str(run_main), *cli_args],
                cwd=str(app_home),
                capture_output=True,
                text=True,
                env=env,
            )

    def test_sql_without_format_exits_two_and_shows_argparse_error(self):
        proc = self._run_cli(["--sql", "select 1"])
        output = f"{proc.stdout}\n{proc.stderr}".lower()

        self.assertEqual(proc.returncode, 2, output)
        self.assertIn("--format", output)
        self.assertTrue(
            any(token in output for token in ("error", "usage", "required")),
            output,
        )


if __name__ == "__main__":
    unittest.main()
