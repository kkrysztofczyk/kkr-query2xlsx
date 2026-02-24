import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import pytest


class TestCliBasicIntegration:
    """Fast unit-level checks for CLI test fixtures."""

    def test_main_script_is_present(self):
        repo_root = Path(__file__).resolve().parents[1]
        assert (repo_root / "main.pyw").exists()


@pytest.mark.slow
class TestCliSubprocessContract:
    """Slow subprocess-based CLI contract tests."""

    @classmethod
    def setup_class(cls):
        cls.repo_root = Path(__file__).resolve().parents[1]
        cls.main_path = cls.repo_root / "main.pyw"

    def _run_cli(self, args, *, data_home: Path, app_home: Path):
        env = os.environ.copy()
        env["XDG_DATA_HOME"] = str(data_home)
        env["KKR_LANG"] = "en"
        env["LC_ALL"] = "C"
        env["LANG"] = "C"
        env["PYTHONUTF8"] = "1"

        run_main = app_home / "main.pyw"
        shutil.copy2(self.main_path, run_main)
        cli_args = [*args, "--lang", "en"]

        return subprocess.run(
            [sys.executable, str(run_main), *cli_args],
            cwd=str(app_home),
            capture_output=True,
            text=True,
            env=env,
        )

    def test_list_connections_output_format(self):
        with tempfile.TemporaryDirectory() as td:
            data_home = Path(td)
            app_home = data_home / "app"
            app_home.mkdir(parents=True, exist_ok=True)
            (app_home / "secure.txt").write_text(
                json.dumps(
                    {
                        "connections": [
                            {"name": "Default MSSQL", "type": "mssql_odbc", "details": {}},
                            {"name": "Demo SQLite", "type": "sqlite", "details": {"path": "demo.db"}},
                        ],
                        "last_selected": "Default MSSQL",
                    }
                ),
                encoding="utf-8",
            )

            proc = self._run_cli(["--list-connections"], data_home=data_home, app_home=app_home)

            assert proc.returncode == 0, proc.stderr
            assert proc.stdout.strip().splitlines() == ["Default MSSQL", "Demo SQLite"]

    def test_help_message_accessibility(self):
        with tempfile.TemporaryDirectory() as td:
            data_home = Path(td)
            app_home = data_home / "app"
            app_home.mkdir(parents=True, exist_ok=True)

            proc = self._run_cli(["--help"], data_home=data_home, app_home=app_home)

            assert proc.returncode == 0, proc.stderr
            help_text = (proc.stdout + "\n" + proc.stderr).lower()
            assert "usage" in help_text
            assert "--list-connections" in help_text
            assert "--self-test" in help_text

    def test_self_test_exit_code_contract(self):
        with tempfile.TemporaryDirectory() as td:
            data_home = Path(td)
            app_home = data_home / "app"
            app_home.mkdir(parents=True, exist_ok=True)

            proc = self._run_cli(["--self-test"], data_home=data_home, app_home=app_home)

            assert proc.returncode == 0, proc.stderr
