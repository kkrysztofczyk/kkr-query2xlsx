import json
import os
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import unittest
from contextlib import closing
from pathlib import Path


class TestCliBasicIntegration(unittest.TestCase):
    """Fast unit-level checks for CLI test fixtures."""

    def test_main_script_is_present(self):
        repo_root = Path(__file__).resolve().parents[1]
        self.assertTrue((repo_root / "main.pyw").exists())


class TestCliSubprocessContract(unittest.TestCase):
    """Slow subprocess-based CLI contract tests (unittest only, no pytest marker)."""

    @classmethod
    def setUpClass(cls):
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

            self.assertEqual(proc.returncode, 0, msg=proc.stderr)
            self.assertEqual(proc.stdout.strip().splitlines(), ["Default MSSQL", "Demo SQLite"])
            stderr = proc.stderr.lower()
            self.assertNotIn("tkinter", stderr)
            self.assertNotIn("traceback", stderr)
            self.assertNotIn("gui", stderr)

    def test_help_message_accessibility(self):
        with tempfile.TemporaryDirectory() as td:
            data_home = Path(td)
            app_home = data_home / "app"
            app_home.mkdir(parents=True, exist_ok=True)

            proc = self._run_cli(["--help"], data_home=data_home, app_home=app_home)

            self.assertEqual(proc.returncode, 0, msg=proc.stderr)
            help_text = (proc.stdout + "\n" + proc.stderr).lower()
            self.assertIn("usage", help_text)
            self.assertIn("--list-connections", help_text)
            self.assertIn("--self-test", help_text)

    def test_self_test_exit_code_contract(self):
        with tempfile.TemporaryDirectory() as td:
            data_home = Path(td)
            app_home = data_home / "app"
            app_home.mkdir(parents=True, exist_ok=True)

            proc = self._run_cli(["--self-test"], data_home=data_home, app_home=app_home)

            self.assertEqual(proc.returncode, 0, msg=proc.stderr)

    def test_sql_requires_format_contract(self):
        with tempfile.TemporaryDirectory() as td:
            data_home = Path(td)
            app_home = data_home / "app"
            app_home.mkdir(parents=True, exist_ok=True)

            sql_path = app_home / "report.sql"
            sql_path.write_text("SELECT 1 AS id;", encoding="utf-8")

            proc = self._run_cli(["--sql", str(sql_path)], data_home=data_home, app_home=app_home)

            combined = (proc.stdout + "\n" + proc.stderr).lower()
            self.assertEqual(proc.returncode, 2, msg=combined)
            self.assertIn("--format", combined)
            self.assertTrue("usage" in combined or "error" in combined, msg=combined)

    def test_noninteractive_sqlite_export_with_cli_flags(self):
        with tempfile.TemporaryDirectory() as td:
            data_home = Path(td)
            app_home = data_home / "app"
            app_home.mkdir(parents=True, exist_ok=True)

            db_path = app_home / "source.sqlite"
            sql_path = app_home / "report.sql"
            sql_path.write_text("SELECT 1 AS id, 'Alice' AS name;", encoding="utf-8")
            out_path = app_home / "sqlite_export.txt"

            (app_home / "secure.txt").write_text(
                json.dumps(
                    {
                        "connections": [
                            {
                                "name": "Demo SQLite",
                                "type": "sqlite",
                                "details": {"path": str(db_path)},
                            }
                        ],
                        "last_selected": "Demo SQLite",
                    }
                ),
                encoding="utf-8",
            )

            proc = self._run_cli(
                [
                    "--connection",
                    "Demo SQLite",
                    "--sql",
                    str(sql_path),
                    "--format",
                    "sqlite",
                    "--output",
                    str(out_path),
                    "--sqlite-table",
                    "Sales Report",
                    "--sqlite-mode",
                    "replace",
                ],
                data_home=data_home,
                app_home=app_home,
            )

            self.assertEqual(proc.returncode, 0, msg=proc.stderr)
            normalized_out = app_home / "sqlite_export.sqlite"
            self.assertTrue(normalized_out.exists(), msg=proc.stdout)

            with closing(sqlite3.connect(normalized_out)) as conn:
                rows = conn.execute('SELECT id, name FROM "Sales_Report"').fetchall()
            self.assertEqual(rows, [(1, "Alice")])

    def test_list_connections_prefers_app_dir_secure_txt(self):
        """App-dir secure.txt wins over user-dir secure.txt; wrong source must not appear."""
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            app_home = tmp / "app"
            app_home.mkdir(parents=True, exist_ok=True)
            user_data_dir = tmp / "userdata" / "kkr-query2xlsx"
            user_data_dir.mkdir(parents=True, exist_ok=True)

            (user_data_dir / "secure.txt").write_text(
                json.dumps({"connections": [{"name": "Wrong Source"}]}),
                encoding="utf-8",
            )
            (app_home / "secure.txt").write_text(
                json.dumps({"connections": [{"name": "App Source"}]}),
                encoding="utf-8",
            )

            env = os.environ.copy()
            env["XDG_DATA_HOME"] = str(tmp / "userdata")
            env["LOCALAPPDATA"] = str(tmp / "userdata")
            env["KKR_LANG"] = "en"
            env["LC_ALL"] = "C"
            env["LANG"] = "C"
            env["PYTHONUTF8"] = "1"

            run_main = app_home / "main.pyw"
            shutil.copy2(self.main_path, run_main)

            proc = subprocess.run(
                [sys.executable, str(run_main), "--list-connections", "--lang", "en"],
                cwd=str(app_home),
                capture_output=True,
                text=True,
                env=env,
            )

            self.assertEqual(proc.returncode, 0, msg=proc.stderr)
            lines = proc.stdout.strip().splitlines()
            self.assertIn("App Source", lines)
            self.assertNotIn("Wrong Source", lines)

    def test_list_connections_does_not_create_logs_dir(self):
        """--list-connections must not create logs/ or write any log file."""
        with tempfile.TemporaryDirectory() as td:
            data_home = Path(td)
            app_home = data_home / "app"
            app_home.mkdir(parents=True, exist_ok=True)
            (app_home / "secure.txt").write_text(
                json.dumps({"connections": [{"name": "Conn A"}]}),
                encoding="utf-8",
            )

            proc = self._run_cli(["--list-connections"], data_home=data_home, app_home=app_home)

            self.assertEqual(proc.returncode, 0, msg=proc.stderr)
            user_data_dir = data_home / "kkr-query2xlsx"
            for work_dir in ("generated_reports", "sql_archive", "templates", "queries", "logs"):
                self.assertFalse(
                    (app_home / work_dir).exists(),
                    msg=f"{work_dir}/ was created in app_home",
                )
                if user_data_dir.exists():
                    self.assertFalse(
                        (user_data_dir / work_dir).exists(),
                        msg=f"{work_dir}/ was created in user data dir",
                    )

    def test_connection_not_found_exits_two_and_shows_hint(self):
        with tempfile.TemporaryDirectory() as td:
            data_home = Path(td)
            app_home = data_home / "app"
            app_home.mkdir(parents=True, exist_ok=True)

            db_path = app_home / "source.sqlite"
            (app_home / "secure.txt").write_text(
                json.dumps(
                    {
                        "connections": [
                            {
                                "name": "Default MSSQL",
                                "type": "sqlite",
                                "details": {"path": str(db_path)},
                            }
                        ],
                        "last_selected": "Default MSSQL",
                    }
                ),
                encoding="utf-8",
            )

            proc = self._run_cli(
                ["--console", "--connection", "NOPE"],
                data_home=data_home,
                app_home=app_home,
            )

            combined = f"{proc.stdout}\n{proc.stderr}"
            self.assertEqual(proc.returncode, 2, msg=combined)
            self.assertIn("NOPE", combined)
            self.assertIn("--list-connections", combined)
            self.assertIn("Default MSSQL", combined)


if __name__ == "__main__":
    unittest.main()
