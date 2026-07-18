import importlib.machinery
import importlib.util
import sqlite3
import tempfile
import threading
import unittest
from contextlib import closing
from pathlib import Path
from unittest import mock

from sqlalchemy import create_engine
from sqlalchemy.pool import NullPool


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_sqlite_export", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_sqlite_export", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


def _fetchall_sqlite(db_path: Path, sql: str):
    with closing(sqlite3.connect(db_path)) as conn:
        cur = conn.cursor()
        try:
            return cur.execute(sql).fetchall()
        finally:
            cur.close()


class SQLiteExportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def test_sqlite_safe_table_name_sanitizes_input(self):
        self.assertEqual(self.app._sqlite_safe_table_name("Sales Report 2026!"), "Sales_Report_2026")
        self.assertEqual(self.app._sqlite_safe_table_name("123 bad name"), "t_123_bad_name")
        self.assertEqual(self.app._sqlite_safe_table_name(""), "results")

    def test_normalize_output_file_path_rewrites_invalid_sqlite_extension(self):
        with tempfile.TemporaryDirectory() as td:
            output_path, ext_mismatch = self.app.normalize_output_file_path(
                output_directory=td,
                default_file_name="report.sqlite",
                output_format="sqlite",
                override_path=str(Path(td) / "report.txt"),
            )

            self.assertTrue(ext_mismatch)
            self.assertEqual(Path(output_path).suffix, ".sqlite")

    def test_export_rows_to_sqlite_replace_overwrites_existing_table(self):
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "replace.sqlite"
            with closing(sqlite3.connect(db_path)) as conn:
                conn.execute('CREATE TABLE "demo_table" ("id" INTEGER, "name" TEXT)')
                conn.execute('INSERT INTO "demo_table" VALUES (99, "old")')
                conn.commit()

            self.app._export_rows_to_sqlite(
                str(db_path),
                table_name="demo_table",
                write_mode="replace",
                columns=["id", "name"],
                rows=[(1, "Alice"), (2, "Bob")],
                timeout_seconds=0,
            )

            rows = _fetchall_sqlite(db_path, 'SELECT id, name FROM "demo_table" ORDER BY id')
            self.assertEqual(rows, [(1, "Alice"), (2, "Bob")])

    def test_export_rows_to_sqlite_append_to_existing_matching_table(self):
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "append.sqlite"
            with closing(sqlite3.connect(db_path)) as conn:
                conn.execute('CREATE TABLE "demo_table" ("id" INTEGER, "name" TEXT)')
                conn.execute('INSERT INTO "demo_table" VALUES (1, "Alice")')
                conn.commit()

            self.app._export_rows_to_sqlite(
                str(db_path),
                table_name="demo_table",
                write_mode="append",
                columns=["id", "name"],
                rows=[(2, "Bob")],
                timeout_seconds=0,
            )

            rows = _fetchall_sqlite(db_path, 'SELECT id, name FROM "demo_table" ORDER BY id')
            self.assertEqual(rows, [(1, "Alice"), (2, "Bob")])

    def test_export_rows_to_sqlite_append_schema_mismatch_raises(self):
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "mismatch.sqlite"
            with closing(sqlite3.connect(db_path)) as conn:
                conn.execute('CREATE TABLE "demo_table" ("id" INTEGER, "different_col" TEXT)')
                conn.commit()

            with self.assertRaises(ValueError):
                self.app._export_rows_to_sqlite(
                    str(db_path),
                    table_name="demo_table",
                    write_mode="append",
                    columns=["id", "name"],
                    rows=[(2, "Bob")],
                    timeout_seconds=0,
                )

    def test_export_rows_to_sqlite_row_width_mismatch_raises_clear_error(self):
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "width_mismatch.sqlite"

            with self.assertRaisesRegex(
                ValueError,
                r"SQLite export row width mismatch: expected 2 values, got 3\.",
            ):
                self.app._export_rows_to_sqlite(
                    str(db_path),
                    table_name="demo_table",
                    write_mode="replace",
                    columns=["id", "name"],
                    rows=[(1, "Alice", "EXTRA")],
                    timeout_seconds=0,
                )

    def test_run_export_sqlite_replace_with_zero_rows_clears_existing_table(self):
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "replace_empty.sqlite"
            with closing(sqlite3.connect(db_path)) as conn:
                conn.execute('CREATE TABLE "demo_table" ("id" INTEGER, "name" TEXT)')
                conn.execute('INSERT INTO "demo_table" VALUES (1, "Alice")')
                conn.commit()

            with mock.patch.object(
                self.app,
                "_run_query_to_rows",
                return_value=([], ["id", "name"], 0.2, 20.0),
            ):
                sql_dur, export_dur, total_dur, rows_count = self.app.run_export(
                    engine=object(),
                    sql_query="SELECT id, name FROM demo WHERE 1=0",
                    output_file_path=str(db_path),
                    output_format="sqlite",
                    sqlite_table="demo_table",
                    sqlite_mode="replace",
                )

            self.assertEqual(rows_count, 0)
            self.assertGreaterEqual(export_dur, 0.0)
            rows = _fetchall_sqlite(db_path, 'SELECT id, name FROM "demo_table"')
            self.assertEqual(rows, [])

    def test_run_console_noninteractive_sqlite_creates_db_and_table(self):
        engine = create_engine("sqlite:///:memory:", poolclass=NullPool)
        try:
            with tempfile.TemporaryDirectory() as td:
                td_path = Path(td)
                sql_path = td_path / "sample.sql"
                sql_path.write_text("SELECT 1 AS id, 'Alice' AS name;", encoding="utf-8")
                output_path = td_path / "output.sqlite"

                exit_code = self.app.run_console_noninteractive(
                    engine,
                    output_directory=str(td_path),
                    selected_connection={"name": "E2E", "type": "sqlite"},
                    sql_path=str(sql_path),
                    output_format="sqlite",
                    output_override=str(output_path),
                    archive_sql=False,
                    sqlite_table="sales report",
                    sqlite_mode="replace",
                )

                self.assertEqual(exit_code, 0)
                rows = _fetchall_sqlite(output_path, 'SELECT id, name FROM "sales_report"')
                self.assertEqual(rows, [(1, "Alice")])
        finally:
            engine.dispose()

    def test_run_export_passes_cancel_event_to_sqlite_export_phase(self):
        cancel_event = threading.Event()

        with mock.patch.object(
            self.app,
            "_run_query_to_rows",
            return_value=([(1,)], ["id"], 0.2, 20.0),
        ), mock.patch.object(self.app, "_export_rows_to_sqlite") as mock_export_sqlite:
            with tempfile.NamedTemporaryFile(suffix=".sqlite") as tmp:
                self.app.run_export(
                    engine=object(),
                    sql_query="SELECT 1",
                    output_file_path=tmp.name,
                    output_format="sqlite",
                    sqlite_table="demo_table",
                    sqlite_mode="append",
                    cancel_event=cancel_event,
                )

        self.assertTrue(mock_export_sqlite.called)
        self.assertIs(mock_export_sqlite.call_args.kwargs["cancel_event"], cancel_event)

    def test_run_export_sqlite_error_preserves_preexisting_database_file(self):
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "existing.sqlite"
            with closing(sqlite3.connect(db_path)) as conn:
                conn.execute('CREATE TABLE "keep_me" ("id" INTEGER)')
                conn.execute('INSERT INTO "keep_me" VALUES (1)')
                conn.commit()

            with mock.patch.object(
                self.app,
                "_run_query_to_rows",
                return_value=([(2,)], ["id"], 0.2, 20.0),
            ), mock.patch.object(
                self.app,
                "_export_rows_to_sqlite",
                side_effect=ValueError("sqlite append failed"),
            ):
                with self.assertRaisesRegex(ValueError, "sqlite append failed"):
                    self.app.run_export(
                        engine=object(),
                        sql_query="SELECT 2 AS id",
                        output_file_path=str(db_path),
                        output_format="sqlite",
                        sqlite_table="keep_me",
                        sqlite_mode="append",
                    )

            self.assertTrue(db_path.exists())
            rows = _fetchall_sqlite(db_path, 'SELECT id FROM "keep_me"')
            self.assertEqual(rows, [(1,)])

    def test_run_export_sqlite_error_removes_new_database_file(self):
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "new.sqlite"
            self.assertFalse(db_path.exists())

            def failing_sqlite_export(output_file_path, **_kwargs):
                Path(output_file_path).write_bytes(b"sqlite-partial")
                raise RuntimeError("sqlite export failed")

            with mock.patch.object(
                self.app,
                "_run_query_to_rows",
                return_value=([(1,)], ["id"], 0.2, 20.0),
            ), mock.patch.object(
                self.app,
                "_export_rows_to_sqlite",
                side_effect=failing_sqlite_export,
            ):
                with self.assertRaisesRegex(RuntimeError, "sqlite export failed"):
                    self.app.run_export(
                        engine=object(),
                        sql_query="SELECT 1 AS id",
                        output_file_path=str(db_path),
                        output_format="sqlite",
                        sqlite_table="results",
                        sqlite_mode="replace",
                    )

            self.assertFalse(db_path.exists())


if __name__ == "__main__":
    unittest.main()
