import decimal
import importlib.machinery
import importlib.util
import io
import sqlite3
import tempfile
import threading
import unittest
from contextlib import closing, redirect_stdout
from pathlib import Path


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_decimal_coerce", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_decimal_coerce", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class DecimalPrecisionTests(unittest.TestCase):
    """Regression tests for Decimal → SQLite precision handling."""

    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def test_coerce_decimal_fractional_preserved_as_text(self):
        value = decimal.Decimal("12345678901234567890.12")
        result = self.app._coerce_sqlite_value(value)
        self.assertIsInstance(result, str)
        self.assertEqual(result, "12345678901234567890.12")

    def test_coerce_decimal_big_integer_preserved_as_text(self):
        value = decimal.Decimal("9223372036854775808")  # INT64_MAX + 1
        result = self.app._coerce_sqlite_value(value)
        self.assertIsInstance(result, str)
        self.assertEqual(result, "9223372036854775808")

    def test_coerce_decimal_small_integer_stored_as_text(self):
        """Even small integral Decimal is TEXT — consistent with all-Decimal-as-TEXT rule."""
        value = decimal.Decimal("42")
        result = self.app._coerce_sqlite_value(value)
        self.assertIsInstance(result, str)
        self.assertEqual(result, "42")

    def test_decimal_roundtrip_in_sqlite_via_export(self):
        """Full roundtrip: Decimal values reach SQLite without precision loss."""
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "dec_test.sqlite"
            columns = ["small_int", "big_int", "fractional"]
            rows = [
                [
                    decimal.Decimal("42"),
                    decimal.Decimal("9223372036854775808"),
                    decimal.Decimal("12345678901234567890.12"),
                ]
            ]
            cancel = threading.Event()
            self.app._export_rows_to_sqlite(
                output_file_path=str(db_path),
                table_name="dec_test",
                write_mode="replace",
                columns=columns,
                rows=rows,
                timeout_seconds=30,
                cancel_event=cancel,
            )
            with closing(sqlite3.connect(db_path)) as conn:
                row = conn.execute(
                    "SELECT small_int, big_int, fractional FROM dec_test"
                ).fetchone()
        self.assertEqual(row[0], "42")
        self.assertEqual(row[1], "9223372036854775808")
        self.assertEqual(row[2], "12345678901234567890.12")


class SqliteColumnNamesTests(unittest.TestCase):
    """Regression tests: column names must be preserved, not sanitized."""

    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def test_original_names_preserved(self):
        cols = self.app._sqlite_safe_column_names(["Order ID", "Total (€)", "Order-ID"])
        self.assertEqual(cols, ["Order ID", "Total (€)", "Order-ID"])

    def test_true_duplicates_get_suffix(self):
        cols = self.app._sqlite_safe_column_names(["a", "a", "b", "a"])
        self.assertEqual(cols, ["a", "a_2", "b", "a_3"])

    def test_dedup_avoids_collision_with_existing_name(self):
        """['a', 'a_2', 'a'] must not produce duplicate 'a_2'."""
        cols = self.app._sqlite_safe_column_names(["a", "a_2", "a"])
        self.assertEqual(len(set(c.lower() for c in cols)), 3, msg=f"Duplicate in: {cols}")
        self.assertEqual(cols[0], "a")
        self.assertEqual(cols[1], "a_2")
        self.assertNotEqual(cols[2], "a_2")

    def test_dedup_case_insensitive(self):
        """SQLite column names are case-insensitive; 'a' and 'A' must be deduplicated."""
        cols = self.app._sqlite_safe_column_names(["a", "A"])
        self.assertEqual(len(set(c.lower() for c in cols)), 2, msg=f"CI duplicate: {cols}")
        self.assertNotEqual(cols[0].lower(), cols[1].lower())

    def test_dedup_a_a_a2_no_collision(self):
        """['a', 'a', 'a_2'] must produce 3 unique names."""
        cols = self.app._sqlite_safe_column_names(["a", "a", "a_2"])
        self.assertEqual(len(set(c.lower() for c in cols)), 3, msg=f"Duplicate in: {cols}")

    def test_blank_name_filled(self):
        cols = self.app._sqlite_safe_column_names(["", None, "val"])
        self.assertEqual(cols, ["column_1", "column_2", "val"])

    def test_original_names_roundtrip_in_sqlite(self):
        """Columns with spaces and special chars survive a SQLite insert/select roundtrip."""
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "cols_test.sqlite"
            columns = ["Order ID", "Total (€)"]
            rows = [["ABC-001", "99.99"]]
            cancel = threading.Event()
            self.app._export_rows_to_sqlite(
                output_file_path=str(db_path),
                table_name="orders",
                write_mode="replace",
                columns=columns,
                rows=rows,
                timeout_seconds=30,
                cancel_event=cancel,
            )
            with closing(sqlite3.connect(db_path)) as conn:
                cursor = conn.execute('SELECT "Order ID", "Total (€)" FROM orders')
                row = cursor.fetchone()
                col_names = [d[0] for d in cursor.description]
        self.assertEqual(col_names, ["Order ID", "Total (€)"])
        self.assertEqual(row, ("ABC-001", "99.99"))


class CoerceCsvValueTests(unittest.TestCase):
    """Tests for _coerce_csv_value with time_format and detect_date_only."""

    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()
        import datetime as _dt
        cls.dt = _dt

    def _coerce(self, value, date_format=None, **kwargs):
        return self.app._coerce_csv_value(value, decimal_sep=".", date_format=date_format, **kwargs)

    def test_time_value_uses_time_format(self):
        t = self.dt.time(14, 30, 5)
        result = self._coerce(t, time_format="%H:%M")
        self.assertEqual(result, "14:30")

    def test_time_value_defaults_to_isoformat(self):
        t = self.dt.time(9, 5, 0)
        result = self._coerce(t)
        self.assertEqual(result, "09:05:00")

    def test_detect_date_only_midnight_datetime(self):
        """Midnight datetime with detect_date_only → date isoformat only."""
        dt = self.dt.datetime(2024, 3, 15, 0, 0, 0)
        result = self._coerce(dt, detect_date_only=True)
        self.assertEqual(result, "2024-03-15")

    def test_detect_date_only_nonmidnight_not_affected(self):
        """Non-midnight datetime with detect_date_only → full datetime preserved."""
        dt = self.dt.datetime(2024, 3, 15, 12, 30, 0)
        result = self._coerce(dt, detect_date_only=True)
        self.assertIn("12:30", result)

    def test_detect_date_only_false_keeps_datetime_at_midnight(self):
        """When detect_date_only=False, midnight datetime is kept as full datetime."""
        dt = self.dt.datetime(2024, 3, 15, 0, 0, 0)
        result = self._coerce(dt, detect_date_only=False)
        self.assertIn("00:00:00", result)

    def test_datetime_uses_combined_date_and_time_format(self):
        """When both date_format and time_format are set, datetime uses both."""
        dt = self.dt.datetime(2024, 3, 15, 14, 30, 5)
        result = self._coerce(dt, date_format="%d.%m.%Y", time_format="%H:%M")
        self.assertEqual(result, "15.03.2024 14:30")

    def test_datetime_time_format_only_falls_back_gracefully(self):
        """When only time_format is set (no date_format), isoformat date + time_format time."""
        dt = self.dt.datetime(2024, 3, 15, 14, 30, 5)
        result = self._coerce(dt, time_format="%H:%M")
        self.assertIn("14:30", result)


class PrintConsoleExportResultTests(unittest.TestCase):
    """Tests for _print_console_export_result — ensures correct single message per scenario."""

    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()
        cls.app.set_lang("en")

    def _capture(self, **kwargs):
        buf = io.StringIO()
        with redirect_stdout(buf):
            self.app._print_console_export_result(**kwargs)
        return buf.getvalue()

    def test_sqlite_replace_zero_rows_prints_cleared_only(self):
        output = self._capture(
            output_file_path="/tmp/x.sqlite",
            rows_count=0,
            existed_before=True,
            output_format="sqlite",
            sqlite_table="demo_table",
            sqlite_mode="replace",
        )
        self.assertIn("recreated", output.lower())
        self.assertNotIn("nothing saved", output.lower())
        self.assertNotIn("no rows", output.lower())

    def test_sqlite_replace_zero_rows_no_double_message(self):
        output = self._capture(
            output_file_path="/tmp/x.sqlite",
            rows_count=0,
            existed_before=True,
            output_format="sqlite",
            sqlite_table="demo_table",
            sqlite_mode="replace",
        )
        self.assertEqual(output.count("\n"), 1, msg=f"Expected single line: {output!r}")

    def test_non_sqlite_zero_rows_prints_nothing_saved(self):
        output = self._capture(
            output_file_path="/tmp/x.xlsx",
            rows_count=0,
            existed_before=True,
            output_format="xlsx",
        )
        self.assertIn("nothing saved", output.lower())
        self.assertNotIn("recreated", output.lower())

    def test_positive_rows_prints_saved_path(self):
        output = self._capture(
            output_file_path="/tmp/result.xlsx",
            rows_count=5,
            existed_before=False,
            output_format="xlsx",
        )
        self.assertIn("result.xlsx", output)


class AppendPrecisionGuardTests(unittest.TestCase):
    """Regression tests: append to typed numeric columns must raise, not silently lose precision."""

    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def _make_db_with_integer_column(self, db_path: Path):
        """Create a SQLite table with an INTEGER-typed column (simulates old-style export)."""
        with closing(sqlite3.connect(db_path)) as conn:
            conn.execute('CREATE TABLE results ("id" INTEGER, "val" TEXT)')
            conn.execute('INSERT INTO results VALUES (1, "hello")')
            conn.commit()

    def test_append_large_int_to_integer_column_raises(self):
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "typed.sqlite"
            self._make_db_with_integer_column(db_path)

            cancel = threading.Event()
            with self.assertRaises(ValueError) as ctx:
                self.app._export_rows_to_sqlite(
                    output_file_path=str(db_path),
                    table_name="results",
                    write_mode="append",
                    columns=["id", "val"],
                    rows=[[9223372036854775808, "world"]],
                    timeout_seconds=30,
                    cancel_event=cancel,
                )
            self.assertIn("precision", str(ctx.exception).lower())
            self.assertIn("replace", str(ctx.exception).lower())

    def test_append_decimal_to_integer_column_raises(self):
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "typed2.sqlite"
            self._make_db_with_integer_column(db_path)

            cancel = threading.Event()
            with self.assertRaises(ValueError) as ctx:
                self.app._export_rows_to_sqlite(
                    output_file_path=str(db_path),
                    table_name="results",
                    write_mode="append",
                    columns=["id", "val"],
                    rows=[[decimal.Decimal("12345678901234567890.12"), "world"]],
                    timeout_seconds=30,
                    cancel_event=cancel,
                )
            self.assertIn("precision", str(ctx.exception).lower())

    def test_append_small_int_to_integer_column_succeeds(self):
        """Normal small ints should still append without error."""
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "typed3.sqlite"
            self._make_db_with_integer_column(db_path)

            cancel = threading.Event()
            self.app._export_rows_to_sqlite(
                output_file_path=str(db_path),
                table_name="results",
                write_mode="append",
                columns=["id", "val"],
                rows=[[42, "ok"]],
                timeout_seconds=30,
                cancel_event=cancel,
            )
            with closing(sqlite3.connect(db_path)) as conn:
                count = conn.execute("SELECT COUNT(*) FROM results").fetchone()[0]
            self.assertEqual(count, 2)

    def test_append_large_int_in_second_row_raises(self):
        """Guard must scan ALL rows: first row small, second row large int must still raise."""
        with tempfile.TemporaryDirectory() as td:
            db_path = Path(td) / "typed4.sqlite"
            self._make_db_with_integer_column(db_path)

            cancel = threading.Event()
            with self.assertRaises(ValueError) as ctx:
                self.app._export_rows_to_sqlite(
                    output_file_path=str(db_path),
                    table_name="results",
                    write_mode="append",
                    columns=["id", "val"],
                    rows=[[1, "first"], [9223372036854775808, "second"]],
                    timeout_seconds=30,
                    cancel_event=cancel,
                )
            self.assertIn("precision", str(ctx.exception).lower())

    def test_sqlite_column_affinity_rules(self):
        """_sqlite_column_affinity must follow official 5-rule SQLite algorithm."""
        affinity = self.app._sqlite_column_affinity
        self.assertEqual(affinity("INTEGER"), "INTEGER")
        self.assertEqual(affinity("BIGINT"), "INTEGER")
        self.assertEqual(affinity("VARCHAR(255)"), "TEXT")
        self.assertEqual(affinity("CLOB"), "TEXT")
        self.assertEqual(affinity("BLOB"), "BLOB")
        self.assertEqual(affinity(""), "BLOB")
        self.assertEqual(affinity("REAL"), "REAL")
        self.assertEqual(affinity("FLOAT"), "REAL")
        self.assertEqual(affinity("DOUBLE PRECISION"), "REAL")
        self.assertEqual(affinity("NUMERIC"), "NUMERIC")
        self.assertEqual(affinity("DECIMAL(10,2)"), "NUMERIC")
        self.assertEqual(affinity("DATE"), "NUMERIC")
        self.assertEqual(affinity("BOOLEAN"), "NUMERIC")


class DetectDateOnlyColumnsTests(unittest.TestCase):
    """Tests for per-column detect_date_only logic."""

    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()
        import datetime as _dt
        cls.dt = _dt

    def test_all_midnight_column_detected(self):
        rows = [
            [self.dt.datetime(2024, 1, 1, 0, 0, 0)],
            [self.dt.datetime(2024, 6, 15, 0, 0, 0)],
        ]
        flags = self.app._detect_date_only_columns(["ts"], rows)
        self.assertEqual(flags, [True])

    def test_mixed_times_column_not_detected(self):
        rows = [
            [self.dt.datetime(2024, 1, 1, 0, 0, 0)],
            [self.dt.datetime(2024, 6, 15, 9, 30, 0)],
        ]
        flags = self.app._detect_date_only_columns(["ts"], rows)
        self.assertEqual(flags, [False])

    def test_per_column_independence(self):
        """First col all-midnight, second col has times → independent flags."""
        rows = [
            [self.dt.datetime(2024, 1, 1, 0, 0, 0), self.dt.datetime(2024, 1, 1, 8, 0, 0)],
            [self.dt.datetime(2024, 2, 1, 0, 0, 0), self.dt.datetime(2024, 2, 1, 9, 15, 0)],
        ]
        flags = self.app._detect_date_only_columns(["date_col", "ts_col"], rows)
        self.assertEqual(flags, [True, False])


if __name__ == "__main__":
    unittest.main()
