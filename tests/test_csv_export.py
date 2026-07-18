import codecs
import csv
import datetime
import importlib.machinery
import importlib.util
import tempfile
import threading
import unittest
from pathlib import Path


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_csv_export", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_csv_export", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class CsvExportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def _export(self, columns, rows, profile_overrides=None):
        profile = dict(self.app.DEFAULT_CSV_PROFILE)
        if profile_overrides:
            profile.update(profile_overrides)
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
            path = f.name
        cancel = threading.Event()
        self.app._export_rows_to_csv(
            output_file_path=path,
            columns=columns,
            rows=rows,
            profile=profile,
            timeout_seconds=30,
            cancel_event=cancel,
        )
        return Path(path)

    # --- BOM ---

    def test_bom_enabled_file_starts_with_bom(self):
        path = self._export(["a"], [["hello"]], {"bom": True, "encoding": "utf-8"})
        raw = path.read_bytes()
        self.assertTrue(raw.startswith(codecs.BOM_UTF8), msg=f"BOM missing: {raw[:10]!r}")

    def test_bom_disabled_file_does_not_start_with_bom(self):
        path = self._export(["a"], [["hello"]], {"bom": False, "encoding": "utf-8"})
        raw = path.read_bytes()
        self.assertFalse(raw.startswith(codecs.BOM_UTF8), msg=f"Unexpected BOM: {raw[:10]!r}")

    # --- datetime with combined date_format + time_format ---

    def test_datetime_combined_date_and_time_format(self):
        dt = datetime.datetime(2024, 3, 15, 14, 30, 5)
        path = self._export(
            ["ts"],
            [[dt]],
            {"date_format": "%d.%m.%Y", "time_format": "%H:%M"},
        )
        rows = list(csv.reader(path.read_text(encoding="utf-8").splitlines()))
        self.assertEqual(rows[1][0], "15.03.2024 14:30")

    def test_datetime_date_format_only_covers_full_format(self):
        dt = datetime.datetime(2024, 3, 15, 9, 5, 0)
        path = self._export(
            ["ts"],
            [[dt]],
            {"date_format": "%Y-%m-%d %H:%M:%S"},
        )
        rows = list(csv.reader(path.read_text(encoding="utf-8").splitlines()))
        self.assertEqual(rows[1][0], "2024-03-15 09:05:00")

    # --- detect_date_only ---

    def test_detect_date_only_strips_midnight_time(self):
        rows_data = [
            [datetime.datetime(2024, 1, 1, 0, 0, 0)],
            [datetime.datetime(2024, 6, 15, 0, 0, 0)],
        ]
        path = self._export(
            ["ts"], rows_data, {"detect_date_only": True, "date_format": "%Y-%m-%d"}
        )
        content = list(csv.reader(path.read_text(encoding="utf-8").splitlines()))
        self.assertEqual(content[1][0], "2024-01-01")
        self.assertEqual(content[2][0], "2024-06-15")

    def test_detect_date_only_preserves_non_midnight_time(self):
        rows_data = [
            [datetime.datetime(2024, 1, 1, 0, 0, 0)],
            [datetime.datetime(2024, 6, 15, 9, 30, 0)],
        ]
        path = self._export(
            ["ts"], rows_data, {"detect_date_only": True}
        )
        content = list(csv.reader(path.read_text(encoding="utf-8").splitlines()))
        self.assertIn("09:30", content[2][0])

    def test_detect_date_only_per_column_independence(self):
        """First col all-midnight (→ date only via isoformat), second col has times (→ full)."""
        rows_data = [
            [datetime.datetime(2024, 1, 1, 0, 0, 0), datetime.datetime(2024, 1, 1, 8, 0, 0)],
            [datetime.datetime(2024, 2, 1, 0, 0, 0), datetime.datetime(2024, 2, 1, 9, 15, 0)],
        ]
        # No date_format: midnight col → date isoformat, non-midnight col → full isoformat
        path = self._export(
            ["date_col", "ts_col"],
            rows_data,
            {"detect_date_only": True, "date_format": ""},
        )
        content = list(csv.reader(path.read_text(encoding="utf-8").splitlines()))
        self.assertEqual(content[1][0], "2024-01-01", msg="date_col should be date-only")
        self.assertIn("08:00", content[1][1], msg="ts_col should keep time")

    # --- Row-width policy: raise on extra fields ---

    def test_row_wider_than_columns_raises_value_error(self):
        """Rows with more values than declared columns must raise ValueError."""
        rows_data = [["a", "b", "EXTRA_VALUE"]]
        with self.assertRaises(ValueError) as ctx:
            self._export(["col1", "col2"], rows_data)
        self.assertIn("mismatch", str(ctx.exception).lower())

    def test_row_narrower_than_columns_raises_value_error(self):
        """Rows with fewer values than declared columns must raise ValueError (no silent empty fields)."""
        rows_data = [[1]]
        with self.assertRaises(ValueError) as ctx:
            self._export(["a", "b"], rows_data)
        self.assertIn("mismatch", str(ctx.exception).lower())


class ValidateTimeFormatTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()
        cls.app.set_lang("en")

    def test_empty_time_format_returns_true_with_default_text(self):
        valid, preview = self.app._validate_time_format("")
        self.assertTrue(valid)
        self.assertIn("ISO", preview)

    def test_valid_time_format_returns_true_with_preview(self):
        valid, preview = self.app._validate_time_format("%H:%M")
        self.assertTrue(valid)
        self.assertIn(":", preview)

    def test_invalid_time_format_returns_false(self):
        valid, preview = self.app._validate_time_format("%Q")
        self.assertFalse(valid)
        self.assertIn("Invalid", preview)

    def test_invalid_date_format_unknown_directive(self):
        valid, _ = self.app._validate_date_format("%Q")
        self.assertFalse(valid)

    def test_percent_c_locale_datetime_is_valid(self):
        """%c is a standard strftime code and must not be rejected."""
        valid, _ = self.app._validate_date_format("%c")
        self.assertTrue(valid)

    def test_percent_c_valid_in_time_format(self):
        valid, _ = self.app._validate_time_format("%c")
        self.assertTrue(valid)

    def test_percent_at_end_of_format_is_invalid(self):
        valid, _ = self.app._validate_time_format("%H:%")
        self.assertFalse(valid)

    def test_escaped_percent_is_valid(self):
        valid, _ = self.app._validate_time_format("%%H")
        self.assertTrue(valid)

    def test_has_unknown_strftime_directive_known(self):
        self.assertFalse(self.app._has_unknown_strftime_directive("%Y-%m-%d"))

    def test_has_unknown_strftime_directive_unknown(self):
        self.assertTrue(self.app._has_unknown_strftime_directive("%Q"))

    def test_has_unknown_strftime_directive_trailing_percent(self):
        self.assertTrue(self.app._has_unknown_strftime_directive("%H%"))

    def test_platform_specific_century_directive_is_rejected(self):
        self.assertTrue(self.app._has_unknown_strftime_directive("%C"))


if __name__ == "__main__":
    unittest.main()
