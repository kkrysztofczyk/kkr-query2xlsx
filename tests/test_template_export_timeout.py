import importlib.machinery
import importlib.util
import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_template_timeout", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_template_timeout", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class _FakeWorksheet:
    def cell(self, row, column):
        return type("Cell", (), {"value": None})()


class _FakeWorkbook:
    def __init__(self):
        self.sheetnames = ["Sheet1"]
        self._ws = _FakeWorksheet()

    def __getitem__(self, key):
        if key != "Sheet1":
            raise KeyError(key)
        return self._ws

    def save(self, path):
        return None

    def close(self):
        return None


class TemplateExportTimeoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def test_template_export_checks_deadline_even_when_no_rows(self):
        with tempfile.TemporaryDirectory() as td:
            template_path = Path(td) / "template.xlsx"
            template_path.write_bytes(b"")
            output_path = Path(td) / "out.xlsx"
            with mock.patch.object(
                self.app,
                "_run_query_to_rows",
                return_value=([], ["id"], 0.1, 10.0),
            ), mock.patch.object(self.app, "_deadline", return_value=object()), mock.patch.object(
                self.app, "_check_deadline"
            ) as mock_check_deadline:
                self.app.run_export_to_template(
                    engine=object(),
                    sql_query="SELECT 1",
                    template_path=str(template_path),
                    output_file_path=str(output_path),
                    sheet_name="Sheet1",
                    start_cell="A1",
                    include_header=False,
                    export_timeout_seconds=60,
                    cancel_event=threading.Event(),
                )

            self.assertGreaterEqual(mock_check_deadline.call_count, 1)

    def test_template_export_checks_deadline_after_save(self):
        fake_wb = _FakeWorkbook()
        with tempfile.TemporaryDirectory() as td:
            template_path = Path(td) / "template.xlsx"
            template_path.write_bytes(b"")
            output_path = Path(td) / "out.xlsx"
            with mock.patch.object(
                self.app,
                "_run_query_to_rows",
                return_value=([("v",)], ["col"], 0.1, 10.0),
            ), mock.patch.object(self.app, "_deadline", return_value=object()), mock.patch.object(
                self.app, "load_workbook", return_value=fake_wb
            ), mock.patch.object(self.app, "_check_deadline") as mock_check_deadline:
                self.app.run_export_to_template(
                    engine=object(),
                    sql_query="SELECT 1",
                    template_path=str(template_path),
                    output_file_path=str(output_path),
                    sheet_name="Sheet1",
                    start_cell="A1",
                    include_header=False,
                    export_timeout_seconds=60,
                    cancel_event=threading.Event(),
                )

            self.assertGreaterEqual(mock_check_deadline.call_count, 3)


if __name__ == "__main__":
    unittest.main()
