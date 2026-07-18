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
    loader = importlib.machinery.SourceFileLoader("app_main_export_output_safety", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_export_output_safety", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class ExportOutputSafetyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def test_run_export_xlsx_error_preserves_existing_output(self):
        with tempfile.TemporaryDirectory() as td:
            output_path = Path(td) / "report.xlsx"
            output_path.write_bytes(b"old-content")

            with mock.patch.object(
                self.app,
                "_run_query_to_rows",
                return_value=([(1,)], ["id"], 0.2, 20.0),
            ), mock.patch.object(
                self.app,
                "_export_rows_to_xlsx",
                side_effect=RuntimeError("boom"),
            ):
                with self.assertRaisesRegex(RuntimeError, "boom"):
                    self.app.run_export(
                        engine=object(),
                        sql_query="SELECT 1",
                        output_file_path=str(output_path),
                        output_format="xlsx",
                    )

            self.assertTrue(output_path.exists())
            self.assertEqual(output_path.read_bytes(), b"old-content")

    def test_run_export_csv_error_preserves_existing_output(self):
        with tempfile.TemporaryDirectory() as td:
            output_path = Path(td) / "report.csv"
            output_path.write_bytes(b"old-content")

            with mock.patch.object(
                self.app,
                "_run_query_to_rows",
                return_value=([(1,)], ["id"], 0.2, 20.0),
            ), mock.patch.object(
                self.app,
                "_export_rows_to_csv",
                side_effect=RuntimeError("boom"),
            ):
                with self.assertRaisesRegex(RuntimeError, "boom"):
                    self.app.run_export(
                        engine=object(),
                        sql_query="SELECT 1",
                        output_file_path=str(output_path),
                        output_format="csv",
                    )

            self.assertTrue(output_path.exists())
            self.assertEqual(output_path.read_bytes(), b"old-content")

    def test_run_export_csv_success_replaces_existing_output(self):
        with tempfile.TemporaryDirectory() as td:
            output_path = Path(td) / "report.csv"
            output_path.write_text("old-content\n", encoding="utf-8")

            with mock.patch.object(
                self.app,
                "_run_query_to_rows",
                return_value=([(1, "Alice")], ["id", "name"], 0.2, 20.0),
            ):
                self.app.run_export(
                    engine=object(),
                    sql_query="SELECT 1, 'Alice'",
                    output_file_path=str(output_path),
                    output_format="csv",
                )

            self.assertEqual(output_path.read_text(encoding="utf-8"), "id,name\n1,Alice\n")

    def test_run_export_error_cleans_temp_file(self):
        with tempfile.TemporaryDirectory() as td:
            for output_name, export_func in (
                ("report.csv", "_export_rows_to_csv"),
                ("report.xlsx", "_export_rows_to_xlsx"),
            ):
                with self.subTest(output_name=output_name):
                    output_path = Path(td) / output_name
                    output_path.write_bytes(b"old-content")

                    def _write_partial_then_fail(path, **_kwargs):
                        Path(path).write_bytes(b"partial")
                        raise RuntimeError("boom")

                    with mock.patch.object(
                        self.app,
                        "_run_query_to_rows",
                        return_value=([(1,)], ["id"], 0.2, 20.0),
                    ), mock.patch.object(
                        self.app,
                        export_func,
                        side_effect=_write_partial_then_fail,
                    ):
                        with self.assertRaisesRegex(RuntimeError, "boom"):
                            self.app.run_export(
                                engine=object(),
                                sql_query="SELECT 1",
                                output_file_path=str(output_path),
                                output_format=output_path.suffix.lstrip("."),
                            )

                    temp_candidates = list(Path(td).glob(f".{output_name}.*.tmp"))
                    self.assertEqual(temp_candidates, [])

    def test_run_export_cancelled_before_final_replace_preserves_existing_csv(self):
        cancel_evt = threading.Event()
        with tempfile.TemporaryDirectory() as td:
            output_path = Path(td) / "report.csv"
            output_path.write_text("old-content\n", encoding="utf-8")

            def _export_then_cancel(path, **_kwargs):
                Path(path).write_text("id\n1\n", encoding="utf-8")
                cancel_evt.set()

            with mock.patch.object(
                self.app,
                "_run_query_to_rows",
                return_value=([(1,)], ["id"], 0.2, 20.0),
            ), mock.patch.object(
                self.app,
                "_export_rows_to_csv",
                side_effect=_export_then_cancel,
            ):
                with self.assertRaises(self.app.UserCancelledError):
                    self.app.run_export(
                        engine=object(),
                        sql_query="SELECT 1",
                        output_file_path=str(output_path),
                        output_format="csv",
                        cancel_event=cancel_evt,
                    )

            self.assertEqual(output_path.read_text(encoding="utf-8"), "old-content\n")
            temp_candidates = list(Path(td).glob(".report.csv.*.tmp"))
            self.assertEqual(temp_candidates, [])

    def test_run_export_cancelled_before_final_replace_preserves_existing_xlsx(self):
        cancel_evt = threading.Event()
        with tempfile.TemporaryDirectory() as td:
            output_path = Path(td) / "report.xlsx"
            output_path.write_bytes(b"old-content")

            def _export_then_cancel(path, **_kwargs):
                Path(path).write_bytes(b"new-content")
                cancel_evt.set()

            with mock.patch.object(
                self.app,
                "_run_query_to_rows",
                return_value=([(1,)], ["id"], 0.2, 20.0),
            ), mock.patch.object(
                self.app,
                "_export_rows_to_xlsx",
                side_effect=_export_then_cancel,
            ):
                with self.assertRaises(self.app.UserCancelledError):
                    self.app.run_export(
                        engine=object(),
                        sql_query="SELECT 1",
                        output_file_path=str(output_path),
                        output_format="xlsx",
                        cancel_event=cancel_evt,
                    )

            self.assertEqual(output_path.read_bytes(), b"old-content")
            temp_candidates = list(Path(td).glob(".report.xlsx.*.tmp"))
            self.assertEqual(temp_candidates, [])

    def test_format_error_for_ui_permission_error_prefers_filename2(self):
        exc = PermissionError("blocked")
        exc.filename = ".report.xlsx.abc.tmp"
        exc.filename2 = "report.xlsx"

        with mock.patch.object(self.app.os.path, "exists", return_value=True), mock.patch.object(
            self.app,
            "shorten_path",
            wraps=self.app.shorten_path,
        ) as mock_shorten:
            message = self.app.format_error_for_ui(
                exc=exc,
                sql_query="SELECT 1",
                context="export",
            )

        self.assertIn("report.xlsx", message)
        self.assertTrue(
            any(call.args and call.args[0] == "report.xlsx" for call in mock_shorten.call_args_list)
        )


if __name__ == "__main__":
    unittest.main()
