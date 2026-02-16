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
    loader = importlib.machinery.SourceFileLoader("app_main_cancel_passthrough", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_cancel_passthrough", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class CancelPassthroughTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def test_run_export_passes_cancel_event_to_db_phase(self):
        cancel_evt = threading.Event()

        def fake_run_query_to_rows(
            engine,
            sql_query,
            timeout_seconds=0,
            cancel_event=None,
            sql_source_path=None,
        ):
            self.assertIs(cancel_event, cancel_evt)
            return [], ["id"], 0.1, 10.0

        with mock.patch.object(self.app, "_run_query_to_rows", side_effect=fake_run_query_to_rows):
            with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
                sql_dur, export_dur, total_dur, rows_count = self.app.run_export(
                    engine=object(),
                    sql_query="SELECT 1",
                    output_file_path=tmp.name,
                    output_format="xlsx",
                    cancel_event=cancel_evt,
                )

        self.assertEqual(rows_count, 0)
        self.assertEqual(export_dur, 0.0)
        self.assertEqual(total_dur, sql_dur)

    def test_run_export_passes_cancel_event_to_csv_export_phase(self):
        cancel_event = threading.Event()

        with mock.patch.object(
            self.app,
            "_run_query_to_rows",
            return_value=([(1,)], ["id"], 0.2, 20.0),
        ), mock.patch.object(self.app, "_export_rows_to_csv") as mock_export_csv:
            with tempfile.NamedTemporaryFile(suffix=".csv") as tmp:
                self.app.run_export(
                    engine=object(),
                    sql_query="SELECT 1",
                    output_file_path=tmp.name,
                    output_format="csv",
                    cancel_event=cancel_event,
                )

        self.assertTrue(mock_export_csv.called)
        self.assertIs(mock_export_csv.call_args.kwargs["cancel_event"], cancel_event)

    def test_run_export_passes_cancel_event_to_xlsx_export_phase(self):
        cancel_event = threading.Event()

        with mock.patch.object(
            self.app,
            "_run_query_to_rows",
            return_value=([(1,)], ["id"], 0.2, 20.0),
        ), mock.patch.object(self.app, "_export_rows_to_xlsx") as mock_export_xlsx:
            with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
                self.app.run_export(
                    engine=object(),
                    sql_query="SELECT 1",
                    output_file_path=tmp.name,
                    output_format="xlsx",
                    cancel_event=cancel_event,
                )

        self.assertTrue(mock_export_xlsx.called)
        self.assertIs(mock_export_xlsx.call_args.kwargs["cancel_event"], cancel_event)

    def test_run_export_removes_partial_file_when_cancelled_before_export(self):
        cancel_event = threading.Event()
        cancel_event.set()

        with mock.patch.object(
            self.app,
            "_run_query_to_rows",
            return_value=([(1,)], ["id"], 0.2, 20.0),
        ):
            with tempfile.TemporaryDirectory() as td:
                output_path = Path(td) / "cancelled.csv"
                with self.assertRaises(self.app.UserCancelledError):
                    self.app.run_export(
                        engine=object(),
                        sql_query="SELECT 1",
                        output_file_path=str(output_path),
                        output_format="csv",
                        cancel_event=cancel_event,
                    )

                self.assertFalse(output_path.exists())


if __name__ == "__main__":
    unittest.main()
