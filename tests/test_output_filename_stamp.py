import importlib.machinery
import importlib.util
import tempfile
import unittest
from datetime import datetime
from pathlib import Path


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_output_stamp", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_output_stamp", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class OutputFilenameStampTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def test_render_output_filename_stamp_tokens(self):
        fixed_dt = datetime(2026, 1, 30, 19, 5, 7)

        rendered = self.app._render_output_filename_stamp("{YYYY-MM-DD hh-mm-ss}", dt=fixed_dt)

        self.assertEqual(rendered, "{2026-01-30 19-05-07}")

    def test_sanitize_filename_stamp_removes_windows_invalid_characters(self):
        stamp = self.app._sanitize_filename_stamp("[2026/01/30 19:05*bad?\"<x>|]")

        self.assertNotIn(":", stamp)
        self.assertNotIn("/", stamp)
        self.assertNotIn("*", stamp)
        self.assertNotIn("?", stamp)
        self.assertNotIn('"', stamp)
        self.assertNotIn("<", stamp)
        self.assertNotIn(">", stamp)
        self.assertNotIn("|", stamp)

    def test_sanitize_filename_stamp_removes_control_chars_and_invalid_windows_chars(self):
        stamp = self.app._sanitize_filename_stamp("[2026-01-30\n19:05]\t\x7f")

        self.assertNotIn("\n", stamp)
        self.assertNotIn("\t", stamp)
        self.assertNotIn("\x7f", stamp)
        self.assertNotIn(":", stamp)

    def test_apply_output_filename_stamp_prefix_and_suffix(self):
        fixed_dt = datetime(2026, 1, 30, 19, 5)

        prefixed = self.app.apply_output_filename_stamp(
            "report.xlsx",
            enabled=True,
            pattern="[YYYY-MM-DD]",
            place="prefix",
            dt=fixed_dt,
        )
        suffixed = self.app.apply_output_filename_stamp(
            "report.xlsx",
            enabled=True,
            pattern="[YYYY-MM-DD]",
            place="suffix",
            dt=fixed_dt,
        )

        self.assertEqual(prefixed, "[2026-01-30]report.xlsx")
        self.assertEqual(suffixed, "report[2026-01-30].xlsx")


    def test_apply_output_filename_stamp_respects_pattern_edge_whitespace_as_separator(self):
        fixed_dt = datetime(2026, 2, 17, 8, 9)

        prefixed = self.app.apply_output_filename_stamp(
            "report__123rows.xlsx",
            enabled=True,
            pattern="[YYYY-MM-DD] ",
            place="prefix",
            dt=fixed_dt,
        )
        suffixed = self.app.apply_output_filename_stamp(
            "report__123rows.xlsx",
            enabled=True,
            pattern=" [YYYY-MM-DD]",
            place="suffix",
            dt=fixed_dt,
        )

        self.assertEqual(prefixed, "[2026-02-17] report__123rows.xlsx")
        self.assertEqual(suffixed, "report__123rows [2026-02-17].xlsx")

    def test_apply_output_filename_stamp_allows_multiple_spaces_from_pattern_edges(self):
        fixed_dt = datetime(2026, 2, 18, 8, 9)

        prefixed = self.app.apply_output_filename_stamp(
            "report__123rows.xlsx",
            enabled=True,
            pattern="[YYYY-MM-DD]        ",
            place="prefix",
            dt=fixed_dt,
        )
        suffixed = self.app.apply_output_filename_stamp(
            "report__123rows.xlsx",
            enabled=True,
            pattern="        [YYYY-MM-DD]",
            place="suffix",
            dt=fixed_dt,
        )

        self.assertEqual(prefixed, "[2026-02-18]        report__123rows.xlsx")
        self.assertEqual(suffixed, "report__123rows        [2026-02-18].xlsx")

    def test_apply_output_filename_stamp_treats_empty_pattern_as_disabled(self):
        unchanged = self.app.apply_output_filename_stamp(
            "report.xlsx",
            enabled=True,
            pattern="   ",
            place="suffix",
        )

        self.assertEqual(unchanged, "report.xlsx")

    def test_ui_config_persists_output_filename_stamp_settings(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            self.app.save_ui_config(
                tmp_path,
                {
                    "ui": {
                        "output_filename_stamp_enabled": True,
                        "output_filename_stamp_pattern": "{YYYY-MM-DD hh-mm}",
                        "output_filename_stamp_place": "prefix",
                    }
                },
            )
            cfg = self.app.load_ui_config(tmp_path)
            ui = cfg.get("ui") or {}

            self.assertTrue(ui.get("output_filename_stamp_enabled"))
            self.assertEqual(ui.get("output_filename_stamp_pattern"), "{YYYY-MM-DD hh-mm}")
            self.assertEqual(ui.get("output_filename_stamp_place"), "prefix")


if __name__ == "__main__":
    unittest.main()
