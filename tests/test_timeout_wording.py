import importlib.machinery
import importlib.util
import unittest
from pathlib import Path


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_timeout_wording", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_timeout_wording", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class TimeoutWordingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()
        cls.readme_text = (Path(__file__).resolve().parents[1] / "README.md").read_text(encoding="utf-8")

    def test_export_timeout_label_mentions_sqlite_in_both_languages(self):
        self.assertEqual(
            self.app.I18N["en"]["LBL_EXPORT_TIMEOUT_MIN"],
            "Export timeout (minutes) - XLSX/CSV/SQLite generation",
        )
        self.assertEqual(
            self.app.I18N["pl"]["LBL_EXPORT_TIMEOUT_MIN"],
            "Limit czasu eksportu (minuty) - generowanie XLSX/CSV/SQLite",
        )

    def test_app_title_mentions_sqlite_in_both_languages(self):
        self.assertEqual(
            self.app.I18N["en"]["APP_TITLE_FULL"],
            "KKr SQL to XLSX/CSV/SQLite",
        )
        self.assertEqual(
            self.app.I18N["pl"]["APP_TITLE_FULL"],
            "KKr SQL to XLSX/CSV/SQLite",
        )

    def test_timeout_minutes_from_seconds_ceils_with_minimum_of_one(self):
        cases = [
            (0, 1),
            (1, 1),
            (59, 1),
            (60, 1),
            (61, 2),
            (119, 2),
            (120, 2),
            (121, 3),
        ]
        for seconds, expected_minutes in cases:
            with self.subTest(seconds=seconds):
                self.assertEqual(
                    self.app._timeout_minutes_from_seconds(seconds),
                    expected_minutes,
                )

    def test_readme_timeout_sections_mention_sqlite(self):
        self.assertIn(
            "`export_seconds` covers **XLSX/CSV/SQLite generation** time.",
            self.readme_text,
        )
        self.assertIn(
            "2) **Export timeout** — XLSX/CSV/SQLite generation.",
            self.readme_text,
        )
        self.assertIn(
            "### Export timeout (XLSX/CSV/SQLite)",
            self.readme_text,
        )


if __name__ == "__main__":
    unittest.main()
