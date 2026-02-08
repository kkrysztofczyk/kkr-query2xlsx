import importlib.machinery
import importlib.util
import tempfile
import unittest
from pathlib import Path


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_sql_validation", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_sql_validation", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class SqlFileValidationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def _write_temp(self, suffix: str, payload: bytes) -> str:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        with tmp:
            tmp.write(payload)
        return tmp.name

    def test_utf8_bom_sql_is_allowed(self):
        path = self._write_temp(".sql", b"\xef\xbb\xbfSELECT 1;\n")
        ok, _ = self.app.validate_sql_text_file(path)
        self.assertTrue(ok)

    def test_utf16_bom_sql_is_allowed(self):
        path = self._write_temp(".sql", "SELECT 1;\n".encode("utf-16"))
        ok, _ = self.app.validate_sql_text_file(path)
        self.assertTrue(ok)

    def test_utf16_without_bom_is_allowed(self):
        path = self._write_temp(".sql", "SELECT 1;\n".encode("utf-16-le"))
        ok, _ = self.app.validate_sql_text_file(path)
        self.assertTrue(ok)

    def test_zip_magic_is_blocked(self):
        path = self._write_temp(".sql", b"PK\x03\x04\x14\x00\x00\x00")
        ok, _ = self.app.validate_sql_text_file(path)
        self.assertFalse(ok)

    def test_spreadsheet_extension_is_blocked(self):
        path = self._write_temp(".xlsx", b"SELECT 1;\n")
        ok, _ = self.app.validate_sql_text_file(path)
        self.assertFalse(ok)

    def test_csv_extension_is_not_blocked(self):
        path = self._write_temp(".csv", b"id,name\n1,Alice\n")
        ok, _ = self.app.validate_sql_text_file(path)
        self.assertTrue(ok)

    def test_sqlite_magic_is_blocked(self):
        path = self._write_temp(".sql", b"SQLite format 3\x00rest")
        ok, _ = self.app.validate_sql_text_file(path)
        self.assertFalse(ok)

    def test_obvious_binary_is_blocked(self):
        path = self._write_temp(".sql", b"\x00\x01\x02\x00" * 128)
        ok, _ = self.app.validate_sql_text_file(path)
        self.assertFalse(ok)


if __name__ == "__main__":
    unittest.main()
