import importlib.machinery
import importlib.util
import unittest
from pathlib import Path


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_connection_status", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_connection_status", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class ConnectionStatusTextFormattingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def test_cli_using_connection_uses_separator_without_nested_parentheses_en(self):
        self.app.set_lang("en")
        rendered = self.app.t(
            "CLI_USING_CONNECTION",
            name="Default MSSQL",
            type="SQL Server (ODBC)",
        )
        self.assertEqual(rendered, "Using connection: Default MSSQL - SQL Server (ODBC).")
        self.assertNotIn("((", rendered)
        self.assertNotIn("))", rendered)

    def test_status_connected_uses_separator_without_nested_parentheses_en(self):
        self.app.set_lang("en")
        rendered = self.app.t(
            "STATUS_CONNECTED",
            name="Default MSSQL",
            type="SQL Server (ODBC)",
        )
        self.assertEqual(rendered, "Connected to Default MSSQL - SQL Server (ODBC).")
        self.assertNotIn("((", rendered)
        self.assertNotIn("))", rendered)

    def test_cli_using_connection_uses_separator_without_nested_parentheses_pl(self):
        self.app.set_lang("pl")
        rendered = self.app.t(
            "CLI_USING_CONNECTION",
            name="Domyślny MSSQL",
            type="SQL Server (ODBC)",
        )
        self.assertEqual(rendered, "Używam połączenia: Domyślny MSSQL - SQL Server (ODBC).")
        self.assertNotIn("((", rendered)
        self.assertNotIn("))", rendered)

    def test_status_connected_uses_separator_without_nested_parentheses_pl(self):
        self.app.set_lang("pl")
        rendered = self.app.t(
            "STATUS_CONNECTED",
            name="Domyślny MSSQL",
            type="SQL Server (ODBC)",
        )
        self.assertEqual(rendered, "Połączono z Domyślny MSSQL - SQL Server (ODBC).")
        self.assertNotIn("((", rendered)
        self.assertNotIn("))", rendered)


if __name__ == "__main__":
    unittest.main()
