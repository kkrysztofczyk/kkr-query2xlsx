import importlib.machinery
import importlib.util
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest import mock


class _ConnectionProxy:
    def __init__(self, conn, on_close):
        self._conn = conn
        self._on_close = on_close

    def close(self):
        self._on_close(self._conn)
        return self._conn.close()

    def __getattr__(self, name):
        return getattr(self._conn, name)

    def __enter__(self):
        self._conn.__enter__()
        return self

    def __exit__(self, exc_type, exc, tb):
        return self._conn.__exit__(exc_type, exc, tb)


class SQLiteLeakProbeTests(unittest.TestCase):
    @staticmethod
    def _load_main_module():
        repo_root = Path(__file__).resolve().parents[1]
        main_path = repo_root / "main.pyw"
        loader = importlib.machinery.SourceFileLoader("app_main_sqlite_leak_probe", str(main_path))
        spec = importlib.util.spec_from_loader("app_main_sqlite_leak_probe", loader)
        module = importlib.util.module_from_spec(spec)
        loader.exec_module(module)
        return module

    def test_sqlite_leak_probe_closes_all_opened_connections_after_dispose(self):
        app = self._load_main_module()
        opened = []
        closed = set()
        real_connect = sqlite3.connect

        def tracking_connect(*args, **kwargs):
            conn = real_connect(*args, **kwargs)
            opened.append(conn)
            return _ConnectionProxy(conn, lambda c: closed.add(id(c)))

        with tempfile.TemporaryDirectory() as tmp_dir:
            out_path = str(Path(tmp_dir) / "out.csv")
            with mock.patch("sqlite3.connect", side_effect=tracking_connect):
                engine = app.create_engine("sqlite:///:memory:")
                try:
                    app.run_export(
                        engine,
                        "SELECT 1 AS x",
                        out_path,
                        "csv",
                        csv_profile=app.DEFAULT_CSV_PROFILE,
                        db_timeout_seconds=10,
                        export_timeout_seconds=10,
                        sql_source_path=None,
                    )
                finally:
                    try:
                        engine.dispose()
                    except Exception:
                        pass

        leaked_ids = sorted(id(conn) for conn in opened if id(conn) not in closed)
        self.assertFalse(
            leaked_ids,
            f"Detected unclosed sqlite DBAPI connections after engine.dispose(): {leaked_ids}",
        )


if __name__ == "__main__":
    unittest.main()
