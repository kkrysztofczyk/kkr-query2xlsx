import importlib.machinery
import importlib.util
import io
import logging
import os
import unittest
from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.pool import NullPool


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_sql_logging", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_sql_logging", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class SqlLoggingPolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def setUp(self):
        self._old_lang = getattr(self.app, "_CURRENT_LANG", "en")
        self.app.set_lang("en")
        self._old_env = os.environ.get("KKR_LOG_FULL_SQL")

    def tearDown(self):
        self.app.set_lang(self._old_lang)
        if self._old_env is None:
            os.environ.pop("KKR_LOG_FULL_SQL", None)
        else:
            os.environ["KKR_LOG_FULL_SQL"] = self._old_env

    def _capture_sql_logs(self, fn):
        stream = io.StringIO()
        handler = logging.StreamHandler(stream)
        logger = self.app.LOGGER
        old_level = logger.level
        logger.setLevel(logging.INFO)
        logger.addHandler(handler)
        try:
            fn()
        finally:
            logger.removeHandler(handler)
            logger.setLevel(old_level)
        return stream.getvalue()

    def test_default_logs_excerpt_only(self):
        os.environ.pop("KKR_LOG_FULL_SQL", None)
        secret = "SECRET_ABC_123"
        sql = "SELECT '" + ("A" * 7000) + secret + "' AS x;"

        engine = create_engine("sqlite:///:memory:", poolclass=NullPool)
        try:
            logs = self._capture_sql_logs(
                lambda: self.app._run_query_to_rows(engine, sql, sql_source_path="/tmp/query.sql")
            )
        finally:
            engine.dispose()

        self.assertIn("Executing SQL (excerpt)", logs)
        self.assertIn("SQL source: query.sql", logs)
        self.assertNotIn(secret, logs)

    def test_env_enabled_logs_full_sql(self):
        os.environ["KKR_LOG_FULL_SQL"] = "1"
        secret = "SECRET_ABC_123"
        sql = "SELECT '" + ("A" * 7000) + secret + "' AS x;"

        engine = create_engine("sqlite:///:memory:", poolclass=NullPool)
        try:
            logs = self._capture_sql_logs(
                lambda: self.app._run_query_to_rows(engine, sql, sql_source_path="/tmp/query.sql")
            )
        finally:
            engine.dispose()

        self.assertIn("Executing SQL (full)", logs)
        self.assertIn(secret, logs)


    def test_exception_logging_uses_max_two_context_entries_plus_traceback(self):
        os.environ.pop("KKR_LOG_FULL_SQL", None)

        def emit():
            try:
                raise ValueError("boom")
            except ValueError as exc:
                self.app._log_sql_exception(
                    "DBAPIError while executing SQL.",
                    "SELECT 1",
                    sql_source_path="/tmp/context.sql",
                    error=exc,
                )

        logs = self._capture_sql_logs(emit)

        self.assertIn("DBAPIError while executing SQL. | SQL source: context.sql", logs)
        self.assertEqual(logs.count("DBAPIError while executing SQL."), 1)
        self.assertEqual(logs.count("SQL source: context.sql"), 1)
        self.assertIn("SQL execution failed (excerpt): boom", logs)
        self.assertIn("Traceback (most recent call last)", logs)

    def test_exception_path_uses_excerpt_when_full_disabled(self):
        os.environ.pop("KKR_LOG_FULL_SQL", None)
        secret = "SECRET_ABC_123"
        sql = "SELECT * FROM missing_table -- " + ("A" * 7000) + secret

        engine = create_engine("sqlite:///:memory:", poolclass=NullPool)
        try:
            def run_bad_query():
                with self.assertRaises(Exception):
                    self.app._run_query_to_rows(engine, sql, sql_source_path="/tmp/bad.sql")

            logs = self._capture_sql_logs(run_bad_query)
        finally:
            engine.dispose()

        self.assertIn("SQL execution failed (excerpt)", logs)
        self.assertIn("SQL source: bad.sql", logs)
        self.assertIn("Traceback (most recent call last)", logs)
        self.assertNotIn(secret, logs)


if __name__ == "__main__":
    unittest.main()
