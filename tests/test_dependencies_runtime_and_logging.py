import importlib.machinery
import importlib.util
import os
import threading
import time
import unittest
import warnings
from pathlib import Path
from unittest import mock


def _load_main_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_testdeps", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_testdeps", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class _DummyMessageBox:
    def __init__(self):
        self.calls = []

    def showerror(self, title, msg):
        self.calls.append((title, msg))


class DependenciesRuntimeAndLoggingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = _load_main_module()

    def test_missing_deps_message_pip_command_has_no_commas(self):
        with mock.patch.object(self.app, "_MISSING_DEPENDENCIES", ["openpyxl", "sqlalchemy"]):
            msg = self.app._missing_deps_message()
        self.assertIn("python -m pip install", msg)
        # should be space-separated, NOT comma-separated
        self.assertIn("python -m pip install openpyxl sqlalchemy", msg)
        self.assertNotIn("pip install openpyxl, sqlalchemy", msg)

    def test_ensure_runtime_dependencies_no_gui_when_show_gui_false(self):
        # Consolidated coverage for _ensure_runtime_dependencies(show_gui=...)
        # lives in this file (deduplicated from test_dependencies_headless.py).
        dummy = _DummyMessageBox()
        with mock.patch.object(self.app, "_MISSING_DEPENDENCIES", ["openpyxl"]), \
             mock.patch.object(self.app, "_TK_AVAILABLE", True), \
             mock.patch.object(self.app, "messagebox", dummy):
            with self.assertRaises(RuntimeError):
                self.app._ensure_runtime_dependencies(show_gui=False)
        self.assertEqual(dummy.calls, [])

    def test_ensure_runtime_dependencies_shows_gui_when_enabled_and_available(self):
        dummy = _DummyMessageBox()
        with mock.patch.object(self.app, "_MISSING_DEPENDENCIES", ["openpyxl"]), \
             mock.patch.object(self.app, "_TK_AVAILABLE", True), \
             mock.patch.object(self.app, "messagebox", dummy):
            with self.assertRaises(RuntimeError):
                self.app._ensure_runtime_dependencies(show_gui=True)
        self.assertEqual(len(dummy.calls), 1)

    def test_ensure_runtime_dependencies_show_gui_true_tk_unavailable_does_not_call_messagebox(self):
        dummy = _DummyMessageBox()
        with mock.patch.object(self.app, "_MISSING_DEPENDENCIES", ["openpyxl"]), \
             mock.patch.object(self.app, "_TK_AVAILABLE", False), \
             mock.patch.object(self.app, "messagebox", dummy):
            with self.assertRaises(RuntimeError):
                self.app._ensure_runtime_dependencies(show_gui=True)
        self.assertEqual(dummy.calls, [])

    def test_sql_log_excerpt_full_when_env_set(self):
        sql = "SELECT 1;\n" * 1000
        with mock.patch.dict(os.environ, {"KKR_LOG_FULL_SQL": "1"}):
            out = self.app._sql_log_excerpt(sql, max_chars=50, max_lines=1)
        # when debug env is set, we should not truncate
        self.assertEqual(out, sql.rstrip())

    def test_sql_log_excerpt_truncates_by_default(self):
        sql = "SELECT 1;\n" * 1000
        with mock.patch.dict(os.environ, {}, clear=True):
            out = self.app._sql_log_excerpt(sql, max_chars=120, max_lines=3)
        # excerpt should be shorter than full SQL
        self.assertLess(len(out), len(sql))
        # should preserve newlines (not flatten into one line)
        self.assertIn("\n", out)

    def test_format_connection_error_sqlalchemy_version_safe_when_missing(self):
        # This test ensures getattr(sqlalchemy, '__version__', ...) doesn't crash
        # even if sqlalchemy is None (or missing __version__).
        with mock.patch.object(self.app, "sqlalchemy", None):
            title, body = self.app._format_connection_error(
                conn_type="generic",
                exc=Exception("boom"),
                details={},
            )
        self.assertTrue(title)
        self.assertTrue(body)


class CancellerThreadLeakTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = _load_main_module()

    def test_canceller_thread_exits_when_not_cancelled(self):
        if getattr(self.app, "sqlalchemy", None) is None:
            self.skipTest("sqlalchemy not available in test env")
        # Integration-ish unit test:
        # Run a small query with cancel_event provided but never set.
        # Ensure the internal daemon threads (watchdog/canceller) finish promptly.
        app = self.app
        sa = app.sqlalchemy
        engine = app.create_engine(
            "sqlite:///:memory:",
            poolclass=sa.pool.NullPool,
        )
        try:
            created = []
            real_thread_cls = app.threading.Thread

            def recording_thread(*args, **kwargs):
                t = real_thread_cls(*args, **kwargs)
                created.append(t)
                return t

            cancel_event = threading.Event()
            with mock.patch.object(app.threading, "Thread", side_effect=recording_thread):
                rows, cols, sql_dur, total_dur = app._run_query_to_rows(
                    engine=engine,
                    sql_query="SELECT 1 AS a",
                    timeout_seconds=5,
                    cancel_event=cancel_event,
                )
            self.assertEqual(rows, [(1,)])
            self.assertEqual(cols, ["a"])
            self.assertGreaterEqual(sql_dur, 0.0)
            self.assertGreaterEqual(total_dur, 0.0)

            # The created threads should stop soon after done.set()
            # Give them a moment to observe done and exit.
            deadline = time.time() + 3.0
            for t in created:
                # they are daemon threads; join is allowed but should not block long
                remaining = max(0.0, deadline - time.time())
                t.join(timeout=remaining)
            still_alive = [t for t in created if t.is_alive()]
            self.assertEqual(still_alive, [], f"Leaked threads: {[t.name for t in still_alive]}")
        finally:
            engine.dispose()
            with warnings.catch_warnings():
                warnings.simplefilter("error", ResourceWarning)
                import gc

                gc.collect()
