import importlib.machinery
import importlib.util
import logging
import os
import shutil
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch



def _rm_tree_strict(path: Path) -> None:
    last_error = None
    for _ in range(3):
        try:
            shutil.rmtree(path)
            return
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            time.sleep(0.2)
    if last_error is not None:
        raise last_error


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_logger", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_logger", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class LoggerHandlerSwitchTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def setUp(self):
        self.logger = logging.getLogger("kkr-query2xlsx")
        self._original_handlers = list(self.logger.handlers)
        for handler in self._original_handlers:
            self.logger.removeHandler(handler)


    def _mkdtemp(self) -> Path:
        temp_dir = Path(tempfile.mkdtemp(prefix="kkr-q2x-logs-"))
        strict = os.getenv("CI") in ("1", "true", "True")
        if strict:
            self.addCleanup(lambda: _rm_tree_strict(temp_dir))
        else:
            self.addCleanup(lambda: shutil.rmtree(temp_dir, ignore_errors=True))
        return temp_dir

    def tearDown(self):
        for handler in list(self.logger.handlers):
            self.logger.removeHandler(handler)
            if handler not in self._original_handlers:
                try:
                    handler.close()
                except Exception:
                    pass
        for handler in self._original_handlers:
            self.logger.addHandler(handler)

    def test_attach_file_handler_removes_fallback_stderr_handler(self):
        fallback_handler = logging.StreamHandler(self.app.sys.stderr)
        setattr(fallback_handler, "_kkr_fallback", True)
        self.logger.addHandler(fallback_handler)

        log_dir = self._mkdtemp()
        attached = self.app._attach_logger_file_handler(str(log_dir))

        self.assertTrue(attached)
        self.assertEqual(sum(isinstance(h, logging.StreamHandler) and not isinstance(h, self.app.RotatingFileHandler) for h in self.logger.handlers), 0)
        self.assertEqual(sum(isinstance(h, self.app.RotatingFileHandler) for h in self.logger.handlers), 1)

    def test_attach_file_handler_keeps_non_fallback_stream_handler(self):
        log_dir = self._mkdtemp()
        custom_stream = (log_dir / "tmp_test_stream.log").open("w", encoding="utf-8")
        self.addCleanup(custom_stream.close)
        custom_handler = logging.StreamHandler(custom_stream)
        self.logger.addHandler(custom_handler)

        attached = self.app._attach_logger_file_handler(str(log_dir))

        self.assertTrue(attached)
        self.assertIn(custom_handler, self.logger.handlers)
        self.assertEqual(sum(isinstance(h, self.app.RotatingFileHandler) for h in self.logger.handlers), 1)

    def test_attach_file_handler_is_idempotent_without_duplicates(self):
        log_dir = self._mkdtemp()

        first_attach = self.app._attach_logger_file_handler(str(log_dir))
        second_attach = self.app._attach_logger_file_handler(str(log_dir))

        self.assertTrue(first_attach)
        self.assertTrue(second_attach)
        file_handlers = [h for h in self.logger.handlers if isinstance(h, self.app.RotatingFileHandler)]
        self.assertEqual(len(file_handlers), 1)

    def test_set_data_dir_repoints_log_handler_to_data_dir_logs(self):
        data_dir = self._mkdtemp()

        self.app._set_data_dir(str(data_dir))

        self.assertEqual(Path(self.app.LOG_DIR), data_dir / "logs")
        self.assertEqual(Path(self.app.LOG_FILE_PATH), data_dir / "logs" / "kkr-query2xlsx.log")
        file_handlers = [h for h in self.logger.handlers if isinstance(h, self.app.RotatingFileHandler)]
        self.assertEqual(len(file_handlers), 1)
        self.assertEqual(Path(file_handlers[0].baseFilename), data_dir / "logs" / "kkr-query2xlsx.log")

    def test_attach_file_handler_failure_keeps_existing_fallback_handler(self):
        fallback_handler = logging.StreamHandler(self.app.sys.stderr)
        setattr(fallback_handler, "_kkr_fallback", True)
        self.logger.addHandler(fallback_handler)

        with patch.object(self.app.os, "makedirs", side_effect=PermissionError("blocked")):
            attached = self.app._attach_logger_file_handler("/unwritable")

        self.assertFalse(attached)
        self.assertIn(fallback_handler, self.logger.handlers)
        self.assertEqual(sum(isinstance(h, self.app.RotatingFileHandler) for h in self.logger.handlers), 0)


if __name__ == "__main__":
    unittest.main()
