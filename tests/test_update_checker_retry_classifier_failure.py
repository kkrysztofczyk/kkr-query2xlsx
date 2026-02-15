import importlib.machinery
import importlib.util
import socket
import unittest
from pathlib import Path
from unittest.mock import patch


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_update_retry_classifier", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_update_retry_classifier", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class UpdateCheckerRetryClassifierFailureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def test_classifier_failure_reraises_original_exception(self):
        original_exc = socket.timeout("boom")

        with patch.object(self.app, "get_update_info", side_effect=original_exc) as mocked_get, patch.object(
            self.app, "_classify_update_check_error", side_effect=RuntimeError("classifier failed")
        ) as mocked_classifier, patch.object(self.app.time, "sleep") as mocked_sleep:
            with self.assertRaises(socket.timeout) as raised:
                self.app._get_update_info_with_retry(retry_once=True, retry_delay_s=0.01)

        self.assertIs(raised.exception, original_exc)
        self.assertEqual(mocked_get.call_count, 1)
        self.assertEqual(mocked_classifier.call_count, 1)
        mocked_sleep.assert_not_called()


if __name__ == "__main__":
    unittest.main()
