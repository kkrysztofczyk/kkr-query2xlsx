import importlib.machinery
import importlib.util
import socket
import unittest
from email.message import Message
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_update_retry", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_update_retry", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


def _mk_headers(d: dict[str, str]) -> Message:
    m = Message()
    for k, v in d.items():
        m[k] = v
    return m


def _mk_http_error(code: int, headers: dict[str, str] | None = None) -> HTTPError:
    hdrs = _mk_headers(headers or {})
    return HTTPError("https://example.invalid", code, "err", hdrs, None)


class UpdateCheckerRetryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def setUp(self):
        self.app.set_lang("en")

    def test_retries_once_on_timeout(self):
        calls = {"n": 0}

        def fake_get_update_info():
            calls["n"] += 1
            if calls["n"] == 1:
                raise socket.timeout("t")
            return {"ok": True}

        with patch.object(self.app, "get_update_info", side_effect=fake_get_update_info), patch.object(
            self.app.time, "sleep"
        ) as sleep_mock:
            out = self.app._get_update_info_with_retry(retry_once=True, retry_delay_s=0.01)

        self.assertEqual(out, {"ok": True})
        self.assertEqual(calls["n"], 2)
        sleep_mock.assert_called_once()

    def test_does_not_retry_on_rate_limit(self):
        err = _mk_http_error(403, {"x-ratelimit-remaining": "0"})
        with patch.object(self.app, "get_update_info", side_effect=err) as mocked, patch.object(
            self.app.time, "sleep"
        ) as sleep_mock:
            with self.assertRaises(HTTPError):
                self.app._get_update_info_with_retry(retry_once=True, retry_delay_s=0.01)

        self.assertEqual(mocked.call_count, 1)
        sleep_mock.assert_not_called()


if __name__ == "__main__":
    unittest.main()
