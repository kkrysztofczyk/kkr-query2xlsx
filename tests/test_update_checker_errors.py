import importlib.machinery
import importlib.util
import socket
import unittest
from email.message import Message
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError, URLError


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_update_checker", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_update_checker", loader)
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


class UpdateCheckerErrorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def setUp(self):
        # deterministycznie: komunikaty EN
        self.app.set_lang("en")

    # --- _parse_retry_hint hardening ---

    def test_parse_retry_hint_huge_retry_after_does_not_raise(self):
        headers = {"retry-after": "9" * 200}  # absurdalnie duża liczba
        try:
            out = self.app._parse_retry_hint(headers)
        except Exception as exc:  # noqa: BLE001
            self.fail(f"_parse_retry_hint must not raise, but raised: {exc!r}")
        self.assertTrue(out is None or isinstance(out, str))

    def test_parse_retry_hint_falls_back_to_reset_when_retry_after_is_invalid(self):
        # Retry-After is absurdly large, but X-RateLimit-Reset is usable -> should not return None.
        with patch.object(self.app.time, "time", return_value=1000):
            headers = {"retry-after": "9" * 200, "x-ratelimit-reset": "1060"}
            out = self.app._parse_retry_hint(headers)
        self.assertIsInstance(out, str)
        self.assertTrue(out)


    def test_parse_retry_hint_prefers_reset_over_non_numeric_retry_after(self):
        with patch.object(self.app.time, "time", return_value=1000):
            headers = {
                "retry-after": "Wed, 21 Oct 2015 07:28:00 GMT",
                "x-ratelimit-reset": "1060",
            }
            out = self.app._parse_retry_hint(headers)
        self.assertIsInstance(out, str)
        self.assertRegex(out, r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")
        self.assertNotEqual(out, headers["retry-after"])

    def test_parse_retry_hint_returns_retry_after_fallback_when_no_reset(self):
        headers = {
            "retry-after": "Wed, 21 Oct 2015 07:28:00 GMT",
        }
        out = self.app._parse_retry_hint(headers)
        self.assertEqual(out, headers["retry-after"][:64])

    def test_parse_retry_hint_huge_ratelimit_reset_does_not_raise(self):
        headers = {"x-ratelimit-reset": str(10**200)}
        try:
            out = self.app._parse_retry_hint(headers)
        except Exception as exc:  # noqa: BLE001
            self.fail(f"_parse_retry_hint must not raise, but raised: {exc!r}")
        self.assertTrue(out is None or isinstance(out, str))

    # --- klasyfikacja błędów ---

    def test_classify_timeout_socket(self):
        key, params = self.app._classify_update_check_error(socket.timeout("t"))
        self.assertEqual(key, "UPD_ERR_TIMEOUT")
        self.assertEqual(params, {})

    def test_classify_network_urlerror(self):
        key, params = self.app._classify_update_check_error(URLError(OSError("dns")))
        self.assertEqual(key, "UPD_ERR_NETWORK")
        self.assertEqual(params, {})

    def test_classify_rate_limited_429_always(self):
        err = _mk_http_error(429, {})  # brak nagłówków
        key, params = self.app._classify_update_check_error(err)
        self.assertEqual(key, "UPD_ERR_RATE_LIMITED")
        self.assertEqual(params.get("status"), 429)
        self.assertIn("retry_at", params)

    def test_classify_rate_limited_from_retry_after(self):
        err = _mk_http_error(429, {"retry-after": "10"})
        key, params = self.app._classify_update_check_error(err)
        self.assertEqual(key, "UPD_ERR_RATE_LIMITED")
        self.assertEqual(params.get("status"), 429)
        self.assertIn("retry_at", params)
        self.assertIsInstance(params.get("retry_at"), str)

    def test_classify_rate_limited_403_remaining_zero(self):
        err = _mk_http_error(403, {"x-ratelimit-remaining": "0"})
        key, params = self.app._classify_update_check_error(err)
        self.assertEqual(key, "UPD_ERR_RATE_LIMITED")
        self.assertEqual(params.get("status"), 403)
        self.assertEqual(params.get("retry_at"), "unknown")

    def test_classify_http_403_without_rate_limit_signals(self):
        err = _mk_http_error(403, {})  # np. proxy/firewall
        key, params = self.app._classify_update_check_error(err)
        self.assertEqual(key, "UPD_ERR_HTTP")
        self.assertEqual(params.get("status"), 403)

    # --- “airbag”: handler błędu nie może rzucić wyjątku ---

    def test_build_message_never_raises_even_if_retry_hint_parser_breaks(self):
        err = _mk_http_error(429, {"retry-after": "10"})
        with patch.object(self.app, "_parse_retry_hint", side_effect=OverflowError("boom")):
            msg = self.app._build_update_check_message_with_hint(err)
        self.assertIsInstance(msg, str)
        self.assertTrue(msg)


if __name__ == "__main__":
    unittest.main()
