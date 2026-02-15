import json
import socket
import time
import unittest
from email.message import Message
from unittest.mock import patch
from urllib.error import HTTPError, URLError

import updater


def _mk_headers(d: dict[str, str]) -> Message:
    m = Message()
    for k, v in d.items():
        m[k] = v
    return m


def _mk_http_error(code: int, headers: dict[str, str] | None = None) -> HTTPError:
    hdrs = _mk_headers(headers or {})
    return HTTPError("https://example.invalid", code, "err", hdrs, None)


class UpdaterErrorTests(unittest.TestCase):
    # --- _parse_retry_hint hardening ---

    def test_parse_retry_hint_without_headers_returns_none(self):
        self.assertIsNone(updater._parse_retry_hint({}))
        self.assertIsNone(updater._parse_retry_hint(None))

    def test_parse_retry_hint_huge_retry_after_does_not_raise(self):
        headers = {"retry-after": "9" * 200}
        try:
            out = updater._parse_retry_hint(headers)
        except Exception as exc:  # noqa: BLE001
            self.fail(f"_parse_retry_hint must not raise, but raised: {exc!r}")
        self.assertTrue(out is None or isinstance(out, str))

    def test_parse_retry_hint_huge_ratelimit_reset_does_not_raise(self):
        headers = {"x-ratelimit-reset": str(10**200)}
        try:
            out = updater._parse_retry_hint(headers)
        except Exception as exc:  # noqa: BLE001
            self.fail(f"_parse_retry_hint must not raise, but raised: {exc!r}")
        self.assertTrue(out is None or isinstance(out, str))

    def test_parse_retry_hint_invalid_retry_after_with_reset_prefers_reset(self):
        with patch.object(updater.time, "time", return_value=1000):
            out = updater._parse_retry_hint(
                {
                    "retry-after": "9" * 200,
                    "x-ratelimit-reset": "1060",
                }
            )
        self.assertIsInstance(out, str)
        self.assertRegex(out, r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")

    def test_parse_retry_hint_non_numeric_retry_after_without_reset_returns_fallback(self):
        headers = {"retry-after": "Wed, 21 Oct 2015 07:28:00 GMT"}
        out = updater._parse_retry_hint(headers)
        self.assertEqual(out, headers["retry-after"][:64])

    def test_parse_retry_hint_non_numeric_retry_after_prefers_reset_when_present(self):
        with patch.object(updater.time, "time", return_value=1000):
            headers = {
                "retry-after": "Wed, 21 Oct 2015 07:28:00 GMT",
                "x-ratelimit-reset": "1060",
            }
            out = updater._parse_retry_hint(headers)
        self.assertIsInstance(out, str)
        self.assertRegex(out, r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")
        self.assertNotEqual(out, headers["retry-after"])

    def test_classify_timeout_socket(self):
        kind, params = updater._classify_update_error(socket.timeout("t"))
        self.assertEqual(kind, "timeout")
        self.assertEqual(params, {})

    def test_classify_network_urlerror(self):
        kind, params = updater._classify_update_error(URLError(OSError("dns")))
        self.assertEqual(kind, "network")
        self.assertEqual(params, {})

    def test_classify_rate_limited_429_always(self):
        err = _mk_http_error(429, {})
        kind, params = updater._classify_update_error(err)
        self.assertEqual(kind, "rate_limited")
        self.assertIsNone(params.get("retry_at"))

    def test_classify_rate_limited_403_remaining_zero(self):
        err = _mk_http_error(403, {"x-ratelimit-remaining": "0"})
        kind, params = updater._classify_update_error(err)
        self.assertEqual(kind, "rate_limited")
        self.assertIsNone(params.get("retry_at"))

    def test_classify_http_403_without_rate_limit_signals(self):
        err = _mk_http_error(403, {})
        kind, params = updater._classify_update_error(err)
        self.assertEqual(kind, "http")
        self.assertEqual(params.get("status"), 403)

    def test_retry_hint_falls_back_to_x_ratelimit_reset_when_retry_after_invalid(self):
        reset_ts = int(time.time()) + 60
        err = _mk_http_error(
            429,
            {"retry-after": "9" * 200, "x-ratelimit-reset": str(reset_ts)},
        )
        kind, params = updater._classify_update_error(err)
        self.assertEqual(kind, "rate_limited")
        self.assertNotEqual(params.get("retry_at"), "unknown")

    def test_classify_json_decode_error(self):
        err = json.JSONDecodeError("bad json", "{}", 1)
        kind, params = updater._classify_update_error(err)
        self.assertEqual(kind, "json")
        self.assertEqual(params, {})

    # --- airbag ---

    def test_build_message_never_raises_even_if_retry_hint_parser_breaks(self):
        err = _mk_http_error(429, {"retry-after": "10"})
        with patch.object(updater, "_parse_retry_hint", side_effect=OverflowError("boom")):
            msg = updater._build_update_error_message(err)
        self.assertIsInstance(msg, str)
        self.assertTrue(msg)


class UpdaterI18NTests(unittest.TestCase):
    def test_detect_updater_lang_reads_ui_lang_from_config(self):
        with patch.object(updater, "_load_app_config", return_value={"ui_lang": "en"}):
            self.assertEqual(updater._detect_updater_lang(), "en")

    def test_detect_updater_lang_defaults_to_polish(self):
        with patch.object(updater, "_load_app_config", return_value={}):
            self.assertEqual(updater._detect_updater_lang(), "pl")

    def test_detect_updater_lang_reads_from_config_file(self):
        cfg_path = updater.APP_CONFIG_PATH
        tmp_path = cfg_path.with_suffix(f"{cfg_path.suffix}.test")
        previous_path = updater.APP_CONFIG_PATH
        try:
            updater.APP_CONFIG_PATH = tmp_path
            tmp_path.write_text(json.dumps({"ui_lang": "en"}), encoding="utf-8")
            self.assertEqual(updater._detect_updater_lang(), "en")
        finally:
            updater.APP_CONFIG_PATH = previous_path
            if tmp_path.exists():
                tmp_path.unlink()

    def test_build_update_error_message_respects_english_language(self):
        err = _mk_http_error(429, {"retry-after": "10"})
        previous = updater._UPDATER_LANG
        updater._UPDATER_LANG = "en"
        try:
            with patch.object(updater.time, "time", return_value=1000):
                msg = updater._build_update_error_message(err)
        finally:
            updater._UPDATER_LANG = previous
        self.assertIn("rate limiting", msg.lower())
        self.assertIn("HTTP 429", msg)

    def test_build_update_error_message_respects_polish_language(self):
        err = _mk_http_error(429, {"retry-after": "10"})
        previous = updater._UPDATER_LANG
        updater._UPDATER_LANG = "pl"
        try:
            with patch.object(updater.time, "time", return_value=1000):
                msg = updater._build_update_error_message(err)
        finally:
            updater._UPDATER_LANG = previous
        self.assertIn("limit zapytań", msg.lower())
        self.assertIn("HTTP 429", msg)


if __name__ == "__main__":
    unittest.main()
