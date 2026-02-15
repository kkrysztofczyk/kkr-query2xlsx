import importlib.machinery
import importlib.util
import unittest
import os
import io
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import patch


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main", str(main_path))
    spec = importlib.util.spec_from_loader("app_main", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class BootstrapDataDirTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def test_required_work_dirs_contains_expected_entries(self):
        output_dir = "/sandbox/generated_reports"
        with patch.object(self.app, "_build_path", side_effect=lambda name: f"/sandbox/{name}"):
            expected = [output_dir, *[f"/sandbox/{subdir}" for subdir in self.app.WORKDIR_SUBDIRS]]
            self.assertEqual(
                self.app._required_work_dirs(output_dir),
                expected,
            )


    def test_select_startup_data_dir_prefers_base_when_base_has_markers(self):
        with patch.object(self.app, "has_data_markers", side_effect=[True, True]):
            selected = self.app.select_startup_data_dir("/base", "/user")
        self.assertEqual(selected, "/base")

    def test_select_startup_data_dir_falls_back_to_user_markers(self):
        with patch.object(self.app, "has_data_markers", side_effect=[False, True]):
            selected = self.app.select_startup_data_dir("/base", "/user")
        self.assertEqual(selected, "/user")

    def test_select_startup_data_dir_defaults_to_base_without_markers(self):
        with patch.object(self.app, "has_data_markers", side_effect=[False, False]):
            selected = self.app.select_startup_data_dir("/base", "/user")
        self.assertEqual(selected, "/base")

    def test_startup_ask_yes_no_returns_false_in_non_interactive_mode(self):
        class NonTtyStdin:
            @staticmethod
            def isatty():
                return False

        with patch("sys.stdin", NonTtyStdin()):
            self.assertFalse(self.app._startup_ask_yes_no("title", "message"))

    def test_startup_ask_yes_no_console_accepts_yes(self):
        class TtyStdin:
            @staticmethod
            def isatty():
                return True

        captured = io.StringIO()
        with redirect_stdout(captured), redirect_stderr(captured), patch("sys.stdin", TtyStdin()), patch(
            "builtins.input", return_value="y"
        ):
            self.assertTrue(
                self.app._startup_ask_yes_no(
                    "title",
                    "message",
                    mode="console",
                )
            )

    def test_startup_ask_yes_no_console_rejects_no(self):
        class TtyStdin:
            @staticmethod
            def isatty():
                return True

        captured = io.StringIO()
        with redirect_stdout(captured), redirect_stderr(captured), patch("sys.stdin", TtyStdin()), patch(
            "builtins.input", return_value="n"
        ):
            self.assertFalse(
                self.app._startup_ask_yes_no(
                    "title",
                    "message",
                    mode="console",
                )
            )

    def test_bootstrap_returns_primary_output_dir_when_first_attempt_succeeds(self):
        with patch.object(self.app, "_suggest_user_data_dir", return_value="/user"), patch.object(
            self.app, "select_startup_data_dir", return_value="/base"
        ), patch.object(self.app, "_set_data_dir") as set_data_dir_mock, patch.object(self.app, "_build_path", return_value="/base/generated_reports"), patch.object(
            self.app,
            "_ensure_required_work_dirs",
        ) as ensure_mock:
            output_dir = self.app.bootstrap_data_dir_and_workdirs_or_exit(prefer_gui_prompt=False)

        set_data_dir_mock.assert_called_once_with("/base")

        self.assertEqual(output_dir, "/base/generated_reports")
        ensure_mock.assert_called_once_with("/base/generated_reports")

    def test_bootstrap_first_failure_user_declines_fallback_exits(self):
        with patch.object(self.app, "_suggest_user_data_dir", return_value="/user"), patch.object(
            self.app, "select_startup_data_dir", return_value="/base"
        ), patch.object(self.app, "_set_data_dir"), patch.object(self.app, "_build_path", return_value="/base/generated_reports"), patch.object(
            self.app,
            "_ensure_required_work_dirs",
            side_effect=OSError("readonly"),
        ), patch.object(self.app, "_startup_ask_yes_no", return_value=False), patch.object(
            self.app,
            "_startup_show_error",
        ) as show_error_mock:
            with self.assertRaises(SystemExit) as ctx:
                self.app.bootstrap_data_dir_and_workdirs_or_exit(prefer_gui_prompt=False)

        self.assertEqual(ctx.exception.code, 1)
        show_error_mock.assert_called_once()
        self.assertEqual(show_error_mock.call_args.args[0], self.app.t("ERR_NO_WRITE_TITLE"))

    def test_bootstrap_first_failure_user_accepts_fallback_and_retry_succeeds(self):
        with patch.object(
            self.app,
            "_build_path",
            side_effect=["/base/generated_reports", "/user/generated_reports"],
        ), patch.object(
            self.app,
            "_ensure_required_work_dirs",
            side_effect=[OSError("readonly"), None],
        ) as ensure_mock, patch.object(
            self.app,
            "_suggest_user_data_dir",
            return_value="/user",
        ), patch.object(
            self.app, "select_startup_data_dir", return_value="/base"
        ), patch.object(
            self.app,
            "_startup_ask_yes_no",
            return_value=True,
        ) as ask_mock, patch.object(
            self.app,
            "_set_data_dir",
        ) as set_data_dir_mock, patch.object(
            self.app,
            "_startup_show_error",
        ) as show_error_mock:
            output_dir = self.app.bootstrap_data_dir_and_workdirs_or_exit(prefer_gui_prompt=False)

        self.assertEqual(output_dir, "/user/generated_reports")
        self.assertEqual(ensure_mock.call_count, 2)
        self.assertEqual(set_data_dir_mock.call_args_list[0].args, ("/base",))
        self.assertEqual(set_data_dir_mock.call_args_list[1].args, ("/user",))
        ask_mock.assert_called_once()
        self.assertEqual(ask_mock.call_args.args[0], self.app.t("ERR_NO_WRITE_TITLE"))
        show_error_mock.assert_not_called()

    def test_bootstrap_fallback_retry_failure_exits(self):
        with patch.object(
            self.app,
            "_build_path",
            side_effect=["/base/generated_reports", "/user/generated_reports"],
        ), patch.object(
            self.app,
            "_ensure_required_work_dirs",
            side_effect=[OSError("readonly"), OSError("still readonly")],
        ), patch.object(
            self.app,
            "_suggest_user_data_dir",
            return_value="/user",
        ), patch.object(
            self.app, "select_startup_data_dir", return_value="/base"
        ), patch.object(
            self.app,
            "_startup_ask_yes_no",
            return_value=True,
        ) as ask_mock, patch.object(
            self.app,
            "_set_data_dir",
        ) as set_data_dir_mock, patch.object(
            self.app,
            "_startup_show_error",
        ) as show_error_mock:
            with self.assertRaises(SystemExit) as ctx:
                self.app.bootstrap_data_dir_and_workdirs_or_exit(prefer_gui_prompt=False)

        self.assertEqual(ctx.exception.code, 1)
        self.assertEqual(set_data_dir_mock.call_args_list[0].args, ("/base",))
        self.assertEqual(set_data_dir_mock.call_args_list[1].args, ("/user",))
        ask_mock.assert_called_once()
        self.assertEqual(ask_mock.call_args.args[0], self.app.t("ERR_NO_WRITE_TITLE"))
        show_error_mock.assert_called_once()
        self.assertEqual(show_error_mock.call_args.args[0], self.app.t("ERR_NO_WRITE_TITLE"))


    def test_bootstrap_headless_first_failure_auto_fallback_without_prompt(self):
        with patch.object(
            self.app,
            "_build_path",
            side_effect=["/base/generated_reports", "/user/generated_reports"],
        ), patch.object(
            self.app,
            "_ensure_required_work_dirs",
            side_effect=[OSError("readonly"), None],
        ) as ensure_mock, patch.object(
            self.app,
            "_suggest_user_data_dir",
            return_value="/user",
        ), patch.object(
            self.app, "select_startup_data_dir", return_value="/base"
        ), patch.object(
            self.app,
            "_startup_ask_yes_no",
        ) as ask_mock, patch.object(
            self.app,
            "_set_data_dir",
        ) as set_data_dir_mock, patch.object(
            self.app,
            "_startup_show_error",
        ) as show_error_mock:
            output_dir = self.app.bootstrap_data_dir_and_workdirs_or_exit(
                prefer_gui_prompt=False,
                headless=True,
            )

        self.assertEqual(output_dir, "/user/generated_reports")
        self.assertEqual(ensure_mock.call_count, 2)
        self.assertEqual(set_data_dir_mock.call_args_list[0].args, ("/base",))
        self.assertEqual(set_data_dir_mock.call_args_list[1].args, ("/user",))
        ask_mock.assert_not_called()
        show_error_mock.assert_not_called()

    def test_bootstrap_headless_both_locations_fail_exits_without_tk(self):
        with patch.object(
            self.app,
            "_build_path",
            side_effect=["/base/generated_reports", "/user/generated_reports"],
        ), patch.object(
            self.app,
            "_ensure_required_work_dirs",
            side_effect=[OSError("readonly"), OSError("still readonly")],
        ), patch.object(
            self.app,
            "_suggest_user_data_dir",
            return_value="/user",
        ), patch.object(
            self.app, "select_startup_data_dir", return_value="/base"
        ), patch.object(
            self.app,
            "_startup_ask_yes_no",
        ) as ask_mock, patch.object(
            self.app,
            "_startup_show_error",
        ) as show_error_mock, patch.object(
            self.app,
            "_set_data_dir",
        ), patch(
            "builtins.print",
        ) as print_mock, patch(
            "sys.stderr",
            new_callable=io.StringIO,
        ) as stderr:
            with self.assertRaises(SystemExit) as ctx:
                self.app.bootstrap_data_dir_and_workdirs_or_exit(
                    prefer_gui_prompt=False,
                    headless=True,
                )

        self.assertEqual(ctx.exception.code, 1)
        ask_mock.assert_not_called()
        show_error_mock.assert_not_called()
        print_mock.assert_called_once()
        self.assertIs(print_mock.call_args.kwargs.get("file"), stderr)
        self.assertIn("Still cannot create working folders in", print_mock.call_args.args[0])
        no_write_title = self.app.t("ERR_NO_WRITE_TITLE")
        self.assertTrue(print_mock.call_args.args[0].startswith(f"{no_write_title}:\n"))

    def test_bootstrap_user_dir_selected_fails_then_base_succeeds_without_prompt(self):
        # When user_dir is auto-selected (markers) but it's not writable, we should try BASE_DIR.
        with patch.object(self.app, "BASE_DIR", "/base"), patch.object(
            self.app, "_suggest_user_data_dir", return_value="/user"
        ), patch.object(
            self.app, "select_startup_data_dir", return_value="/user"
        ), patch.object(
            self.app,
            "_build_path",
            side_effect=lambda name: os.path.join(self.app.DATA_DIR, name),
        ), patch.object(
            self.app, "_ensure_required_work_dirs",
            side_effect=[OSError("user readonly"), None],
        ) as ensure_mock, patch.object(
            self.app, "_startup_ask_yes_no"
        ) as ask_mock, patch.object(
            self.app, "_startup_show_error"
        ) as show_error_mock, patch.object(
            self.app, "_set_data_dir"
        ) as set_data_dir_mock:
            set_data_dir_mock.side_effect = lambda path: setattr(self.app, "DATA_DIR", os.path.abspath(path))
            output_dir = self.app.bootstrap_data_dir_and_workdirs_or_exit(prefer_gui_prompt=False)

        user_output = os.path.join(os.path.abspath("/user"), "generated_reports")
        expected = os.path.join(os.path.abspath("/base"), "generated_reports")
        self.assertEqual(output_dir, expected)
        self.assertEqual(ensure_mock.call_count, 2)
        self.assertEqual(ensure_mock.call_args_list[0].args, (user_output,))
        self.assertEqual(ensure_mock.call_args_list[1].args, (expected,))
        self.assertEqual(set_data_dir_mock.call_args_list[0].args, ("/user",))
        self.assertEqual(set_data_dir_mock.call_args_list[1].args, ("/base",))
        ask_mock.assert_not_called()
        show_error_mock.assert_not_called()

    def test_bootstrap_user_dir_selected_and_base_also_fails_exits_without_prompt(self):
        with patch.object(self.app, "BASE_DIR", "/base"), patch.object(
            self.app, "_suggest_user_data_dir", return_value="/user"
        ), patch.object(
            self.app, "select_startup_data_dir", return_value="/user"
        ), patch.object(
            self.app,
            "_build_path",
            side_effect=lambda name: os.path.join(self.app.DATA_DIR, name),
        ), patch.object(
            self.app, "_ensure_required_work_dirs",
            side_effect=[OSError("user readonly"), OSError("base readonly")],
        ) as ensure_mock, patch.object(
            self.app, "_startup_ask_yes_no"
        ) as ask_mock, patch.object(
            self.app, "_startup_show_error"
        ) as show_error_mock, patch.object(
            self.app, "_set_data_dir"
        ) as set_data_dir_mock:
            set_data_dir_mock.side_effect = lambda path: setattr(self.app, "DATA_DIR", os.path.abspath(path))
            with self.assertRaises(SystemExit) as ctx:
                self.app.bootstrap_data_dir_and_workdirs_or_exit(prefer_gui_prompt=False)

        user_output = os.path.join(os.path.abspath("/user"), "generated_reports")
        base_output = os.path.join(os.path.abspath("/base"), "generated_reports")
        self.assertEqual(ctx.exception.code, 1)
        self.assertEqual(ensure_mock.call_args_list[0].args, (user_output,))
        self.assertEqual(ensure_mock.call_args_list[1].args, (base_output,))
        self.assertEqual(set_data_dir_mock.call_args_list[0].args, ("/user",))
        self.assertEqual(set_data_dir_mock.call_args_list[1].args, ("/base",))
        ask_mock.assert_not_called()
        show_error_mock.assert_called_once()
        self.assertEqual(show_error_mock.call_args.args[0], self.app.t("ERR_NO_WRITE_TITLE"))


if __name__ == "__main__":
    unittest.main()
