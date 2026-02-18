import importlib.machinery
import importlib.util
import unittest
from pathlib import Path
import tempfile


def load_app_module():
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    loader = importlib.machinery.SourceFileLoader("app_main_window_geometry", str(main_path))
    spec = importlib.util.spec_from_loader("app_main_window_geometry", loader)
    module = importlib.util.module_from_spec(spec)
    loader.exec_module(module)
    return module


class _FakeRoot:
    def __init__(self):
        self.last_geometry = None
        self._winfo_geometry = "1100x760+10+20"
        self.minsize_calls = []
        self.update_calls = 0
        self._vrootx = 0
        self._vrooty = 0
        self._vrootw = 1920
        self._vrooth = 1080
        self._screenw = 1920
        self._screenh = 1080

    def geometry(self, value):
        self.last_geometry = value

    def minsize(self, w, h):
        self.minsize_calls.append((w, h))

    def update_idletasks(self):
        self.update_calls += 1

    def winfo_geometry(self):
        return self._winfo_geometry

    def winfo_screenwidth(self):
        return self._screenw

    def winfo_screenheight(self):
        return self._screenh

    def winfo_vrootx(self):
        return self._vrootx

    def winfo_vrooty(self):
        return self._vrooty

    def winfo_vrootwidth(self):
        return self._vrootw

    def winfo_vrootheight(self):
        return self._vrooth


class MainWindowGeometryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_module()

    def test_apply_main_window_geometry_uses_default_for_tiny_geometry(self):
        root = _FakeRoot()

        self.app.apply_main_window_geometry(root, "200x100+0+0")

        self.assertEqual(root.last_geometry, "900x760+510+160")
        self.assertEqual(root.minsize_calls[-1], (self.app.MIN_MAIN_WINDOW_WIDTH, self.app.MIN_MAIN_WINDOW_HEIGHT))

    def test_apply_main_window_geometry_accepts_valid_geometry(self):
        root = _FakeRoot()

        self.app.apply_main_window_geometry(root, "1200x800+4+5")

        # apply clamps/normalizes; still should keep it unchanged here
        self.assertEqual(root.last_geometry, "1200x800+4+5")
        self.assertEqual(root.minsize_calls[-1], (self.app.MIN_MAIN_WINDOW_WIDTH, self.app.MIN_MAIN_WINDOW_HEIGHT))

    def test_apply_main_window_geometry_rejects_malformed_geometry(self):
        root = _FakeRoot()

        self.app.apply_main_window_geometry(root, "1200x800oops")

        self.assertEqual(root.last_geometry, "900x760+510+160")


    def test_apply_main_window_geometry_rejects_malformed_geometry_on_multimonitor(self):
        root = _FakeRoot()
        root._vrootw = 3840
        root._vrooth = 1080

        self.app.apply_main_window_geometry(root, "1200x800oops")

        self.assertEqual(root.last_geometry, "900x760+510+160")

    def test_apply_main_window_geometry_clamps_offscreen_offsets(self):
        root = _FakeRoot()

        # way off to the right; should clamp to (1920-1200)=720
        self.app.apply_main_window_geometry(root, "1200x800+4000+100")

        self.assertEqual(root.last_geometry, "1200x800+720+100")

    def test_apply_main_window_geometry_centers_on_primary_when_multimonitor(self):
        root = _FakeRoot()
        # simulate multi-monitor virtual desktop
        root._vrootw = 3840
        root._vrooth = 1080

        self.app.apply_main_window_geometry(root, "900x760+2500+100")

        self.assertEqual(root.last_geometry, "900x760+510+160")

    def test_main_window_geometry_to_save_returns_default_for_tiny_current_geometry(self):
        root = _FakeRoot()
        root._winfo_geometry = "300x200+0+0"

        saved = self.app.main_window_geometry_to_save(root)

        self.assertEqual(saved, self.app.DEFAULT_MAIN_WINDOW_GEOMETRY)

    def test_main_window_geometry_to_save_returns_geometry_when_valid(self):
        root = _FakeRoot()
        root._winfo_geometry = "1200x800+15+25"

        saved = self.app.main_window_geometry_to_save(root)

        self.assertEqual(saved, "1200x800+15+25")

    def test_load_ui_config_reads_window_geometry(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            (tmp_path / self.app.UI_CONFIG_FILENAME).write_text(
                '{"ui": {"window_geometry": "1200x800+1+2"}}',
                encoding="utf-8",
            )
            cfg = self.app.load_ui_config(tmp_path)
            self.assertEqual((cfg.get("ui") or {}).get("window_geometry"), "1200x800+1+2")

    def test_save_ui_config_persists_window_geometry(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            self.app.save_ui_config(tmp_path, {"ui": {"window_geometry": "1200x800+1+2"}})
            cfg = self.app.load_ui_config(tmp_path)
            self.assertEqual((cfg.get("ui") or {}).get("window_geometry"), "1200x800+1+2")

    def test_persist_main_window_geometry_does_not_overwrite_other_ui_settings(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            original = {
                "ui": {
                    "window_geometry": "1200x800+1+2",
                    "output_filename_stamp_enabled": True,
                    "output_filename_stamp_pattern": "[YYYY-MM-DD hh:mm]",
                    "output_filename_stamp_place": "prefix",
                    "sql_highlight_enabled": True,
                    "hide_data_dir_notice": True,
                }
            }
            self.app.save_ui_config(tmp_path, original)

            self.app.persist_main_window_geometry(tmp_path, "1300x850+7+8")

            cfg = self.app.load_ui_config(tmp_path)
            ui = cfg.get("ui") or {}
            self.assertEqual(ui.get("window_geometry"), "1300x850+7+8")
            self.assertTrue(ui.get("output_filename_stamp_enabled"))
            self.assertEqual(ui.get("output_filename_stamp_pattern"), "[YYYY-MM-DD hh:mm]")
            self.assertEqual(ui.get("output_filename_stamp_place"), "prefix")
            self.assertTrue(ui.get("sql_highlight_enabled"))
            self.assertTrue(ui.get("hide_data_dir_notice"))

if __name__ == "__main__":
    unittest.main()
