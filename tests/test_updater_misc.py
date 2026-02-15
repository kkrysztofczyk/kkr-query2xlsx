import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import updater


class UpdaterMiscTests(unittest.TestCase):
    def test_select_windows_asset_prefers_windows_zip(self):
        assets = [
            {
                "name": "kkr-query2xlsx-linux.zip",
                "browser_download_url": "https://example.invalid/linux",
            },
            {
                "name": "kkr-query2xlsx-windows.zip",
                "browser_download_url": "https://example.invalid/windows",
            },
        ]
        selected = updater._select_windows_asset(assets)
        self.assertIsNotNone(selected)
        self.assertEqual(selected.get("name"), "kkr-query2xlsx-windows.zip")

    def test_select_windows_asset_falls_back_to_first_zip(self):
        assets = [
            {
                "name": "kkr-query2xlsx-macos.zip",
                "browser_download_url": "https://example.invalid/macos",
            },
            {
                "name": "kkr-query2xlsx-linux.zip",
                "browser_download_url": "https://example.invalid/linux",
            },
        ]
        selected = updater._select_windows_asset(assets)
        self.assertIsNotNone(selected)
        self.assertEqual(selected.get("name"), "kkr-query2xlsx-macos.zip")

    def test_select_windows_asset_returns_none_when_no_zip(self):
        assets = [
            {"name": "README.txt"},
            {"name": "kkr-query2xlsx-windows.exe"},
        ]
        self.assertIsNone(updater._select_windows_asset(assets))

    def test_guard_install_root_blocks_git_checkout(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / ".git").mkdir()
            msg = updater._guard_install_root(root)
        self.assertEqual(msg, updater.t_upd("UPD_ERR_BUNDLE_ONLY"))

    def test_guard_install_root_blocks_non_bundle_layout(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            msg = updater._guard_install_root(root)
        self.assertEqual(msg, updater.t_upd("UPD_ERR_BUNDLE_ONLY"))

    def test_guard_install_root_allows_portable_bundle_layout(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / updater.APP_EXE_NAME).write_bytes(b"exe")
            (root / "_internal").mkdir()
            msg = updater._guard_install_root(root)
        self.assertIsNone(msg)

    def test_update_files_stages_updater_when_updating_self(self):
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            bundle_root = td_path / "bundle"
            install_root = td_path / "install"
            bundle_root.mkdir()
            install_root.mkdir()

            # Prepare bundle updater executable (new version)
            src_updater = bundle_root / updater.UPDATER_EXE_NAME
            src_updater.write_bytes(b"NEW")

            # Prepare currently-running updater executable path
            dst_updater = install_root / updater.UPDATER_EXE_NAME
            dst_updater.write_bytes(b"OLD")

            staged_path = install_root / updater.UPDATER_STAGED_EXE_NAME
            self.assertFalse(staged_path.exists())

            with patch.object(updater.sys, "executable", str(dst_updater)), patch.object(
                updater.sys, "frozen", True, create=True
            ), patch.object(
                updater,
                "_set_pending_updater_update",
            ) as mock_set_pending:
                updater._update_files(bundle_root, install_root, latest_tag="v1.2.3")

            # Updater should be staged, not overwritten in-place
            self.assertTrue(staged_path.exists())
            self.assertEqual(staged_path.read_bytes(), b"NEW")
            self.assertEqual(dst_updater.read_bytes(), b"OLD")
            mock_set_pending.assert_called_once_with(
                updater.UPDATER_STAGED_EXE_NAME,
                latest_tag="v1.2.3",
            )

    def test_update_files_replaces_internal_dir(self):
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            bundle_root = td_path / "bundle"
            install_root = td_path / "install"
            bundle_root.mkdir()
            install_root.mkdir()

            # Old runtime in install
            old_internal = install_root / "_internal"
            old_internal.mkdir(parents=True)
            (old_internal / "old.txt").write_text("OLD", encoding="utf-8")

            # New runtime in bundle
            new_internal = bundle_root / "_internal"
            new_internal.mkdir(parents=True)
            (new_internal / "new.txt").write_text("NEW", encoding="utf-8")

            updater._update_files(bundle_root, install_root, latest_tag="v9.9.9")

            self.assertTrue((install_root / "_internal").is_dir())
            self.assertFalse((install_root / "_internal" / "old.txt").exists())
            self.assertTrue((install_root / "_internal" / "new.txt").exists())
            self.assertEqual(
                (install_root / "_internal" / "new.txt").read_text(encoding="utf-8"),
                "NEW",
            )


if __name__ == "__main__":
    unittest.main()
