import argparse
import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import time
import tkinter as tk
from pathlib import Path
from tkinter import messagebox
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from zipfile import ZipFile

APP_EXE_NAME = "kkr-query2xlsx.exe"
UPDATER_EXE_NAME = "kkr-query2xlsx-updater.exe"
UPDATER_STAGED_EXE_NAME = "kkr-query2xlsx-updater.new.exe"
APP_CONFIG_NAME = "kkr-query2xlsx.json"
UPDATER_TITLE = "Update"
GITHUB_REPO_OWNER = "kkrysztofczyk"
GITHUB_REPO_NAME = "kkr-query2xlsx"
GITHUB_RELEASES_LATEST_URL = (
    f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases/latest"
)


def _get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = _get_base_dir()
APP_CONFIG_PATH = BASE_DIR / APP_CONFIG_NAME
LOG_DIR = BASE_DIR / "logs"
LOG_PATH = LOG_DIR / "update.log"


def _setup_logger() -> logging.Logger:
    logger = logging.getLogger("kkr-query2xlsx-updater")
    logger.setLevel(logging.INFO)
    handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    handler.setFormatter(formatter)
    logger.handlers = [handler]
    return logger


def _get_logger() -> logging.Logger:
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        return _setup_logger()
    except Exception:
        temp_log = Path(tempfile.gettempdir()) / "kkr-query2xlsx-update.log"
        logger = logging.getLogger("kkr-query2xlsx-updater")
        logger.setLevel(logging.INFO)
        handler = logging.FileHandler(temp_log, encoding="utf-8")
        formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
        handler.setFormatter(formatter)
        logger.handlers = [handler]
        return logger


LOGGER = _get_logger()


def _load_app_config() -> dict:
    if not APP_CONFIG_PATH.exists():
        return {}
    try:
        with APP_CONFIG_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception as exc:  # noqa: BLE001
        LOGGER.warning("Failed to read %s: %s", APP_CONFIG_PATH, exc, exc_info=exc)
        return {}


def _save_app_config(cfg: dict) -> None:
    if not isinstance(cfg, dict):
        cfg = {}
    tmp_path = APP_CONFIG_PATH.with_suffix(f"{APP_CONFIG_PATH.suffix}.tmp")
    try:
        with tmp_path.open("w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, APP_CONFIG_PATH)
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def _set_pending_updater_update(
    staged_file_name: str, latest_tag: str = ""
) -> None:
    cfg = _load_app_config()
    updates = cfg.get("_updates")
    if not isinstance(updates, dict):
        updates = {}
    updates["pending_updater"] = {
        "file": staged_file_name,
        "tag": latest_tag or "",
        "ts": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
    }
    cfg["_updates"] = updates
    _save_app_config(cfg)


def _show_error(message: str, exc: Exception | None = None) -> None:
    if exc:
        LOGGER.error("%s: %s", message, exc, exc_info=exc)
    else:
        LOGGER.error("%s", message)
    messagebox.showerror(UPDATER_TITLE, message)


def _show_info(message: str) -> None:
    LOGGER.info("%s", message)
    messagebox.showinfo(UPDATER_TITLE, message)


def _select_windows_asset(assets: list[dict]) -> dict | None:
    for asset in assets:
        name = (asset.get("name") or "").lower()
        if "windows" in name and name.endswith(".zip"):
            return asset
    for asset in assets:
        name = (asset.get("name") or "").lower()
        if name.endswith(".zip"):
            return asset
    return None


def _fetch_latest_release() -> dict:
    req = Request(
        GITHUB_RELEASES_LATEST_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "kkr-query2xlsx-updater",
        },
    )
    with urlopen(req, timeout=20) as resp:  # noqa: S310
        payload = resp.read().decode("utf-8")
    return json.loads(payload)


def _download_asset(url: str, dest: Path) -> None:
    req = Request(url, headers={"User-Agent": "kkr-query2xlsx-updater"})
    with urlopen(req, timeout=30) as resp:  # noqa: S310
        with open(dest, "wb") as f:
            shutil.copyfileobj(resp, f)


def _find_bundle_root(root: Path) -> Path | None:
    for dirpath, dirnames, filenames in os.walk(root):
        if APP_EXE_NAME in filenames and "_internal" in dirnames:
            return Path(dirpath)
    return None


def _find_git_root(start: Path) -> Path | None:
    for candidate in [start, *start.parents]:
        if (candidate / ".git").exists():
            return candidate
    return None


def _looks_like_portable_bundle(root: Path) -> bool:
    return (root / APP_EXE_NAME).exists() and (root / "_internal").is_dir()


def _guard_install_root(install_root: Path) -> str | None:
    if _find_git_root(install_root):
        return (
            "Updater działa tylko dla paczki EXE (Release ZIP). "
            "Dla źródeł użyj `git pull` lub pobierz nowszy ZIP."
        )
    if not _looks_like_portable_bundle(install_root):
        return (
            "Updater działa tylko dla paczki EXE (Release ZIP). "
            "Dla źródeł użyj `git pull` lub pobierz nowszy ZIP."
        )
    return None


def _pid_exists(pid: int) -> bool:
    try:
        proc = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}"],
            check=False,
            capture_output=True,
            text=True,
        )
        return str(pid) in proc.stdout
    except Exception as exc:
        LOGGER.warning("PID check failed: %s", exc, exc_info=exc)
        return False


def _is_app_running() -> bool:
    try:
        proc = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {APP_EXE_NAME}"],
            check=False,
            capture_output=True,
            text=True,
        )
        return APP_EXE_NAME.lower() in proc.stdout.lower()
    except Exception as exc:
        LOGGER.warning("Process check failed: %s", exc, exc_info=exc)
        return False


def _wait_for_pid(pid: int, timeout_s: float = 60.0) -> bool:
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        if not _pid_exists(pid):
            return True
        time.sleep(0.5)
    return False


def _update_files(bundle_root: Path, install_root: Path, latest_tag: str = "") -> None:
    files_to_replace = [
        "kkr-query2xlsx.exe",
        UPDATER_EXE_NAME,
        "README.md",
        "LICENSE",
        "secure.sample.json",
        "queries.sample.txt",
    ]
    dirs_to_replace = [
        "_internal",
        "docs",
        "examples",
    ]

    for item in files_to_replace:
        src = bundle_root / item
        if not src.exists():
            continue
        dst = install_root / item
        if (
            item == UPDATER_EXE_NAME
            and getattr(sys, "frozen", False)
            and dst.resolve() == Path(sys.executable).resolve()
        ):
            staged_path = install_root / UPDATER_STAGED_EXE_NAME
            shutil.copy2(src, staged_path)
            _set_pending_updater_update(UPDATER_STAGED_EXE_NAME, latest_tag=latest_tag)
            LOGGER.info("Staged updater update: %s", staged_path)
            continue
        shutil.copy2(src, dst)
        LOGGER.info("Updated file: %s", dst)

    for item in dirs_to_replace:
        src = bundle_root / item
        if not src.exists():
            continue
        dst = install_root / item
        if dst.exists():
            shutil.rmtree(dst, ignore_errors=True)
        shutil.copytree(src, dst, dirs_exist_ok=True)
        LOGGER.info("Updated dir: %s", dst)


def run_update(wait_pid: int | None = None) -> None:
    guard_message = _guard_install_root(BASE_DIR)
    if guard_message:
        _show_error(guard_message)
        return
    if wait_pid is not None:
        if not _wait_for_pid(wait_pid):
            _show_error("Timed out waiting for the app to close.")
            return

    if _is_app_running():
        _show_error("Close the application before updating.")
        return

    try:
        release = _fetch_latest_release()
    except (HTTPError, URLError, json.JSONDecodeError) as exc:
        _show_error("Could not check latest release.", exc)
        return
    except Exception as exc:  # noqa: BLE001
        _show_error("Could not check latest release.", exc)
        return

    asset = _select_windows_asset(release.get("assets") or [])
    if not asset:
        _show_error("No Windows ZIP asset found in latest release.")
        return
    download_url = asset.get("browser_download_url")
    latest_tag = release.get("tag_name") or ""
    if not download_url:
        _show_error("Release asset is missing a download URL.")
        return

    with tempfile.TemporaryDirectory(prefix="kkr-query2xlsx-update-") as temp_dir:
        temp_path = Path(temp_dir)
        zip_path = temp_path / "update.zip"
        try:
            _download_asset(download_url, zip_path)
        except Exception as exc:  # noqa: BLE001
            _show_error("Failed to download update package.", exc)
            return

        extract_root = temp_path / "extract"
        extract_root.mkdir(parents=True, exist_ok=True)
        try:
            with ZipFile(zip_path) as zf:
                zf.extractall(extract_root)
        except Exception as exc:  # noqa: BLE001
            _show_error("Failed to unpack update package.", exc)
            return

        bundle_root = _find_bundle_root(extract_root)
        if not bundle_root:
            _show_error("Update package layout not recognized.")
            return

        try:
            _update_files(bundle_root, BASE_DIR, latest_tag=latest_tag)
        except Exception as exc:  # noqa: BLE001
            _show_error("Failed to apply update.", exc)
            return

    _show_info(f"Zaktualizowano do {latest_tag or 'najnowszej wersji'}.")
    if messagebox.askyesno(UPDATER_TITLE, "Uruchomić aplikację?"):
        try:
            subprocess.Popen([str(BASE_DIR / APP_EXE_NAME)], cwd=BASE_DIR)  # noqa: S603
        except Exception as exc:  # noqa: BLE001
            _show_error("Failed to launch the application.", exc)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--wait-pid", type=int)
    args = parser.parse_args()

    root = tk.Tk()
    root.withdraw()
    run_update(wait_pid=args.wait_pid)


if __name__ == "__main__":
    main()
