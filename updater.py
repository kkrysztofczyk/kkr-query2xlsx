import argparse
import json
import logging
import os
import socket
import shutil
import subprocess
import sys
import tempfile
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from zipfile import ZipFile

APP_EXE_NAME = "kkr-query2xlsx.exe"
UPDATER_EXE_NAME = "kkr-query2xlsx-updater.exe"
UPDATER_STAGED_EXE_NAME = "kkr-query2xlsx-updater.new.exe"
APP_CONFIG_NAME = "kkr-query2xlsx.json"
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

_UPD_MAX_RETRY_AFTER_SECONDS = 7 * 24 * 60 * 60
_UPD_MAX_RESET_FUTURE_SECONDS = 365 * 24 * 60 * 60

I18N_UPDATER: dict[str, dict[str, str]] = {
    "en": {
        "TITLE": "Update",
        "UPD_ERR_NETWORK": "Could not reach GitHub (network/DNS problem). Check your connection and try again.",
        "UPD_ERR_TIMEOUT": "GitHub did not respond in time (timeout). Try again in a moment.",
        "UPD_ERR_HTTP": "GitHub returned HTTP {status} while checking updates.",
        "UPD_ERR_RATE_LIMITED": "GitHub returned HTTP {status}. It looks like rate limiting. Try again after: {retry_at}.",
        "UPD_ERR_JSON": "Received an invalid response from GitHub (JSON parse error).",
        "UPD_ERR_GENERIC": "Failed to check or download updates.",
        "UPD_ERR_BUNDLE_ONLY": "Updater works only for the EXE package (Release ZIP). For source code use `git pull` or download a newer ZIP.",
        "UPD_ERR_WAIT_CLOSE_TIMEOUT": "Timed out waiting for the app to close.",
        "UPD_ERR_CLOSE_APP_FIRST": "Close the application before updating.",
        "UPD_ERR_NO_WINDOWS_ZIP": "No Windows ZIP asset found in latest release.",
        "UPD_ERR_MISSING_DOWNLOAD_URL": "Release asset is missing a download URL.",
        "UPD_ERR_UNPACK": "Failed to unpack update package.",
        "UPD_ERR_LAYOUT": "Update package layout not recognized.",
        "UPD_ERR_APPLY": "Failed to apply update.",
        "UPD_INFO_DONE": "Updated to {version}.",
        "UPD_LATEST_VERSION_LABEL": "the latest version",
        "UPD_ASK_LAUNCH": "Launch the application?",
        "UPD_ERR_LAUNCH": "Failed to launch the application.",
        "UPD_UNKNOWN": "unknown",
    },
    "pl": {
        "TITLE": "Aktualizacja",
        "UPD_ERR_NETWORK": "Brak połączenia z GitHub (problem sieci/DNS). Sprawdź połączenie i spróbuj ponownie.",
        "UPD_ERR_TIMEOUT": "GitHub nie odpowiedział na czas (timeout). Spróbuj ponownie za chwilę.",
        "UPD_ERR_HTTP": "GitHub zwrócił HTTP {status} podczas sprawdzania aktualizacji.",
        "UPD_ERR_RATE_LIMITED": "GitHub zwrócił HTTP {status}. Wygląda na limit zapytań. Spróbuj ponownie po: {retry_at}.",
        "UPD_ERR_JSON": "Otrzymano nieprawidłową odpowiedź z GitHub (błąd parsowania JSON).",
        "UPD_ERR_GENERIC": "Nie udało się sprawdzić lub pobrać aktualizacji.",
        "UPD_ERR_BUNDLE_ONLY": "Updater działa tylko dla paczki EXE (Release ZIP). Dla źródeł użyj `git pull` lub pobierz nowszy ZIP.",
        "UPD_ERR_WAIT_CLOSE_TIMEOUT": "Przekroczono czas oczekiwania na zamknięcie aplikacji.",
        "UPD_ERR_CLOSE_APP_FIRST": "Zamknij aplikację przed aktualizacją.",
        "UPD_ERR_NO_WINDOWS_ZIP": "Nie znaleziono paczki ZIP dla Windows w najnowszym wydaniu.",
        "UPD_ERR_MISSING_DOWNLOAD_URL": "Brakuje adresu URL do pobrania zasobu wydania.",
        "UPD_ERR_UNPACK": "Nie udało się rozpakować paczki aktualizacji.",
        "UPD_ERR_LAYOUT": "Nie rozpoznano struktury paczki aktualizacji.",
        "UPD_ERR_APPLY": "Nie udało się zastosować aktualizacji.",
        "UPD_INFO_DONE": "Zaktualizowano do {version}.",
        "UPD_LATEST_VERSION_LABEL": "najnowszej wersji",
        "UPD_ASK_LAUNCH": "Uruchomić aplikację?",
        "UPD_ERR_LAUNCH": "Nie udało się uruchomić aplikacji.",
        "UPD_UNKNOWN": "nieznane",
    },
}


def _normalize_ui_lang(lang: str | None) -> str | None:
    normalized = (lang or "").lower()
    return normalized if normalized in I18N_UPDATER else None


def _detect_updater_lang() -> str:
    app_cfg = _load_app_config()
    ui_lang = _normalize_ui_lang(app_cfg.get("ui_lang"))
    return ui_lang or "pl"


_UPDATER_LANG = "pl"


def t_upd(key: str, **kwargs) -> str:
    s = (
        I18N_UPDATER.get(_UPDATER_LANG, {}).get(key)
        or I18N_UPDATER["en"].get(key)
        or key
    )
    return s.format(**kwargs) if kwargs else s


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
    messagebox.showerror(t_upd("TITLE"), message)


def _show_info(message: str) -> None:
    LOGGER.info("%s", message)
    messagebox.showinfo(t_upd("TITLE"), message)


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


def _format_local_ts(ts: float) -> str | None:
    try:
        return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ts))
    except (OverflowError, OSError, ValueError):
        return None


def _parse_retry_hint(headers) -> str | None:  # noqa: ANN001
    if not headers:
        return None

    retry_after_fallback_str: str | None = None
    retry_after_ts: str | None = None

    try:
        retry_after = headers.get("retry-after")
    except Exception:  # noqa: BLE001
        retry_after = None
    if retry_after is not None:
        s = str(retry_after).strip()
        if s:
            if s.isdigit():
                try:
                    seconds = int(s)
                except (OverflowError, ValueError):
                    seconds = None
                if seconds is not None and 0 <= seconds <= _UPD_MAX_RETRY_AFTER_SECONDS:
                    retry_after_ts = _format_local_ts(time.time() + seconds)
            else:
                retry_after_fallback_str = s[:64]

    try:
        reset_raw = headers.get("x-ratelimit-reset")
    except Exception:  # noqa: BLE001
        reset_raw = None
    if reset_raw is not None:
        s = str(reset_raw).strip()
        if s.isdigit():
            try:
                reset_ts = int(s)
            except (OverflowError, ValueError):
                reset_ts = None
            if reset_ts is not None:
                now = int(time.time())
                if 0 <= reset_ts <= now + _UPD_MAX_RESET_FUTURE_SECONDS:
                    try:
                        dt = datetime.fromtimestamp(reset_ts)
                        return dt.strftime("%Y-%m-%d %H:%M:%S")
                    except (OverflowError, OSError, ValueError):
                        pass

    if retry_after_ts:
        return retry_after_ts
    return retry_after_fallback_str


def _classify_update_error(exc: Exception) -> tuple[str, dict]:
    if isinstance(exc, TimeoutError | socket.timeout):
        return "timeout", {}
    if isinstance(exc, HTTPError):
        status = exc.code
        headers = exc.headers or {}
        retry_at = _parse_retry_hint(headers)
        try:
            remaining = str(headers.get("x-ratelimit-remaining", "")).strip()
        except Exception:  # noqa: BLE001
            remaining = ""
        if status == 429:
            return "rate_limited", {"status": status, "retry_at": retry_at}
        if status == 403 and (retry_at or remaining == "0"):
            return "rate_limited", {"status": status, "retry_at": retry_at}
        return "http", {"status": status}
    if isinstance(exc, URLError):
        reason = getattr(exc, "reason", None)
        if isinstance(reason, TimeoutError | socket.timeout):
            return "timeout", {}
        return "network", {}
    if isinstance(exc, json.JSONDecodeError | UnicodeDecodeError | ValueError):
        return "json", {}
    return "unknown", {}


def _build_update_error_message(exc: Exception) -> str:
    try:
        kind, params = _classify_update_error(exc)
        if kind == "network":
            return t_upd("UPD_ERR_NETWORK")
        if kind == "timeout":
            return t_upd("UPD_ERR_TIMEOUT")
        if kind == "http":
            return t_upd("UPD_ERR_HTTP", status=params.get("status"))
        if kind == "rate_limited":
            retry_at = params.get("retry_at") or t_upd("UPD_UNKNOWN")
            status = params.get("status") or 429
            return t_upd(
                "UPD_ERR_RATE_LIMITED",
                status=status,
                retry_at=retry_at,
            )
        if kind == "json":
            return t_upd("UPD_ERR_JSON")
    except Exception:  # noqa: BLE001
        pass
    return t_upd("UPD_ERR_GENERIC")


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
    data = json.loads(payload)
    if not isinstance(data, dict):
        raise ValueError("GitHub release payload is not a JSON object")
    return data


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
        return t_upd("UPD_ERR_BUNDLE_ONLY")
    if not _looks_like_portable_bundle(install_root):
        return t_upd("UPD_ERR_BUNDLE_ONLY")
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
            shutil.rmtree(dst)
        shutil.copytree(src, dst, dirs_exist_ok=True)
        LOGGER.info("Updated dir: %s", dst)


def run_update(wait_pid: int | None = None) -> None:
    guard_message = _guard_install_root(BASE_DIR)
    if guard_message:
        _show_error(guard_message)
        return
    if wait_pid is not None:
        if not _wait_for_pid(wait_pid):
            _show_error(t_upd("UPD_ERR_WAIT_CLOSE_TIMEOUT"))
            return

    if _is_app_running():
        _show_error(t_upd("UPD_ERR_CLOSE_APP_FIRST"))
        return

    try:
        release = _fetch_latest_release()
    except Exception as exc:  # noqa: BLE001
        _show_error(_build_update_error_message(exc), exc)
        return

    asset = _select_windows_asset(release.get("assets") or [])
    if not asset:
        _show_error(t_upd("UPD_ERR_NO_WINDOWS_ZIP"))
        return
    download_url = asset.get("browser_download_url")
    latest_tag = release.get("tag_name") or ""
    if not download_url:
        _show_error(t_upd("UPD_ERR_MISSING_DOWNLOAD_URL"))
        return

    with tempfile.TemporaryDirectory(prefix="kkr-query2xlsx-update-") as temp_dir:
        temp_path = Path(temp_dir)
        zip_path = temp_path / "update.zip"
        try:
            _download_asset(download_url, zip_path)
        except Exception as exc:  # noqa: BLE001
            _show_error(_build_update_error_message(exc), exc)
            return

        extract_root = temp_path / "extract"
        extract_root.mkdir(parents=True, exist_ok=True)
        try:
            with ZipFile(zip_path) as zf:
                zf.extractall(extract_root)
        except Exception as exc:  # noqa: BLE001
            _show_error(t_upd("UPD_ERR_UNPACK"), exc)
            return

        bundle_root = _find_bundle_root(extract_root)
        if not bundle_root:
            _show_error(t_upd("UPD_ERR_LAYOUT"))
            return

        try:
            _update_files(bundle_root, BASE_DIR, latest_tag=latest_tag)
        except Exception as exc:  # noqa: BLE001
            _show_error(t_upd("UPD_ERR_APPLY"), exc)
            return

    _show_info(
        t_upd(
            "UPD_INFO_DONE",
            version=latest_tag or t_upd("UPD_LATEST_VERSION_LABEL"),
        )
    )
    if messagebox.askyesno(t_upd("TITLE"), t_upd("UPD_ASK_LAUNCH")):
        try:
            subprocess.Popen([str(BASE_DIR / APP_EXE_NAME)], cwd=BASE_DIR)  # noqa: S603
        except Exception as exc:  # noqa: BLE001
            _show_error(t_upd("UPD_ERR_LAUNCH"), exc)


def main() -> None:
    global _UPDATER_LANG
    _UPDATER_LANG = _detect_updater_lang()

    parser = argparse.ArgumentParser()
    parser.add_argument("--wait-pid", type=int)
    args = parser.parse_args()

    root = tk.Tk()
    root.withdraw()
    run_update(wait_pid=args.wait_pid)


if __name__ == "__main__":
    main()
