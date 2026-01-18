import csv
import json
import logging
import textwrap
import traceback
import os
import shutil
import subprocess
import sys
import time
import tkinter as tk
import tkinter.font as tkfont
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
from urllib.parse import quote_plus

from logging.handlers import RotatingFileHandler

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.exc import DBAPIError, NoSuchModuleError
from openpyxl import load_workbook
from openpyxl.utils import coordinate_to_tuple


BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _build_path(name: str) -> str:
    return os.path.join(BASE_DIR, name)


def _center_window(win, parent=None):
    """Center the window on the parent or screen, clamping to screen bounds."""
    win.update()
    w = win.winfo_width() or win.winfo_reqwidth()
    h = win.winfo_height() or win.winfo_reqheight()
    if parent and parent.winfo_ismapped():
        parent.update()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width() or parent.winfo_reqwidth()
        ph = parent.winfo_height() or parent.winfo_reqheight()
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
    else:
        vroot_x = win.winfo_vrootx()
        vroot_y = win.winfo_vrooty()
        vroot_w = win.winfo_vrootwidth()
        vroot_h = win.winfo_vrootheight()
        x = vroot_x + (vroot_w - w) // 2
        y = vroot_y + (vroot_h - h) // 2
    vroot_x = win.winfo_vrootx()
    vroot_y = win.winfo_vrooty()
    vroot_w = win.winfo_vrootwidth()
    vroot_h = win.winfo_vrootheight()
    if x + w > vroot_x + vroot_w:
        x = vroot_x + vroot_w - w
    if y + h > vroot_y + vroot_h:
        y = vroot_y + vroot_h - h
    x = max(x, vroot_x)
    y = max(y, vroot_y)
    win.geometry(f"+{x}+{y}")


SECURE_PATH = _build_path("secure.txt")
QUERIES_PATH = _build_path("queries.txt")
CSV_PROFILES_PATH = _build_path("csv_profiles.json")


BUILTIN_CSV_PROFILES = [
    {
        "name": "CSV standard (comma, dot)",
        "encoding": "utf-8-sig",
        "delimiter": ",",
        "delimiter_replacement": "",
        "decimal": ".",
        "lineterminator": "\n",
        "quotechar": '"',
        "quoting": "minimal",
        "escapechar": "",
        "doublequote": True,
        "date_format": "",
    },
    {
        "name": "CSV Excel Europe (semicolon, comma)",
        "encoding": "utf-8-sig",
        "delimiter": ";",
        "delimiter_replacement": "",
        "decimal": ",",
        "lineterminator": "\n",
        "quotechar": '"',
        "quoting": "minimal",
        "escapechar": "",
        "doublequote": True,
        "date_format": "",
    },
]

BUILTIN_CSV_PROFILE_NAMES = {p["name"] for p in BUILTIN_CSV_PROFILES}


def is_builtin_csv_profile(name: str) -> bool:
    return name in BUILTIN_CSV_PROFILE_NAMES


def shorten_path(path, max_len=80):
    if not path:
        return ""
    if len(path) <= max_len:
        return path
    head, tail = os.path.split(path)
    short = f"...{os.sep}{tail}"
    if len(short) > max_len:
        short = short[-max_len:]
    return short


def remove_bom(content: bytes) -> str:
    """
    Decode text from bytes, handling UTF-8/16/32 BOM if present.

    Falls back to UTF-8 when no BOM is detected and attempts legacy
    codepages when UTF-8 decoding fails (useful for queries saved with
    Windows encodings).
    """
    # UTF-8 with BOM
    if content.startswith(b"\xef\xbb\xbf"):
        return content[3:].decode("utf-8")

    # UTF-16 LE / BE with BOM
    if content.startswith(b"\xff\xfe") or content.startswith(b"\xfe\xff"):
        return content.decode("utf-16")

    # UTF-32 LE / BE with BOM
    if content.startswith(b"\x00\x00\xfe\xff") or content.startswith(b"\xff\xfe\x00\x00"):
        return content.decode("utf-32")

    # Default: attempt UTF-8 without BOM, then try common Windows encodings
    try:
        return content.decode("utf-8")
    except UnicodeDecodeError as exc:
        for fallback_encoding in ("cp1250", "cp1252"):
            try:
                LOGGER.warning(
                    "Failed to decode bytes as UTF-8. Falling back to %s.",
                    fallback_encoding,
                )
                return content.decode(fallback_encoding)
            except UnicodeDecodeError:
                continue

        LOGGER.error(
            "Failed to decode bytes with UTF-8 and Windows fallbacks. "
            "Replacing invalid bytes.",
            exc_info=exc,
        )
        return content.decode("utf-8", errors="replace")


def _normalize_connections(data):
    """Return normalized structure for saved connections with legacy support."""

    def _default_store():
        return {"connections": [], "last_selected": None}

    if isinstance(data, str):
        legacy_str = data.strip()
        if not legacy_str:
            return _default_store()
        name = "Domyślne MSSQL"
        return {
            "connections": [
                {
                    "name": name,
                    "type": "mssql_odbc",
                    "url": f"mssql+pyodbc:///?odbc_connect={quote_plus(legacy_str)}",
                    "details": {"odbc_connect": legacy_str},
                }
            ],
            "last_selected": name,
        }

    connections = []
    last_selected = None

    if isinstance(data, dict):
        last_selected = data.get("last_selected")
        for item in data.get("connections", []):
            name = str(item.get("name") or "").strip()
            url = str(item.get("url") or "").strip()
            conn_type = str(item.get("type") or "custom").strip() or "custom"
            details = item.get("details")
            if not name or not url:
                continue
            connections.append(
                {
                    "name": name,
                    "type": conn_type,
                    "url": url,
                    "details": details if isinstance(details, dict) else {},
                }
            )

    if not connections:
        return _default_store()

    names = {c["name"] for c in connections}
    if last_selected not in names:
        last_selected = connections[0]["name"]

    return {"connections": connections, "last_selected": last_selected}


def load_connections(path=SECURE_PATH):
    if not os.path.exists(path):
        return _normalize_connections({})

    with open(path, "r", encoding="utf-8") as file:
        raw = file.read().strip()

    if not raw:
        return _normalize_connections({})

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        data = raw

    return _normalize_connections(data)


def save_connections(store, path=SECURE_PATH):
    normalized = _normalize_connections(store)
    with open(path, "w", encoding="utf-8") as file:
        json.dump(normalized, file, ensure_ascii=False, indent=2)


def load_query_paths(queries_file=QUERIES_PATH):
    paths = []
    if not os.path.exists(queries_file):
        return paths

    with open(queries_file, "r", encoding="utf-8") as f:
        for line in f:
            path = line.strip()
            if path:
                paths.append(path)
    return paths


def save_query_paths(paths, queries_file=QUERIES_PATH):
    with open(queries_file, "w", encoding="utf-8") as f:
        for path in paths:
            f.write(f"{path}\n")


DEFAULT_CSV_PROFILE = {
    "name": "UTF-8 (comma)",
    "encoding": "utf-8",
    "delimiter": ",",
    "delimiter_replacement": "",
    "decimal": ".",
    "lineterminator": "\n",
    "quotechar": '"',
    "quoting": "minimal",
    "escapechar": "",
    "doublequote": True,
    "date_format": "",
}


# --- Logging setup ----------------------------------------------------------


def _setup_logger():
    """
    Prosty globalny logger zapisujący błędy do pliku logs/kkr_query2sheet.log.
    Log ma rotację (ok. 1 MB na plik, 3 backupy).
    """
    log_dir = os.path.join(BASE_DIR, "logs")
    os.makedirs(log_dir, exist_ok=True)

    logger = logging.getLogger("kkr_query2sheet")
    logger.setLevel(logging.INFO)

    # Nie dodawaj handlerów ponownie przy imporcie
    if not logger.handlers:
        log_path = os.path.join(log_dir, "kkr_query2sheet.log")
        handler = RotatingFileHandler(
            log_path,
            maxBytes=1_000_000,
            backupCount=3,
            encoding="utf-8",
        )
        formatter = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s - %(message)s"
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger


LOGGER = _setup_logger()
logger = logging.getLogger(__name__)
logger.setLevel(LOGGER.level)
if not logger.handlers:
    logger.handlers = LOGGER.handlers
    logger.propagate = False


def _log_unhandled_exception(exc_type, exc_value, exc_traceback):
    """Wszystkie nieobsłużone wyjątki lądują w logu.

    Ctrl+C zostawiamy w spokoju.
    """
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    LOGGER.critical(
        "Unhandled exception",
        exc_info=(exc_type, exc_value, exc_traceback),
    )


sys.excepthook = _log_unhandled_exception


def handle_db_driver_error(exc, db_type, profile_name=None, show_message=None):
    """Show user-friendly info for missing DB drivers and log details.

    Returns True when the error was recognized and presented to the user.
    """

    LOGGER.exception(
        "Database driver or library issue (type=%s, profile=%s)",
        db_type,
        profile_name,
        exc_info=exc,
    )

    show = show_message or messagebox.showerror
    exc_text = str(exc).lower()

    def _notify(title, msg):
        try:
            show(title, msg)
        except Exception:
            # In console mode, show may be a simple print function
            try:
                show(msg)
            except Exception:
                pass

    if db_type == "mssql_odbc":
        missing_pyodbc = isinstance(exc, (ImportError, ModuleNotFoundError)) and (
            getattr(exc, "name", "") == "pyodbc" or "pyodbc" in exc_text
        )
        missing_driver = any(
            signature in exc_text
            for signature in (
                "data source name not found",
                "driver not found",
                "no default driver",
                "odbc driver 17",
                "odbc driver 18",
                "native client",
                "unable to open lib",
            )
        )

        if isinstance(exc, NoSuchModuleError) and "pyodbc" in exc_text:
            missing_pyodbc = True

        if missing_pyodbc or missing_driver:
            msg = (
                "Nie można połączyć z SQL Server. Wymagany sterownik ODBC "
                "('ODBC Driver 17 for SQL Server' lub zgodny) lub biblioteka pyodbc "
                "nie jest zainstalowana. Zainstaluj sterownik i spróbuj ponownie."
            )
            _notify("Brak sterownika ODBC", msg)
            return True

    if db_type == "postgresql":
        missing_psycopg2 = isinstance(exc, (ImportError, ModuleNotFoundError)) and (
            getattr(exc, "name", "") == "psycopg2" or "psycopg2" in exc_text
        )
        if isinstance(exc, NoSuchModuleError) and "psycopg2" in exc_text:
            missing_psycopg2 = True

        if missing_psycopg2:
            msg = (
                "Nie można połączyć z PostgreSQL. Wymagana biblioteka Pythona (np. "
                "psycopg2) nie jest zainstalowana. Zainstaluj brakującą bibliotekę i "
                "spróbuj ponownie."
            )
            _notify("Brak biblioteki PostgreSQL", msg)
            return True

    return False


def _escape_visible(text):
    return (
        text.replace("\\", "\\\\")
        .replace("\n", "\\n")
        .replace("\r", "\\r")
        .replace("\t", "\\t")
    )


def _unescape_visible(text):
    return (
        text.replace("\\n", "\n")
        .replace("\\r", "\r")
        .replace("\\t", "\t")
        .replace("\\\\", "\\")
    )


def _normalize_user_csv_profiles(raw_profiles):
    if not isinstance(raw_profiles, list):
        raw_profiles = []

    normalized_profiles = []
    seen = set()
    for raw in raw_profiles:
        name = str(raw.get("name") or "").strip()
        if not name or name in seen or is_builtin_csv_profile(name):
            continue
        seen.add(name)
        normalized_profiles.append(
            {
                "name": name,
                "encoding": raw.get("encoding") or DEFAULT_CSV_PROFILE["encoding"],
                "delimiter": raw.get("delimiter") or DEFAULT_CSV_PROFILE["delimiter"],
                "delimiter_replacement": raw.get("delimiter_replacement", ""),
                "decimal": raw.get("decimal") or DEFAULT_CSV_PROFILE["decimal"],
                "lineterminator": raw.get("lineterminator")
                or DEFAULT_CSV_PROFILE["lineterminator"],
                "quotechar": raw.get("quotechar") or DEFAULT_CSV_PROFILE["quotechar"],
                "quoting": (raw.get("quoting") or DEFAULT_CSV_PROFILE["quoting"]).lower(),
                "escapechar": raw.get("escapechar", ""),
                "doublequote": bool(
                    raw.get("doublequote")
                    if raw.get("doublequote") is not None
                    else DEFAULT_CSV_PROFILE["doublequote"]
                ),
                "date_format": raw.get("date_format", ""),
            }
        )

    return normalized_profiles


def _merge_builtin_and_user_profiles(user_profiles):
    profiles_by_name = {p["name"]: dict(p) for p in BUILTIN_CSV_PROFILES}
    for prof in user_profiles:
        name = prof.get("name")
        if not name or is_builtin_csv_profile(name):
            continue
        profiles_by_name[name] = prof

    profiles = list(profiles_by_name.values())
    _sort_csv_profiles_in_place(profiles)
    return profiles


def _sort_csv_profiles_in_place(profiles):
    profiles.sort(
        key=lambda p: (0 if is_builtin_csv_profile(p["name"]) else 1, p["name"].lower())
    )


def _read_csv_profiles_from_file(path=CSV_PROFILES_PATH):
    if not os.path.exists(path):
        return []

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:  # noqa: BLE001
        logging.exception("Nie udało się odczytać csv_profiles.json")
        return []

    if isinstance(data, dict) and "profiles" in data:
        profiles = data.get("profiles") or []
    else:
        profiles = data if isinstance(data, list) else []

    return [p for p in profiles if p.get("name") not in BUILTIN_CSV_PROFILE_NAMES]


def get_all_csv_profiles(path=CSV_PROFILES_PATH):
    user_profiles = _normalize_user_csv_profiles(_read_csv_profiles_from_file(path))
    return _merge_builtin_and_user_profiles(user_profiles)


def _normalize_csv_config(data):
    profiles = []
    default_profile = None
    if isinstance(data, dict):
        default_profile = data.get("default_profile")
        profiles = data.get("profiles") or []
    elif isinstance(data, list):
        profiles = data

    user_profiles = _normalize_user_csv_profiles(profiles)
    merged_profiles = _merge_builtin_and_user_profiles(user_profiles)

    if not merged_profiles:
        merged_profiles = [DEFAULT_CSV_PROFILE.copy()]

    names = {p["name"] for p in merged_profiles}
    if default_profile not in names:
        default_profile = merged_profiles[0]["name"]

    return {"default_profile": default_profile, "profiles": merged_profiles}


def load_csv_profiles(path=CSV_PROFILES_PATH):
    data = {}
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:  # noqa: BLE001
            logging.exception("Nie udało się odczytać csv_profiles.json")

    config = _normalize_csv_config(data)
    config["profiles"] = get_all_csv_profiles(path)
    if config["default_profile"] not in {p["name"] for p in config["profiles"]}:
        config["default_profile"] = config["profiles"][0]["name"]

    return config


def save_csv_profiles(config, path=CSV_PROFILES_PATH):
    normalized = _normalize_csv_config(config)
    data_to_save = {
        "default_profile": normalized.get("default_profile"),
        "profiles": [
            p for p in normalized.get("profiles", []) if not is_builtin_csv_profile(p.get("name", ""))
        ],
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data_to_save, f, ensure_ascii=False, indent=2)


def ensure_directories(paths):
    for path in paths:
        os.makedirs(path, exist_ok=True)


def get_csv_profile(config, name):
    for profile in config.get("profiles", []):
        if profile.get("name") == name:
            return profile
    return None


def csv_profile_to_kwargs(profile):
    quoting_map = {
        "all": csv.QUOTE_ALL,
        "none": csv.QUOTE_NONE,
        "minimal": csv.QUOTE_MINIMAL,
        "nonnumeric": csv.QUOTE_NONNUMERIC,
    }

    quoting_value = quoting_map.get(
        (profile.get("quoting") or DEFAULT_CSV_PROFILE["quoting"]).lower(),
        csv.QUOTE_MINIMAL,
    )

    line_terminator_value = (
        profile.get("line_terminator")
        or profile.get("lineterminator")
        or DEFAULT_CSV_PROFILE["lineterminator"]
    )

    return {
        "sep": profile.get("delimiter") or DEFAULT_CSV_PROFILE["delimiter"],
        "encoding": profile.get("encoding") or DEFAULT_CSV_PROFILE["encoding"],
        "decimal": profile.get("decimal") or DEFAULT_CSV_PROFILE["decimal"],
        # pandas expects the keyword "lineterminator" (without underscore)
        "lineterminator": line_terminator_value,
        "quotechar": profile.get("quotechar") or DEFAULT_CSV_PROFILE["quotechar"],
        "quoting": quoting_value,
        "escapechar": profile.get("escapechar") or None,
        "doublequote": bool(
            profile.get("doublequote", DEFAULT_CSV_PROFILE["doublequote"])
        ),
        "date_format": profile.get("date_format") or None,
    }


def _run_query_to_rows(engine, sql_query):
    """
    Execute SQL with retry/deadlock handling and return:
    rows, columns, sql_duration, sql_start.
    """
    max_retries = 3
    last_exception = None

    for attempt in range(1, max_retries + 1):
        try:
            sql_start = time.perf_counter()
            with engine.connect() as connection:
                result = connection.execute(text(sql_query))

                if result.returns_rows:
                    rows = result.fetchall()
                    columns = result.keys()
                else:
                    rows = []
                    columns = []
            sql_end = time.perf_counter()
            sql_duration = sql_end - sql_start

            return rows, columns, sql_duration, sql_start

        except DBAPIError as e:
            msg = str(getattr(e, "orig", e))
            msg_lower = msg.lower()
            last_exception = e

            retryable_error_signatures = (
                # SQL Server deadlocks/serialization failures (existing behavior)
                "1205",
                "40001",
                "deadlocked on lock",
                # PostgreSQL messages for deadlocks/serialization failures
                "deadlock detected",
                "could not serialize access due to",
            )

            if any(signature in msg_lower for signature in retryable_error_signatures) and attempt < max_retries:
                wait_seconds = 2 ** attempt
                LOGGER.warning(
                    "Deadlock-like DBAPIError while executing SQL "
                    "(attempt %s/%s, waiting %s s). Query:\n%s",
                    attempt,
                    max_retries,
                    wait_seconds,
                    sql_query,
                )
                time.sleep(wait_seconds)
                continue

            LOGGER.exception(
                "DBAPIError while executing SQL. Query:\n%s",
                sql_query,
            )
            raise

        except Exception:
            last_exception = sys.exc_info()[1]
            LOGGER.exception(
                "Unexpected error while executing SQL. Query:\n%s",
                sql_query,
            )
            raise

    if last_exception:
        raise last_exception


def format_error_for_ui(exc: Exception, sql_query: str, max_chars: int = 2000) -> str:
    """
    Zapisuje pełny błąd do loga i zwraca skrócony komunikat do pokazania w UI.
    """
    # pełny traceback + SQL tylko do loga
    logger.exception("Błąd podczas wykonywania zapytania SQL. Query:\n%s", sql_query)

    full_tb = traceback.format_exc()
    first_line = full_tb.strip().splitlines()[0] if full_tb else str(exc)

    # pierwsza linia komunikatu z bazy
    db_msg_first_line = str(exc).splitlines()[0] if str(exc) else ""

    # SQL w jednej linii + skrócenie
    sql_one_line = " ".join(sql_query.split())
    sql_preview = textwrap.shorten(sql_one_line, width=600, placeholder=" ...")

    hints: list[str] = []
    if isinstance(exc, PermissionError):
        blocked_file = getattr(exc, "filename", "")
        shortened = shorten_path(blocked_file, max_len=200) if blocked_file else ""
        exists = os.path.exists(blocked_file) if blocked_file else False

        if exists:
            hints.append(
                "Plik docelowy już istnieje i może być otwarty w innej aplikacji "
                "(np. Excel). Zamknij go i spróbuj ponownie."
                + (f"\nŚcieżka: {shortened}" if shortened else "")
            )
        else:
            hints.append(
                "Brak uprawnień do zapisu pliku docelowego lub ścieżka jest "
                "niedostępna. Sprawdź lokalizację pliku."
            )

    msg = (
        f"{first_line}\n\n"
        f"Komunikat bazy (fragment):\n{db_msg_first_line}\n\n"
        f"SQL (początek):\n{sql_preview}\n\n"
        f"Pełny błąd zapisany w pliku kkr_query2sheet.log"
    )

    if hints:
        msg += "\n\nPodpowiedź:\n" + "\n".join(hints)

    if len(msg) > max_chars:
        msg = msg[:max_chars] + "\n...\n(Przycięto w UI, pełna treść w kkr_query2sheet.log)"

    return msg


def run_export(engine, sql_query, output_file_path, output_format, csv_profile=None):
    """Execute SQL, export the result, and return timing + row count details."""
    rows, columns, sql_duration, sql_start = _run_query_to_rows(engine, sql_query)

    export_duration = 0.0
    if rows:
        df = pd.DataFrame(rows, columns=columns)

        export_start = time.perf_counter()
        if output_format == "xlsx":
            df.to_excel(output_file_path, index=False)
        else:
            profile = csv_profile or DEFAULT_CSV_PROFILE
            delimiter = profile.get("delimiter") or DEFAULT_CSV_PROFILE["delimiter"]
            delimiter_replacement = profile.get("delimiter_replacement", "")

            export_df = df
            if delimiter and delimiter_replacement:
                # Intentionally replace delimiters globally in all string values
                # to match current CSV profile behaviour and avoid escaping.
                export_df = df.applymap(
                    lambda value: value.replace(delimiter, delimiter_replacement)
                    if isinstance(value, str)
                    else value
                )

            export_df.to_csv(output_file_path, index=False, **csv_profile_to_kwargs(profile))
        export_end = time.perf_counter()
        export_duration = export_end - export_start
        total_duration = export_end - sql_start
    else:
        total_duration = sql_duration

    return sql_duration, export_duration, total_duration, len(rows)


def run_export_to_template(
    engine,
    sql_query,
    template_path,
    output_file_path,
    sheet_name,
    start_cell,
    include_header,
):
    """
    Execute SQL, copy XLSX template and paste data into given sheet starting at start_cell.

    Returns the same tuple as run_export:
    (sql_duration, export_duration, total_duration, rows_count).
    """
    rows, columns, sql_duration, sql_start = _run_query_to_rows(engine, sql_query)

    export_start = time.perf_counter()
    # Zawsze kopiujemy template, nawet jeśli nie ma wierszy z SQL
    shutil.copyfile(template_path, output_file_path)

    rows_count = len(rows)
    if rows_count:
        df = pd.DataFrame(rows, columns=columns)

        wb = load_workbook(output_file_path)
        if sheet_name not in wb.sheetnames:
            wb.close()
            raise ValueError(f"Arkusz '{sheet_name}' nie istnieje w pliku template.")

        ws = wb[sheet_name]

        start_row, start_col = coordinate_to_tuple(start_cell)

        data_start_row = start_row
        if include_header:
            for c_offset, col_name in enumerate(df.columns):
                cell = ws.cell(row=start_row, column=start_col + c_offset)
                cell.value = col_name
            data_start_row = start_row + 1

        for r_offset, row in enumerate(df.itertuples(index=False)):
            for c_offset, value in enumerate(row):
                cell = ws.cell(
                    row=data_start_row + r_offset,
                    column=start_col + c_offset,
                )
                cell.value = value

        wb.save(output_file_path)

    export_end = time.perf_counter()
    export_duration = export_end - export_start
    total_duration = export_end - sql_start

    return sql_duration, export_duration, total_duration, rows_count


def run_console(engine, output_directory, selected_connection):
    sql_query_file_paths = load_query_paths()
    csv_config = load_csv_profiles()

    sql_query_file_path = None
    if sql_query_file_paths:
        print("Available SQL query files:")
        print("0: [Custom path]")
        for idx, path in enumerate(sql_query_file_paths, start=1):
            print(f"{idx}: {path}")

        while True:
            try:
                selection = int(
                    input(
                        "Please enter the number of the SQL query file to execute "
                        f"(0 for custom path, 1-{len(sql_query_file_paths)}): "
                    )
                )
                if selection == 0:
                    custom_path = input("Please enter full path to the .sql file: ").strip()
                    if not os.path.isfile(custom_path):
                        print("File does not exist. Please try again.")
                        continue
                    sql_query_file_path = custom_path
                    break
                if 1 <= selection <= len(sql_query_file_paths):
                    sql_query_file_path = sql_query_file_paths[selection - 1]
                    break
                print(f"Please enter a number between 0 and {len(sql_query_file_paths)}.")
            except ValueError:
                print("Invalid input. Please enter a number.")
    else:
        print("No SQL query file paths found in queries.txt")
        while True:
            custom_path = input("Please enter full path to the .sql file: ").strip()
            if os.path.isfile(custom_path):
                sql_query_file_path = custom_path
                break
            print("File does not exist. Please try again.")

    while True:
        output_format = input("Please enter the desired output format (xlsx or csv): ").strip().lower()
        if output_format in ["xlsx", "csv"]:
            break
        print("Invalid input. Please enter 'xlsx' or 'csv'.")

    selected_csv_profile = get_csv_profile(csv_config, csv_config.get("default_profile"))
    if output_format == "csv":
        profiles = csv_config.get("profiles", [])
        profile_names = [p.get("name") for p in profiles]
        default_profile_name = csv_config.get("default_profile") or profile_names[0]

        print("Available CSV profiles:")
        for idx, name in enumerate(profile_names, start=1):
            default_marker = " (default)" if name == default_profile_name else ""
            print(f"{idx}: {name}{default_marker}")

        while True:
            selection = input(
                "Enter CSV profile number to use or press Enter to use the default: "
            ).strip()
            if not selection:
                break
            if selection.isdigit():
                idx = int(selection)
                if 1 <= idx <= len(profile_names):
                    selected_csv_profile = profiles[idx - 1]
                    break
            print("Invalid selection. Please try again.")

    with open(sql_query_file_path, "rb") as file:
        content = file.read()

    sql_query = remove_bom(content).strip()
    if selected_connection.get("type") == "mssql_odbc" and sql_query:
        sql_query = (
            "SET ARITHABORT ON;\n"
            "SET NOCOUNT ON;\n"
            "SET ANSI_WARNINGS OFF;\n"
            + sql_query
        )

    base_name = os.path.basename(sql_query_file_path)
    output_file_name = os.path.splitext(base_name)[0] + (".xlsx" if output_format == "xlsx" else ".csv")
    output_file_path = os.path.join(output_directory, output_file_name)

    sql_dur, export_dur, total_dur, rows_count = run_export(
        engine, sql_query, output_file_path, output_format, csv_profile=selected_csv_profile
    )

    if rows_count > 0:
        print(f"Query results have been saved to: {output_file_path}")
    else:
        print("The query did not return any rows.")
    print(f"Data fetch time (SQL): {sql_dur:.2f} seconds")
    if rows_count > 0:
        print(f"Export time ({output_format}): {export_dur:.2f} seconds")
        print(f"Total time: {total_dur:.2f} seconds")


def _create_mssql_frame(parent):
    frame = tk.LabelFrame(parent, text="MSSQL (ODBC)")
    frame.grid(row=3, column=0, columnspan=4, sticky="we", padx=10, pady=(5, 0))
    for idx in range(4):
        frame.columnconfigure(idx, weight=1)

    driver_var = tk.StringVar(value="ODBC Driver 17 for SQL Server")
    server_var = tk.StringVar()
    database_var = tk.StringVar()
    username_var = tk.StringVar()
    password_var = tk.StringVar()
    trusted_var = tk.BooleanVar(value=False)
    encrypt_var = tk.BooleanVar(value=True)
    trust_cert_var = tk.BooleanVar(value=True)

    tk.Label(frame, text="Sterownik ODBC").grid(row=0, column=0, sticky="w", padx=5, pady=(5, 0))
    tk.Entry(frame, textvariable=driver_var, width=30).grid(
        row=0, column=1, columnspan=3, sticky="we", padx=5, pady=(5, 0)
    )

    tk.Label(frame, text="Serwer").grid(row=1, column=0, sticky="w", padx=5, pady=(5, 0))
    tk.Entry(frame, textvariable=server_var, width=30).grid(
        row=1, column=1, columnspan=3, sticky="we", padx=5, pady=(5, 0)
    )

    tk.Label(frame, text="Nazwa bazy").grid(row=2, column=0, sticky="w", padx=5, pady=(5, 0))
    tk.Entry(frame, textvariable=database_var, width=30).grid(
        row=2, column=1, columnspan=3, sticky="we", padx=5, pady=(5, 0)
    )

    tk.Label(frame, text="Login").grid(row=3, column=0, sticky="w", padx=5, pady=(5, 0))
    tk.Entry(frame, textvariable=username_var, width=25).grid(
        row=3, column=1, sticky="we", padx=5, pady=(5, 0)
    )

    tk.Label(frame, text="Hasło").grid(row=3, column=2, sticky="w", padx=5, pady=(5, 0))
    tk.Entry(frame, textvariable=password_var, show="*", width=25).grid(
        row=3, column=3, sticky="we", padx=5, pady=(5, 0)
    )

    tk.Checkbutton(
        frame, text="Logowanie Windows (Trusted_Connection)", variable=trusted_var
    ).grid(
        row=4, column=1, sticky="w", padx=5, pady=(5, 0)
    )
    tk.Checkbutton(frame, text="Encrypt", variable=encrypt_var).grid(
        row=4, column=2, sticky="w", padx=5, pady=(5, 0)
    )
    tk.Checkbutton(frame, text="TrustServerCertificate", variable=trust_cert_var).grid(
        row=4, column=3, sticky="w", padx=5, pady=(5, 0)
    )

    return frame, {
        "driver": driver_var,
        "server": server_var,
        "database": database_var,
        "username": username_var,
        "password": password_var,
        "trusted": trusted_var,
        "encrypt": encrypt_var,
        "trust_cert": trust_cert_var,
    }


def _create_pg_frame(parent):
    frame = tk.LabelFrame(parent, text="PostgreSQL")
    frame.grid(row=4, column=0, columnspan=4, sticky="we", padx=10, pady=(5, 0))
    for idx in range(4):
        frame.columnconfigure(idx, weight=1)

    host_var = tk.StringVar(value="localhost")
    port_var = tk.StringVar(value="5432")
    db_var = tk.StringVar()
    user_var = tk.StringVar()
    password_var = tk.StringVar()

    tk.Label(frame, text="Host").grid(row=0, column=0, sticky="w", padx=5, pady=(5, 0))
    tk.Entry(frame, textvariable=host_var).grid(
        row=0, column=1, sticky="we", padx=5, pady=(5, 0)
    )

    tk.Label(frame, text="Port").grid(row=0, column=2, sticky="w", padx=5, pady=(5, 0))
    tk.Entry(frame, textvariable=port_var, width=8).grid(
        row=0, column=3, sticky="we", padx=5, pady=(5, 0)
    )

    tk.Label(frame, text="Database").grid(row=1, column=0, sticky="w", padx=5, pady=(5, 0))
    tk.Entry(frame, textvariable=db_var).grid(
        row=1, column=1, sticky="we", padx=5, pady=(5, 0)
    )

    tk.Label(frame, text="User").grid(row=1, column=2, sticky="w", padx=5, pady=(5, 0))
    tk.Entry(frame, textvariable=user_var).grid(
        row=1, column=3, sticky="we", padx=5, pady=(5, 0)
    )

    tk.Label(frame, text="Password").grid(row=2, column=0, sticky="w", padx=5, pady=(5, 0))
    tk.Entry(frame, textvariable=password_var, show="*").grid(
        row=2, column=1, sticky="we", padx=5, pady=(5, 0)
    )

    return frame, {
        "host": host_var,
        "port": port_var,
        "database": db_var,
        "user": user_var,
        "password": password_var,
    }


def _create_sqlite_frame(parent):
    frame = tk.LabelFrame(parent, text="SQLite")
    frame.grid(row=5, column=0, columnspan=4, sticky="we", padx=10, pady=(5, 0))
    frame.columnconfigure(1, weight=1)

    path_var = tk.StringVar()

    tk.Label(frame, text="Ścieżka do pliku").grid(
        row=0, column=0, sticky="w", padx=5, pady=(5, 0)
    )
    tk.Entry(frame, textvariable=path_var).grid(
        row=0, column=1, sticky="we", padx=5, pady=(5, 0)
    )
    tk.Button(
        frame,
        text="Wybierz",
        command=lambda: path_var.set(
            filedialog.asksaveasfilename(
                title="Wybierz lub utwórz plik SQLite",
                defaultextension=".db",
                filetypes=[("SQLite", "*.db"), ("All files", "*.*")],
            )
        ),
    ).grid(row=0, column=2, padx=5, pady=(5, 0))

    return frame, {"path": path_var}


def _load_connection_details(conn_type, vars_by_type, details):
    details = details or {}
    if conn_type == "mssql_odbc":
        vars_by_type["driver"].set(details.get("driver", ""))
        vars_by_type["server"].set(details.get("server", ""))
        vars_by_type["database"].set(details.get("database", ""))
        vars_by_type["username"].set(details.get("username", ""))
        vars_by_type["password"].set(details.get("password", ""))
        vars_by_type["trusted"].set(details.get("trusted", False))
        vars_by_type["encrypt"].set(details.get("encrypt", True))
        vars_by_type["trust_cert"].set(details.get("trust_server_certificate", True))
    elif conn_type == "postgresql":
        vars_by_type["host"].set(details.get("host", ""))
        vars_by_type["port"].set(str(details.get("port", "5432")))
        vars_by_type["database"].set(details.get("database", ""))
        vars_by_type["user"].set(details.get("user", ""))
        vars_by_type["password"].set(details.get("password", ""))
    elif conn_type == "sqlite":
        vars_by_type["path"].set(details.get("path", ""))


def _build_connection_entry(conn_type, vars_by_type, name):
    base_entry = {"name": name, "type": conn_type, "url": "", "details": {}}
    if conn_type == "mssql_odbc":
        driver = vars_by_type["driver"].get().strip()
        server = vars_by_type["server"].get().strip()
        database = vars_by_type["database"].get().strip()
        username = vars_by_type["username"].get().strip()
        password = vars_by_type["password"].get()

        if not driver or not server or not database:
            messagebox.showerror(
                "Błąd danych",
                "Wypełnij: sterownik, serwer i nazwę bazy danych.",
            )
            return None

        parts = [
            f"DRIVER={{{driver}}}",
            f"SERVER={server}",
            f"DATABASE={database}",
        ]

        if vars_by_type["trusted"].get():
            parts.append("Trusted_Connection=yes")
        else:
            if not username or not password:
                messagebox.showerror(
                    "Błąd danych",
                    "Podaj login i hasło lub zaznacz logowanie Windows (Trusted_Connection).",
                )
                return None
            parts.extend([f"UID={username}", f"PWD={password}"])

        if vars_by_type["encrypt"].get():
            parts.append("Encrypt=yes")
        if vars_by_type["trust_cert"].get():
            parts.append("TrustServerCertificate=yes")

        connection_str = ";".join(parts)
        base_entry["url"] = f"mssql+pyodbc:///?odbc_connect={quote_plus(connection_str)}"
        base_entry["details"] = {
            "driver": driver,
            "server": server,
            "database": database,
            "username": username,
            "password": password,
            "trusted": vars_by_type["trusted"].get(),
            "encrypt": vars_by_type["encrypt"].get(),
            "trust_server_certificate": vars_by_type["trust_cert"].get(),
        }
        return base_entry

    if conn_type == "postgresql":
        host = vars_by_type["host"].get().strip()
        port = vars_by_type["port"].get().strip() or "5432"
        database = vars_by_type["database"].get().strip()
        user = vars_by_type["user"].get().strip()
        password = vars_by_type["password"].get()

        if not host or not database or not user:
            messagebox.showerror(
                "Błąd danych",
                "Wypełnij: host, nazwę bazy i użytkownika.",
            )
            return None

        base_entry["url"] = (
            f"postgresql+psycopg2://{quote_plus(user)}:{quote_plus(password)}@"
            f"{host}:{port}/{database}"
        )
        base_entry["details"] = {
            "host": host,
            "port": port,
            "database": database,
            "user": user,
            "password": password,
        }
        return base_entry

    db_path = vars_by_type["path"].get().strip()
    if not db_path:
        messagebox.showerror(
            "Błąd danych", "Wskaż ścieżkę do pliku bazy SQLite."
        )
        return None
    base_entry["url"] = f"sqlite:///{os.path.abspath(db_path)}"
    base_entry["details"] = {"path": os.path.abspath(db_path)}
    return base_entry


# Map internal DB types to user-friendly labels (UI-only)
DB_TYPE_LABELS = {
    "mssql_odbc": "SQL Server (ODBC)",
    "postgresql": "PostgreSQL",
    "sqlite": "SQLite (plik .db)",
}

DB_TYPE_BY_LABEL = {label: key for key, label in DB_TYPE_LABELS.items()}


def _build_connection_dialog_ui(root, selected_connection_var):
    dlg = tk.Toplevel(root)
    dlg.title("Dodaj lub zaktualizuj połączenie")
    dlg.transient(root)
    dlg.grab_set()

    for idx in range(4):
        dlg.columnconfigure(idx, weight=1)

    tk.Label(
        dlg,
        text=(
            "Aby utworzyć nowe połączenie, wpisz nową nazwę.\n"
            "Aby edytować istniejące połączenie, zostaw nazwę i popraw szczegóły."
        ),
        justify="left",
    ).grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(10, 0))

    tk.Label(dlg, text="Nazwa połączenia").grid(
        row=1, column=0, sticky="w", padx=10, pady=(10, 0)
    )
    name_var = tk.StringVar(value=selected_connection_var.get())
    tk.Entry(dlg, textvariable=name_var, width=40).grid(
        row=1, column=1, columnspan=3, sticky="we", padx=10, pady=(10, 0)
    )

    tk.Label(dlg, text="Typ bazy").grid(row=2, column=0, sticky="w", padx=10, pady=(5, 0))
    type_var = tk.StringVar(value="mssql_odbc")
    type_label_var = tk.StringVar(value=DB_TYPE_LABELS["mssql_odbc"])
    type_combo = ttk.Combobox(
        dlg,
        textvariable=type_label_var,
        values=list(DB_TYPE_LABELS.values()),
        state="readonly",
    )
    type_combo.grid(row=2, column=1, sticky="w", padx=10, pady=(5, 0))

    mssql_frame, mssql_vars = _create_mssql_frame(dlg)
    pg_frame, pg_vars = _create_pg_frame(dlg)
    sqlite_frame, sqlite_vars = _create_sqlite_frame(dlg)

    type_sections = {
        "mssql_odbc": (mssql_frame, mssql_vars),
        "postgresql": (pg_frame, pg_vars),
        "sqlite": (sqlite_frame, sqlite_vars),
    }

    def show_type_frame(*_):  # noqa: ANN001
        for frame, _ in type_sections.values():
            frame.grid_remove()
        frame, _ = type_sections.get(type_var.get(), (mssql_frame, mssql_vars))
        frame.grid()

    def update_type_from_label(*_):  # noqa: ANN001
        type_var.set(DB_TYPE_BY_LABEL.get(type_label_var.get(), "mssql_odbc"))
        show_type_frame()

    show_type_frame()
    type_var.trace_add("write", show_type_frame)
    type_label_var.trace_add("write", update_type_from_label)

    return dlg, name_var, type_var, type_label_var, type_sections, show_type_frame


def _load_existing_connection(
    name_var,
    type_var,
    type_label_var,
    type_sections,
    show_type_frame,
    selected_connection_var,
    get_connection_by_name,
):
    existing = get_connection_by_name(selected_connection_var.get())
    if not existing:
        return
    name_var.set(existing.get("name", ""))
    conn_type = existing.get("type", "mssql_odbc")
    type_var.set(conn_type)
    type_label_var.set(DB_TYPE_LABELS.get(conn_type, DB_TYPE_LABELS["mssql_odbc"]))
    section = type_sections.get(conn_type)
    if section:
        show_type_frame()
        _load_connection_details(conn_type, section[1], existing.get("details"))


def _build_and_test_connection_entry(
    name, conn_type, type_sections, create_engine_from_entry, handle_db_driver_error
):
    section = type_sections.get(conn_type)
    if not section:
        messagebox.showerror("Błąd danych", "Nieprawidłowy typ połączenia.")
        return None

    new_entry = _build_connection_entry(conn_type, section[1], name)
    if not new_entry:
        return None

    try:
        engine = create_engine_from_entry(new_entry)
        with engine.connect() as connection:
            connection.execute(text("SELECT 1"))
    except Exception as exc:  # noqa: BLE001
        if handle_db_driver_error(exc, conn_type, name):
            return None
        LOGGER.exception(
            "Connection creation failed for %s (%s)",
            name,
            conn_type,
            exc_info=exc,
        )
        messagebox.showerror(
            "Błąd połączenia",
            "Nie udało się połączyć przy użyciu podanych danych.\n\n"
            f"Szczegóły techniczne:\n{exc}",
        )
        return None

    return new_entry


def _replace_or_append_connection(connections_state, new_entry):
    replaced = False
    for idx, c in enumerate(connections_state["store"].get("connections", [])):
        if c.get("name") == new_entry.get("name"):
            connections_state["store"]["connections"][idx] = new_entry
            replaced = True
            break

    if not replaced:
        connections_state["store"].setdefault("connections", []).append(new_entry)


def _save_connection_without_test(
    name_var,
    type_var,
    type_sections,
    connections_state,
    set_selected_connection,
    persist_connections,
    refresh_connection_combobox,
):
    name = name_var.get().strip()
    if not name:
        messagebox.showerror("Błąd danych", "Nazwa połączenia nie może być pusta.")
        return False

    conn_type = type_var.get()
    section = type_sections.get(conn_type)
    if not section:
        messagebox.showerror("Błąd danych", "Nieprawidłowy typ połączenia.")
        return False

    new_entry = _build_connection_entry(conn_type, section[1], name)
    if not new_entry:
        return False

    _replace_or_append_connection(connections_state, new_entry)

    set_selected_connection(name)
    persist_connections()
    refresh_connection_combobox()

    messagebox.showinfo(
        "Zapisano",
        "Połączenie zapisane bez testu.\nUżyj przycisku „Testuj połączenie”, aby je sprawdzić.",
    )
    return True


def _save_connection_from_dialog(
    name_var,
    type_var,
    type_sections,
    connections_state,
    set_selected_connection,
    persist_connections,
    refresh_connection_combobox,
    apply_selected_connection,
    handle_db_driver_error,
    create_engine_from_entry,
):
    name = name_var.get().strip()
    if not name:
        messagebox.showerror("Błąd danych", "Nazwa połączenia nie może być pusta.")
        return False

    new_entry = _build_and_test_connection_entry(
        name, type_var.get(), type_sections, create_engine_from_entry, handle_db_driver_error
    )
    if not new_entry:
        return False

    _replace_or_append_connection(connections_state, new_entry)

    set_selected_connection(name)
    persist_connections()
    refresh_connection_combobox()
    apply_selected_connection(show_success=True)
    return True


def open_connection_dialog_gui(
    root,
    connections_state,
    selected_connection_var,
    get_connection_by_name,
    set_selected_connection,
    persist_connections,
    refresh_connection_combobox,
    apply_selected_connection,
    handle_db_driver_error,
    create_engine_from_entry,
):
    (
        dlg,
        name_var,
        type_var,
        type_label_var,
        type_sections,
        show_type_frame,
    ) = _build_connection_dialog_ui(root, selected_connection_var)

    _load_existing_connection(
        name_var,
        type_var,
        type_label_var,
        type_sections,
        show_type_frame,
        selected_connection_var,
        get_connection_by_name,
    )

    def on_save(*_):
        saved = _save_connection_from_dialog(
            name_var,
            type_var,
            type_sections,
            connections_state,
            set_selected_connection,
            persist_connections,
            refresh_connection_combobox,
            apply_selected_connection,
            handle_db_driver_error,
            create_engine_from_entry,
        )
        if saved:
            dlg.destroy()

    def on_save_without_test(*_):
        saved = _save_connection_without_test(
            name_var,
            type_var,
            type_sections,
            connections_state,
            set_selected_connection,
            persist_connections,
            refresh_connection_combobox,
        )
        if saved:
            dlg.destroy()

    def on_cancel(*_):
        dlg.destroy()

    button_frame = tk.Frame(dlg)
    button_frame.grid(row=6, column=0, columnspan=4, pady=10)

    tk.Button(button_frame, text="Zapisz", command=on_save, width=14).pack(
        side="left", padx=(0, 5)
    )
    tk.Button(
        button_frame, text="Zapisz bez testu", command=on_save_without_test, width=18
    ).pack(side="left", padx=(0, 5))
    tk.Button(button_frame, text="Anuluj", command=on_cancel, width=12).pack(side="left")

    dlg.bind("<Return>", on_save)
    dlg.bind("<Escape>", on_cancel)

    _center_window(dlg, root)


def _init_csv_profile_vars():
    return {
        "name": tk.StringVar(value=""),
        "encoding": tk.StringVar(value="utf-8"),
        "delimiter": tk.StringVar(value=","),
        "delimiter_replacement": tk.StringVar(value=""),
        "decimal": tk.StringVar(value="."),
        "lineterminator": tk.StringVar(value="\\n"),
        "quotechar": tk.StringVar(value='"'),
        "quoting": tk.StringVar(value="minimal"),
        "escapechar": tk.StringVar(value=""),
        "doublequote": tk.BooleanVar(value=True),
        "date_format": tk.StringVar(value=""),
        "date_preview": tk.StringVar(value=""),
    }


def _validate_date_format(raw):
    if not raw:
        example = datetime.now().isoformat(sep=" ", timespec="seconds")
        return (
            True,
            f"Domyślny format Pandas (przykład: {example})",
        )
    try:
        example = datetime.now().strftime(raw)
    except (ValueError, TypeError):
        return (
            False,
            "Nieprawidłowy wzorzec daty (użyj składni strftime, np. %Y-%m-%d).",
        )
    return True, f"Bieżący czas w tym formacie: {example}"


def _load_csv_profile(vars_dict, profile):
    vars_dict["name"].set(profile.get("name", ""))
    vars_dict["encoding"].set(profile.get("encoding", ""))
    vars_dict["delimiter"].set(profile.get("delimiter", ","))
    vars_dict["delimiter_replacement"].set(profile.get("delimiter_replacement", ""))
    vars_dict["decimal"].set(profile.get("decimal", "."))
    vars_dict["lineterminator"].set(_escape_visible(profile.get("lineterminator", "\\n")))
    vars_dict["quotechar"].set(profile.get("quotechar", '"'))
    vars_dict["quoting"].set((profile.get("quoting") or "minimal").lower())
    vars_dict["escapechar"].set(profile.get("escapechar", ""))
    vars_dict["doublequote"].set(bool(profile.get("doublequote", True)))
    vars_dict["date_format"].set(profile.get("date_format", ""))


def _read_csv_profile(vars_dict):
    valid_format, preview = _validate_date_format(vars_dict["date_format"].get())
    vars_dict["date_preview"].set(preview)
    if not valid_format:
        return None
    return {
        "name": vars_dict["name"].get().strip(),
        "encoding": vars_dict["encoding"].get().strip()
        or DEFAULT_CSV_PROFILE["encoding"],
        "delimiter": vars_dict["delimiter"].get() or DEFAULT_CSV_PROFILE["delimiter"],
        "delimiter_replacement": vars_dict["delimiter_replacement"].get(),
        "decimal": vars_dict["decimal"].get() or DEFAULT_CSV_PROFILE["decimal"],
        "lineterminator": _unescape_visible(
            vars_dict["lineterminator"].get() or DEFAULT_CSV_PROFILE["lineterminator"]
        ),
        "quotechar": vars_dict["quotechar"].get()
        or DEFAULT_CSV_PROFILE["quotechar"],
        "quoting": vars_dict["quoting"].get() or DEFAULT_CSV_PROFILE["quoting"],
        "escapechar": vars_dict["escapechar"].get(),
        "doublequote": bool(vars_dict["doublequote"].get()),
        "date_format": vars_dict["date_format"].get(),
    }


def _csv_field_help():
    return {
        "name": (
            "Nazwa profilu",
            "Dowolna, unikalna nazwa ułatwiająca wybór profilu, np. "
            '"UTF-8 (przecinek)" lub "Windows-1250 (średnik)".',
        ),
        "encoding": (
            "Kodowanie",
            "Sposób kodowania znaków w pliku CSV. Domyślnie UTF-8; dla "
            "starszych arkuszy Excel można użyć windows-1250.",
        ),
        "delimiter": (
            "Separator pól",
            "Znak oddzielający kolumny. Najczęściej przecinek (,) lub "
            "średnik (;), zgodnie z ustawieniami regionalnymi arkusza.",
        ),
        "delimiter_replacement": (
            "Zastąp separator w wartościach",
            "Opcjonalnie zamienia znak separatora w wartościach na inny (np. "
            "średnik na przecinek). Przydatne, gdy system importujący nie "
            "obsługuje poprawnego eskapowania separatorów w polach. "
            "Uwaga: zamiana jest globalna dla wszystkich pól tekstowych "
            "(również JSON/ID w formie tekstu).",
        ),
        "decimal": (
            "Separator dziesiętny",
            "Znak rozdzielający część całkowitą od ułamkowej. Kropka (.) "
            "dla układu angielskiego, przecinek (,) dla polskiego.",
        ),
        "lineterminator": (
            "Znak końca linii",
            "Domyślnie \n. Dla pełnej zgodności z Windows można użyć "
            "\r\n. Zmień tylko gdy import wymaga konkretnego formatu.",
        ),
        "quotechar": (
            "Znak cudzysłowu",
            'Najczęściej " . Używany do otaczania pól wymagających '
            "cytowania (np. zawierających separator).",
        ),
        "quoting": (
            "Strategia cudzysłowów",
            "minimal – tylko gdy potrzebne (zalecane), all – zawsze, "
            "nonnumeric – dla tekstu, none – bez cytowania (wymaga "
            "escapechar).",
        ),
        "escapechar": (
            "Separator w polu",
            "Znak ucieczki używany, gdy quoting=none lub pola mogą "
            "zawierać separator. Zostaw pusty, jeżeli stosujesz "
            "standardowe cytowanie.",
        ),
        "doublequote": (
            "Podwajanie cudzysłowów",
            'Gdy zaznaczone, wewnętrzny " w polu staje się "". '
            "Zostaw włączone, chyba że system importujący wymaga inaczej.",
        ),
        "date_format": (
            "Format daty",
            "Opcjonalny wzorzec strftime, np. %Y-%m-%d lub %d.%m.%Y. "
            "Pozostaw puste, aby użyć domyślnego formatowania Pandas.",
        ),
    }


def _build_csv_profile_list_ui(dlg):
    list_var = tk.StringVar(value=[])
    listbox = tk.Listbox(
        dlg,
        listvariable=list_var,
        width=40,
        height=8,
        activestyle="dotbox",
        exportselection=False,
    )
    listbox.grid(row=1, column=0, sticky="nsew", padx=(10, 5), pady=10)

    normal_font = tkfont.nametofont(listbox.cget("font"))
    bold_font = normal_font.copy()
    bold_font.configure(weight="bold")
    listbox._fonts = {"normal": normal_font, "bold": bold_font}

    scrollbar = tk.Scrollbar(dlg, orient="vertical", command=listbox.yview)
    listbox.config(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=1, column=1, sticky="ns", pady=10)
    return listbox, list_var


def _build_csv_profile_form_ui(form_frame, form_vars, field_help):
    def show_field_help(key):
        title, message = field_help.get(key, ("Informacja", ""))
        messagebox.showinfo(title, message)

    def add_info_button(row, key):
        tk.Button(
            form_frame, text="i", width=2, command=lambda k=key: show_field_help(k)
        ).grid(row=row, column=2, sticky="w", padx=(5, 0))

    date_preview_var = form_vars["date_preview"]

    def update_date_preview(*_args):  # noqa: ANN001
        valid, preview = _validate_date_format(form_vars["date_format"].get())
        date_preview_var.set(preview)
        return valid

    widgets = []

    tk.Label(form_frame, text="Nazwa profilu:").grid(row=0, column=0, sticky="w")
    name_entry = tk.Entry(form_frame, textvariable=form_vars["name"])
    name_entry.grid(
        row=0, column=1, sticky="we"
    )
    widgets.append((name_entry, "normal"))
    add_info_button(0, "name")

    tk.Label(form_frame, text="Kodowanie:").grid(row=1, column=0, sticky="w")
    encoding_entry = tk.Entry(form_frame, textvariable=form_vars["encoding"])
    encoding_entry.grid(
        row=1, column=1, sticky="we"
    )
    widgets.append((encoding_entry, "normal"))
    add_info_button(1, "encoding")

    tk.Label(form_frame, text="Separator pól:").grid(row=2, column=0, sticky="w")
    delimiter_entry = tk.Entry(form_frame, textvariable=form_vars["delimiter"], width=5)
    delimiter_entry.grid(
        row=2, column=1, sticky="w"
    )
    widgets.append((delimiter_entry, "normal"))
    add_info_button(2, "delimiter")

    tk.Label(form_frame, text="Zastąp separator w wartościach:").grid(
        row=3, column=0, sticky="w"
    )
    delimiter_replacement_entry = tk.Entry(
        form_frame, textvariable=form_vars["delimiter_replacement"], width=5
    )
    delimiter_replacement_entry.grid(
        row=3, column=1, sticky="w"
    )
    widgets.append((delimiter_replacement_entry, "normal"))
    add_info_button(3, "delimiter_replacement")

    tk.Label(form_frame, text="Separator dziesiętny:").grid(row=4, column=0, sticky="w")
    decimal_entry = tk.Entry(form_frame, textvariable=form_vars["decimal"], width=5)
    decimal_entry.grid(
        row=4, column=1, sticky="w"
    )
    widgets.append((decimal_entry, "normal"))
    add_info_button(4, "decimal")

    tk.Label(form_frame, text="Znak końca linii:").grid(row=5, column=0, sticky="w")
    lineterminator_entry = tk.Entry(
        form_frame, textvariable=form_vars["lineterminator"], width=10
    )
    lineterminator_entry.grid(
        row=5, column=1, sticky="w"
    )
    widgets.append((lineterminator_entry, "normal"))
    add_info_button(5, "lineterminator")

    tk.Label(form_frame, text="Znak cudzysłowu:").grid(row=6, column=0, sticky="w")
    quotechar_entry = tk.Entry(form_frame, textvariable=form_vars["quotechar"], width=5)
    quotechar_entry.grid(
        row=6, column=1, sticky="w"
    )
    widgets.append((quotechar_entry, "normal"))
    add_info_button(6, "quotechar")

    tk.Label(form_frame, text="Cudzysłów:").grid(row=7, column=0, sticky="w")
    quoting_combo = ttk.Combobox(
        form_frame,
        textvariable=form_vars["quoting"],
        values=["minimal", "all", "nonnumeric", "none"],
        state="readonly",
        width=15,
    )
    quoting_combo.grid(row=7, column=1, sticky="w")
    widgets.append((quoting_combo, "readonly"))
    add_info_button(7, "quoting")

    tk.Label(
        form_frame,
        text="Separator w polu:",
    ).grid(row=8, column=0, sticky="w")
    escapechar_entry = tk.Entry(form_frame, textvariable=form_vars["escapechar"], width=5)
    escapechar_entry.grid(
        row=8, column=1, sticky="w"
    )
    widgets.append((escapechar_entry, "normal"))
    add_info_button(8, "escapechar")
    tk.Label(
        form_frame,
        text="(znak ucieczki; puste = cytowanie)",
        fg="gray",
    ).grid(row=8, column=3, sticky="w", padx=(5, 0))

    doublequote_check = tk.Checkbutton(
        form_frame,
        text="Podwajaj cudzysłowy w polach",
        variable=form_vars["doublequote"],
    )
    doublequote_check.grid(row=9, column=0, columnspan=2, sticky="w")
    widgets.append((doublequote_check, "normal"))
    add_info_button(9, "doublequote")

    tk.Label(form_frame, text="Format daty:").grid(row=10, column=0, sticky="w")
    date_format_entry = tk.Entry(form_frame, textvariable=form_vars["date_format"])
    date_format_entry.grid(
        row=10, column=1, columnspan=1, sticky="we"
    )
    widgets.append((date_format_entry, "normal"))
    add_info_button(10, "date_format")
    tk.Label(
        form_frame,
        textvariable=date_preview_var,
        fg="gray",
    ).grid(row=11, column=0, columnspan=4, sticky="w", pady=(2, 0))
    form_vars["date_format"].trace_add("write", update_date_preview)
    update_date_preview()

    return update_date_preview, widgets


def _create_csv_profiles_dialog(root, csv_profile_state):
    dlg = tk.Toplevel(root)
    dlg.title("Profile CSV")
    dlg.transient(root)
    dlg.grab_set()
    dlg.resizable(True, True)

    working_profiles = [dict(p) for p in csv_profile_state["config"].get("profiles", [])]
    _sort_csv_profiles_in_place(working_profiles)
    default_profile_var = tk.StringVar(value=csv_profile_state["config"].get("default_profile"))
    display_default_var = tk.StringVar(
        value=f"Domyślny profil: {default_profile_var.get() or ''}"
    )

    dlg.columnconfigure(1, weight=1)
    dlg.rowconfigure(1, weight=1)

    tk.Label(dlg, textvariable=display_default_var, anchor="w").grid(
        row=0, column=0, columnspan=3, sticky="we", padx=10, pady=(10, 0)
    )

    return dlg, working_profiles, default_profile_var, display_default_var


def _refresh_csv_profile_list(
    listbox, list_var, working_profiles, default_profile_var, display_default_var
):
    display = []
    for prof in working_profiles:
        suffix = " (domyślny)" if prof["name"] == default_profile_var.get() else ""
        builtin_tag = " [wbudowany]" if is_builtin_csv_profile(prof["name"]) else ""
        display.append(f"{prof['name']}{builtin_tag}{suffix}")
    list_var.set(display)

    fonts = getattr(listbox, "_fonts", {})
    for idx, prof in enumerate(working_profiles):
        font_to_use = (
            fonts.get("bold") if is_builtin_csv_profile(prof["name"]) else fonts.get("normal")
        )
        if font_to_use is not None:
            listbox.itemconfig(idx, font=font_to_use)

    display_default_var.set(f"Domyślny profil: {default_profile_var.get() or ''}")


def _ensure_default_profile(default_profile_var, working_profiles, preferred_name=None):
    names = {p["name"] for p in working_profiles}
    if default_profile_var.get() in names:
        return
    if preferred_name and preferred_name in names:
        default_profile_var.set(preferred_name)
    else:
        default_profile_var.set(working_profiles[0]["name"])


def _read_profile_from_form_or_warn(form_vars):
    profile = _read_csv_profile(form_vars)
    if profile is None:
        messagebox.showwarning(
            "Uwaga",
            "Podany format daty jest nieprawidłowy. Skorzystaj ze składni strftime.",
        )
    elif not profile.get("name"):
        messagebox.showwarning("Uwaga", "Nazwa profilu nie może być pusta.")
        profile = None
    return profile


def _save_csv_profile_config(
    csv_profile_state, default_profile_var, working_profiles, refresh_csv_profile_controls
):
    config = {
        "default_profile": default_profile_var.get() or working_profiles[0]["name"],
        "profiles": working_profiles,
    }
    save_csv_profiles(config)
    csv_profile_state["config"] = load_csv_profiles()
    refresh_csv_profile_controls(csv_profile_state["config"].get("default_profile"))


def open_csv_profiles_manager_gui(
    root,
    csv_profile_state,
    selected_csv_profile_var,
    refresh_csv_profile_controls,
):
    current_profile_name = {"name": None}
    unsaved_changes = False

    dlg, working_profiles, default_profile_var, display_default_var = _create_csv_profiles_dialog(
        root, csv_profile_state
    )

    listbox, list_var = _build_csv_profile_list_ui(dlg)

    form_frame = tk.LabelFrame(dlg, text="Szczegóły profilu", padx=10, pady=10)
    form_frame.grid(row=1, column=2, sticky="nsew", padx=(5, 10), pady=10)
    form_frame.columnconfigure(1, weight=1)
    form_frame.columnconfigure(3, weight=1)

    form_vars = _init_csv_profile_vars()
    update_date_preview, form_widgets = _build_csv_profile_form_ui(
        form_frame, form_vars, _csv_field_help()
    )
    builtin_notice_var = tk.StringVar(value="")
    tk.Label(
        form_frame,
        textvariable=builtin_notice_var,
        fg="gray",
    ).grid(row=12, column=0, columnspan=4, sticky="w", pady=(2, 0))

    def refresh_list():
        _sort_csv_profiles_in_place(working_profiles)
        _refresh_csv_profile_list(
            listbox,
            list_var,
            working_profiles,
            default_profile_var,
            display_default_var,
        )

        selected_name = current_profile_name.get("name")
        if selected_name:
            for idx, prof in enumerate(working_profiles):
                if prof.get("name") == selected_name:
                    listbox.selection_clear(0, tk.END)
                    listbox.selection_set(idx)
                    listbox.see(idx)
                    break

    def set_editable_state(_is_builtin):
        for widget, normal_state in form_widgets:
            widget.configure(state=tk.DISABLED if _is_builtin else normal_state)

    def update_builtin_indicator(idx=None):
        is_builtin = False
        if idx is not None and 0 <= idx < len(working_profiles):
            is_builtin = is_builtin_csv_profile(working_profiles[idx].get("name", ""))
        builtin_notice_var.set(
            "Profil wbudowany: nie można zapisać zmian ani usuwać. Użyj Zapisz jako nowy, aby stworzyć własny wariant."
            if is_builtin
            else ""
        )
        set_editable_state(is_builtin)
        update_button.configure(state=tk.DISABLED if is_builtin else tk.NORMAL)
        delete_button.configure(state=tk.DISABLED if is_builtin else tk.NORMAL)

    def load_profile(idx):
        if idx < 0 or idx >= len(working_profiles):
            return
        _load_csv_profile(form_vars, working_profiles[idx])
        update_date_preview()
        update_builtin_indicator(idx)
        current_profile_name["name"] = working_profiles[idx].get("name")

    def sync_selection(event=None):  # noqa: ANN001
        sel = listbox.curselection()
        if sel:
            load_profile(sel[0])
        else:
            selected_name = current_profile_name.get("name")
            reselect_idx = None
            if selected_name:
                reselect_idx = next(
                    (
                        idx
                        for idx, prof in enumerate(working_profiles)
                        if prof.get("name") == selected_name
                    ),
                    None,
                )

            if reselect_idx is not None:
                listbox.selection_set(reselect_idx)
                listbox.see(reselect_idx)
                load_profile(reselect_idx)
            else:
                update_builtin_indicator()

    listbox.bind("<<ListboxSelect>>", sync_selection)

    def add_profile():
        nonlocal unsaved_changes
        prof = _read_profile_from_form_or_warn(form_vars)
        if not prof:
            return
        if is_builtin_csv_profile(prof["name"]):
            messagebox.showerror(
                "Błąd",
                "Ta nazwa jest zarezerwowana dla wbudowanego profilu. Wybierz inną nazwę.",
            )
            return
        if any(p["name"] == prof["name"] for p in working_profiles):
            messagebox.showwarning("Uwaga", "Profil o podanej nazwie już istnieje.")
            return
        working_profiles.append(prof)
        unsaved_changes = True
        refresh_list()
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(len(working_profiles) - 1)
        sync_selection()

    def update_profile():
        nonlocal unsaved_changes
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("Brak profilu", "Zaznacz profil na liście.")
            return
        prof = _read_profile_from_form_or_warn(form_vars)
        if not prof:
            return
        selected_profile = working_profiles[sel[0]]
        if is_builtin_csv_profile(prof["name"]):
            messagebox.showerror(
                "Błąd",
                "Nie możesz nadpisać wbudowanego profilu. Zmień nazwę i zapisz jako nowy profil.",
            )
            return
        for idx, existing in enumerate(working_profiles):
            if idx != sel[0] and existing["name"] == prof["name"]:
                messagebox.showwarning("Uwaga", "Profil o podanej nazwie już istnieje.")
                return
        if is_builtin_csv_profile(selected_profile.get("name", "")):
            working_profiles.append(prof)
            unsaved_changes = True
            refresh_list()
            new_idx = next(
                (i for i, p in enumerate(working_profiles) if p.get("name") == prof["name"]),
                None,
            )
            if new_idx is not None:
                listbox.selection_clear(0, tk.END)
                listbox.selection_set(new_idx)
            sync_selection()
            return

        working_profiles[sel[0]] = prof
        _ensure_default_profile(default_profile_var, working_profiles, prof["name"])
        unsaved_changes = True
        refresh_list()
        listbox.selection_set(sel[0])
        sync_selection()

    def delete_profile():
        nonlocal unsaved_changes
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("Brak profilu", "Zaznacz profil na liście.")
            return
        idx = sel[0]
        if is_builtin_csv_profile(working_profiles[idx].get("name", "")):
            messagebox.showinfo("Informacja", "Wbudowanych profili nie można usuwać.")
            return
        working_profiles.pop(idx)
        if not working_profiles:
            working_profiles.append(DEFAULT_CSV_PROFILE.copy())
        _ensure_default_profile(default_profile_var, working_profiles)
        unsaved_changes = True
        refresh_list()
        listbox.selection_set(0)
        sync_selection()

    def set_default_profile():
        nonlocal unsaved_changes
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("Brak profilu", "Zaznacz profil na liście.")
            return
        selected_name = working_profiles[sel[0]]["name"]
        if default_profile_var.get() != selected_name:
            default_profile_var.set(selected_name)
            unsaved_changes = True
            refresh_list()

    def save_and_close():
        nonlocal unsaved_changes
        if not working_profiles:
            messagebox.showwarning("Uwaga", "Musi istnieć co najmniej jeden profil CSV.")
            return
        _save_csv_profile_config(
            csv_profile_state, default_profile_var, working_profiles, refresh_csv_profile_controls
        )
        messagebox.showinfo(
            "Zapisano",
            "Profile CSV zapisane. Będą używane przy kolejnych eksportach.",
        )
        unsaved_changes = False
        dlg.destroy()

    def on_close():
        if not unsaved_changes:
            dlg.destroy()
            return

        resp = messagebox.askyesnocancel(
            "Niezapisane zmiany",
            "Masz niezapisane zmiany profili CSV. Zapisać przed zamknięciem?",
        )
        if resp is True:
            save_and_close()
        elif resp is False:
            dlg.destroy()

    button_frame = tk.Frame(dlg)
    button_frame.grid(row=2, column=0, columnspan=3, pady=(0, 10))

    tk.Button(button_frame, text="Zapisz jako nowy", command=add_profile, width=14).pack(
        side="left", padx=(0, 5)
    )
    update_button = tk.Button(
        button_frame, text="Zaktualizuj profil", command=update_profile, width=14
    )
    update_button.pack(side="left", padx=(0, 5))
    delete_button = tk.Button(button_frame, text="Usuń", command=delete_profile, width=10)
    delete_button.pack(side="left", padx=(0, 5))
    tk.Button(button_frame, text="Ustaw jako domyślny", command=set_default_profile, width=18).pack(
        side="left", padx=(0, 5)
    )
    tk.Button(button_frame, text="Zamknij i zapisz", command=save_and_close, width=14).pack(side="left")

    refresh_list()
    if working_profiles:
        preferred_name = selected_csv_profile_var.get() or default_profile_var.get()
        selected_idx = next(
            (
                idx
                for idx, prof in enumerate(working_profiles)
                if prof.get("name") == preferred_name
            ),
            0,
        )
        listbox.selection_set(selected_idx)
        listbox.see(selected_idx)
        sync_selection()
    else:
        update_builtin_indicator()

    # 1) Set keyboard focus to the list of profiles
    listbox.focus_set()

    # 2) Esc korzysta z on_close (uwzględnia niezapisane zmiany)
    dlg.bind("<Escape>", lambda *_: on_close())

    # 3) Profesjonalne centrowanie względem okna rodzica (aplikacji)
    dlg.update()
    _center_window(dlg, root)

    dlg.protocol("WM_DELETE_WINDOW", on_close)

    dlg.wait_window(dlg)


def run_gui(connection_store, output_directory):
    query_paths_state = {"paths": load_query_paths()}
    csv_profile_state = {"config": load_csv_profiles(), "combobox": None}
    connections_state = {
        "store": connection_store or {"connections": [], "last_selected": None},
        "combobox": None,
    }

    root = tk.Tk()
    root.title("KKr SQL to XLSX/CSV")

    selected_sql_path_full = tk.StringVar(value="")
    sql_label_var = tk.StringVar(value="")
    format_var = tk.StringVar(value="xlsx")
    selected_csv_profile_var = tk.StringVar(value="")
    default_csv_label_var = tk.StringVar(value="")
    result_info_var = tk.StringVar(value="")
    last_output_path = {"path": None}
    engine_holder = {"engine": None}
    connection_status_var = tk.StringVar(value="")
    secure_edit_state = {"button": None}
    start_button_holder = {"widget": None}
    error_display = {"widget": None}
    selected_connection_var = tk.StringVar(
        value=connections_state["store"].get("last_selected") or ""
    )

    # Template-related state (GUI only; console mode has no template support)
    use_template_var = tk.BooleanVar(value=False)
    template_path_var = tk.StringVar(value="")
    template_label_var = tk.StringVar(value="")
    sheet_name_var = tk.StringVar(value="")
    start_cell_var = tk.StringVar(value="A2")
    include_header_var = tk.BooleanVar(value=False)
    template_state = {
        "sheet_combobox": None,
        "choose_button": None,
        "start_cell_entry": None,
        "include_header_check": None,
    }

    def _set_sql_path(path):
        selected_sql_path_full.set(path)
        sql_label_var.set(shorten_path(path))

    def set_connection_status(message, connected):
        connection_status_var.set(message)
        btn = start_button_holder.get("widget")
        if btn is not None:
            btn_state = tk.NORMAL if connected else tk.DISABLED
            btn.config(state=btn_state)

    def apply_engine(new_engine):
        old_engine = engine_holder.get("engine")
        if old_engine is not None and old_engine is not new_engine:
            old_engine.dispose()
        engine_holder["engine"] = new_engine

    def get_connection_by_name(name):
        for conn in connections_state["store"].get("connections", []):
            if conn.get("name") == name:
                return conn
        return None

    def persist_connections():
        save_connections(connections_state["store"])
        refresh_secure_edit_button()

    def refresh_connection_combobox():
        combo = connections_state.get("combobox")
        if combo is None:
            return
        names = [c.get("name", "") for c in connections_state["store"].get("connections", [])]
        combo["values"] = names
        if selected_connection_var.get() not in names:
            selected_connection_var.set(names[0] if names else "")

    def set_selected_connection(name):
        if name and get_connection_by_name(name):
            connections_state["store"]["last_selected"] = name
            selected_connection_var.set(name)
            persist_connections()

    def on_connection_change(*_):  # noqa: ANN001
        name = selected_connection_var.get()
        if not name:
            set_connection_status("Brak połączenia. Utwórz nowe połączenie.", False)
            apply_engine(None)
            return
        set_selected_connection(name)
        apply_selected_connection(show_success=False)

    def create_engine_from_entry(entry):
        if not entry:
            raise ValueError("Brak konfiguracji połączenia")
        engine_kwargs = {}
        if entry.get("type") == "mssql_odbc":
            engine_kwargs["isolation_level"] = "AUTOCOMMIT"
        return create_engine(entry["url"], **engine_kwargs)

    def apply_selected_connection(show_success=False):
        conn = get_connection_by_name(selected_connection_var.get())
        if not conn:
            set_connection_status("Brak połączenia. Utwórz nowe połączenie.", False)
            apply_engine(None)
            return
        try:
            engine = create_engine_from_entry(conn)
            with engine.connect() as connection:
                connection.execute(text("SELECT 1"))
            apply_engine(engine)
        except Exception as exc:  # noqa: BLE001
            set_connection_status("Błąd połączenia. Utwórz nowe połączenie.", False)
            if handle_db_driver_error(exc, conn.get("type"), conn.get("name")):
                return
            LOGGER.exception(
                "Connection test failed for %s (%s)",
                conn.get("name"),
                conn.get("type"),
                exc_info=exc,
            )
            messagebox.showerror(
                "Błąd połączenia",
                (
                    "Nie udało się nawiązać połączenia.\n\n"
                    f"Szczegóły techniczne:\n{exc}"
                ),
            )
            return

        set_connection_status(
            f"Połączono z {conn.get('name', '')} ({conn.get('type', '')}).",
            True,
        )
        if show_success:
            messagebox.showinfo(
                "Połączenie działa", f"Połączenie {conn.get('name', '')} powiodło się."
            )

    def update_error_display(message):
        widget = error_display.get("widget")
        if widget is None:
            return
        widget.config(state="normal")
        widget.delete("1.0", tk.END)
        if message:
            widget.insert("1.0", message)
            widget.see(tk.END)
        widget.config(state="disabled")

    def show_error_popup(ui_msg):
        popup = tk.Toplevel(root)
        popup.title("Błąd zapytania")
        popup.transient(root)
        popup.attributes("-topmost", True)
        popup.grab_set()

        popup.geometry("900x450")
        _center_window(popup, root)

        text_widget = tk.Text(
            popup, wrap="none", width=100, height=25, font=("Consolas", 9)
        )
        y_scroll = tk.Scrollbar(popup, orient="vertical", command=text_widget.yview)
        x_scroll = tk.Scrollbar(popup, orient="horizontal", command=text_widget.xview)
        text_widget.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        text_widget.insert("1.0", ui_msg)
        text_widget.config(state="disabled")

        button_frame = tk.Frame(popup)

        def copy_error():
            popup.clipboard_clear()
            popup.clipboard_append(ui_msg)

        copy_btn = tk.Button(button_frame, text="Kopiuj", command=copy_error)
        close_btn = tk.Button(button_frame, text="Zamknij", command=popup.destroy)

        y_scroll.pack(side="right", fill="y")
        x_scroll.pack(side="bottom", fill="x")
        text_widget.pack(side="left", fill="both", expand=True)
        button_frame.pack(side="bottom", fill="x", pady=5)
        copy_btn.pack(side="right", padx=(0, 10))
        close_btn.pack(side="right")

        popup.bind("<Escape>", lambda *_: popup.destroy())
        popup.focus_set()

    def refresh_secure_edit_button():
        btn = secure_edit_state.get("button")
        if btn is None:
            return
        exists = os.path.exists(SECURE_PATH)
        btn_state = tk.NORMAL if exists else tk.DISABLED
        btn.config(state=btn_state)

    def test_connection_only():
        if not connections_state["store"].get("connections"):
            messagebox.showerror(
                "Brak połączenia",
                "Brak zapisanych połączeń. Utwórz i zapisz nowe połączenie.",
            )
            return
        apply_selected_connection(show_success=True)

    def delete_selected_connection():
        connections = connections_state["store"].get("connections", [])
        name = selected_connection_var.get()
        if not connections or not name:
            messagebox.showerror(
                "Brak połączenia",
                "Brak połączenia do usunięcia.",
            )
            return

        if not messagebox.askyesno(
            "Usuń połączenie", f"Czy na pewno chcesz usunąć połączenie {name}?"
        ):
            return

        connections_state["store"]["connections"] = [
            c for c in connections if c.get("name") != name
        ]

        remaining = connections_state["store"].get("connections", [])
        if remaining:
            new_selection = remaining[0].get("name", "")
            set_selected_connection(new_selection)
            refresh_connection_combobox()
            apply_selected_connection(show_success=False)
        else:
            selected_connection_var.set("")
            connections_state["store"]["last_selected"] = None
            persist_connections()
            refresh_connection_combobox()
            apply_engine(None)
            set_connection_status("Brak połączenia. Utwórz nowe połączenie.", False)

    def open_secure_editor():
        dlg = tk.Toplevel(root)
        dlg.title("Edytuj secure.txt")
        dlg.transient(root)
        dlg.grab_set()

        pretty = json.dumps(connections_state["store"], ensure_ascii=False, indent=2)
        text_widget = tk.Text(dlg, width=80, height=16, wrap="word")
        text_widget.insert("1.0", pretty)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))

        def save_and_close(*_):
            new_content = text_widget.get("1.0", tk.END).strip()
            try:
                parsed = json.loads(new_content) if new_content else {}
                connections_state["store"] = _normalize_connections(parsed)
                persist_connections()
                refresh_connection_combobox()
                dlg.destroy()
                apply_selected_connection(show_success=False)
                messagebox.showinfo(
                    "Zapisano", "Zaktualizowano zawartość pliku secure.txt."
                )
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror(
                    "Błąd zapisu",
                    f"Nie udało się zapisać pliku secure.txt.\n\nSzczegóły techniczne:\n{exc}",
                )

        def cancel(*_):
            dlg.destroy()

        btn_frame = tk.Frame(dlg)
        btn_frame.pack(pady=(0, 10))

        tk.Button(btn_frame, text="Zapisz", width=12, command=save_and_close).pack(
            side="left", padx=(0, 5)
        )
        tk.Button(btn_frame, text="Anuluj", width=12, command=cancel).pack(
            side="left"
        )

        dlg.bind("<Return>", save_and_close)
        dlg.bind("<Escape>", cancel)

        _center_window(dlg, root)

    def choose_sql_file():
        path = filedialog.askopenfilename(
            title="Wybierz plik SQL",
            filetypes=[("SQL files", "*.sql"), ("All files", "*.*")],
        )
        if path:
            _set_sql_path(path)

    def update_template_controls_state():
        enabled = use_template_var.get()

        choose_btn = template_state.get("choose_button")
        if choose_btn is not None:
            choose_btn_state = tk.NORMAL if enabled else tk.DISABLED
            choose_btn.config(state=choose_btn_state)

        sheet_combo = template_state.get("sheet_combobox")
        if sheet_combo is not None:
            sheet_state = "readonly" if enabled else "disabled"
            sheet_combo.config(state=sheet_state)

        start_cell_entry = template_state.get("start_cell_entry")
        if start_cell_entry is not None:
            start_cell_state = tk.NORMAL if enabled else tk.DISABLED
            start_cell_entry.config(state=start_cell_state)

        include_header_check = template_state.get("include_header_check")
        if include_header_check is not None:
            include_header_state = tk.NORMAL if enabled else tk.DISABLED
            include_header_check.config(state=include_header_state)

    def update_csv_profile_controls_state():
        enabled = format_var.get() == "csv"

        combo = csv_profile_state.get("combobox")
        if combo is not None:
            combo_state = "readonly" if enabled else "disabled"
            combo.config(state=combo_state)

        manage_btn = csv_profile_state.get("manage_button")
        if manage_btn is not None:
            manage_state = tk.NORMAL if enabled else tk.DISABLED
            manage_btn.config(state=manage_state)

    def on_toggle_template():
        # Template jest sensowny tylko dla XLSX; jeśli zaznaczono template przy CSV, przełącz na XLSX.
        if use_template_var.get() and format_var.get() != "xlsx":
            format_var.set("xlsx")

        update_template_controls_state()
        update_csv_profile_controls_state()

    def on_format_change(*_):
        """Keep template option consistent with selected output format."""
        if format_var.get() == "csv":
            use_template_var.set(False)

        update_template_controls_state()
        update_csv_profile_controls_state()

    def choose_template_file():
        path = filedialog.askopenfilename(
            title="Wybierz plik template XLSX",
            filetypes=[("Pliki Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return

        template_path_var.set(path)
        template_label_var.set(shorten_path(path))
        use_template_var.set(True)
        on_toggle_template()

        wb = None
        try:
            wb = load_workbook(path, read_only=True)
            sheetnames = wb.sheetnames
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(
                "Błąd template",
                "Nie można odczytać arkuszy z pliku template.\n\n"
                f"Szczegóły techniczne:\n{e}",
            )
            sheetnames = []
        finally:
            try:
                if wb is not None:
                    wb.close()
            except Exception:
                pass

        combo = template_state.get("sheet_combobox")
        if combo is not None:
            combo["values"] = sheetnames

        if sheetnames:
            sheet_name_var.set(sheetnames[0])
        else:
            sheet_name_var.set("")

    def refresh_csv_profile_controls(selected_name=None):
        config = csv_profile_state.get("config", {"profiles": []})
        profiles = config.get("profiles", [])
        names = [p["name"] for p in profiles]

        combo = csv_profile_state.get("combobox")
        if combo is not None:
            combo["values"] = names

        chosen = selected_name or selected_csv_profile_var.get()
        if chosen and chosen in names:
            selected_csv_profile_var.set(chosen)
        elif names:
            selected_csv_profile_var.set(config.get("default_profile") or names[0])
        else:
            selected_csv_profile_var.set("")

        default_name = config.get("default_profile")
        if default_name:
            default_csv_label_var.set(f"Domyślny profil CSV: {default_name}")
        else:
            default_csv_label_var.set("")

    def open_queries_manager():
        dlg = tk.Toplevel(root)
        dlg.title("Edycja queries.txt")
        dlg.transient(root)
        dlg.grab_set()
        dlg.resizable(True, True)

        paths = load_query_paths()
        query_paths_state["paths"] = list(paths)

        dlg.columnconfigure(0, weight=1)
        dlg.rowconfigure(0, weight=1)
        dlg.rowconfigure(1, weight=0)

        list_frame = tk.Frame(dlg)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 5))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        list_var = tk.StringVar(value=[])
        listbox = tk.Listbox(
            list_frame,
            listvariable=list_var,
            width=80,
            height=10,
            activestyle="dotbox",
            selectmode=tk.EXTENDED,
        )
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.config(yscrollcommand=scrollbar.set)
        listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        def refresh_list():
            list_var.set(list(paths))
            update_delete_state()

        def update_delete_state(*_):  # noqa: ANN001
            has_selection = bool(listbox.curselection())
            delete_state = tk.NORMAL if has_selection else tk.DISABLED
            delete_btn.config(state=delete_state)

        def add_from_dialog():
            path = filedialog.askopenfilename(
                title="Dodaj plik SQL",
                filetypes=[("SQL files", "*.sql"), ("All files", "*.*")],
            )
            if not path:
                return

            if path in paths:
                messagebox.showinfo("Informacja", "Ta ścieżka jest już na liście.")
                return

            paths.append(path)
            refresh_list()
            new_idx = len(paths) - 1
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(new_idx)
            listbox.activate(new_idx)
            listbox.see(new_idx)

        def edit_selected(event=None):  # noqa: ANN001
            sel = listbox.curselection()
            if not sel:
                return

            idx = sel[0]
            current_path = paths[idx]
            new_path = simpledialog.askstring(
                "Edytuj ścieżkę zapytania",
                "Edytuj ścieżkę zapytania:",
                initialvalue=current_path,
                parent=dlg,
            )
            if new_path is None:
                return

            new_path = new_path.strip()
            if not new_path:
                return

            if new_path in paths and new_path != current_path:
                messagebox.showinfo("Informacja", "Ta ścieżka jest już na liście.")
                return

            paths[idx] = new_path
            refresh_list()
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(idx)
            listbox.activate(idx)
            listbox.see(idx)

        def delete_selected(event=None):  # noqa: ANN001
            sel = listbox.curselection()
            if not sel:
                messagebox.showinfo("Informacja", "Zaznacz wpis do usunięcia.")
                return "break" if event is not None else None

            for idx in reversed(sel):
                paths.pop(idx)

            refresh_list()

            if paths:
                next_idx = min(sel[0], len(paths) - 1)
                listbox.selection_set(next_idx)
                listbox.activate(next_idx)
                listbox.see(next_idx)

            return "break" if event is not None else None

        def save_and_close(event=None):  # noqa: ANN001
            try:
                save_query_paths(paths)
            except OSError as exc:
                messagebox.showerror(
                    "Błąd zapisu",
                    "Nie można zapisać queries.txt.\n\n" f"Szczegóły techniczne:\n{exc}",
                )
                return

            query_paths_state["paths"] = list(paths)
            dlg.destroy()

        def cancel_dialog(event=None):  # noqa: ANN001
            dlg.destroy()

        button_frame = tk.Frame(dlg)
        button_frame.grid(row=1, column=0, pady=(0, 10), padx=10, sticky="e")

        add_btn = tk.Button(button_frame, text="Dodaj plik...", command=add_from_dialog, width=15)
        delete_btn = tk.Button(
            button_frame,
            text="Usuń zaznaczone",
            command=delete_selected,
            width=18,
        )
        save_btn = tk.Button(button_frame, text="Zapisz", command=save_and_close, width=12)
        cancel_btn = tk.Button(button_frame, text="Anuluj", command=cancel_dialog, width=12)

        add_btn.pack(side="left", padx=(0, 5))
        delete_btn.pack(side="left", padx=(0, 20))
        save_btn.pack(side="left", padx=(0, 5))
        cancel_btn.pack(side="left")

        listbox.bind("<<ListboxSelect>>", update_delete_state)
        listbox.bind("<Double-Button-1>", edit_selected)
        listbox.bind("<Delete>", delete_selected)
        dlg.bind("<Return>", save_and_close)
        dlg.bind("<Escape>", cancel_dialog)
        dlg.protocol("WM_DELETE_WINDOW", cancel_dialog)

        refresh_list()

        if paths:
            listbox.selection_set(0)
            listbox.activate(0)
            listbox.see(0)
        update_delete_state()

        listbox.focus_set()

        _center_window(dlg, root)

    def choose_from_list():
        current_paths = query_paths_state["paths"]
        if not current_paths:
            messagebox.showerror("Error", "Brak raportów w queries.txt")
            return

        dlg = tk.Toplevel(root)
        dlg.title("Wybierz raport z listy")

        dlg.transient(root)
        dlg.resizable(True, True)

        dlg.columnconfigure(0, weight=1)
        dlg.rowconfigure(0, weight=1)

        list_frame = tk.Frame(dlg)
        list_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=10, pady=(10, 5))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        listbox = tk.Listbox(list_frame, width=80, activestyle="dotbox")
        scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.config(yscrollcommand=scrollbar.set)
        listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        for p in current_paths:
            listbox.insert(tk.END, os.path.basename(p))

        path_label_var = tk.StringVar(value="")
        path_label = tk.Label(
            dlg,
            textvariable=path_label_var,
            anchor="w",
            justify="left",
            wraplength=600,
        )
        path_label.grid(row=1, column=0, columnspan=2, sticky="we", padx=10, pady=(0, 10))

        button_frame = tk.Frame(dlg)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(0, 10), padx=10, sticky="e")

        def update_path_label(*_):
            sel = listbox.curselection()
            if sel:
                path_label_var.set(current_paths[sel[0]])
            else:
                path_label_var.set("")

        def on_ok(*_):
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("Uwaga", "Nie wybrano żadnego raportu.")
                return
            idx = sel[0]
            _set_sql_path(current_paths[idx])
            dlg.destroy()

        def on_cancel(*_):
            dlg.destroy()

        ok_btn = tk.Button(button_frame, text="OK", width=12, command=on_ok)
        cancel_btn = tk.Button(button_frame, text="Anuluj", width=12, command=on_cancel)
        ok_btn.pack(side="left", padx=(0, 5))
        cancel_btn.pack(side="left")

        listbox.bind("<<ListboxSelect>>", update_path_label)
        listbox.bind("<Double-Button-1>", on_ok)
        dlg.bind("<Return>", on_ok)
        dlg.bind("<Escape>", on_cancel)

        if current_paths:
            listbox.selection_set(0)
            listbox.activate(0)
            listbox.see(0)
            update_path_label()

        listbox.focus_set()

        _center_window(dlg, root)

        dlg.grab_set()
        root.wait_window(dlg)

    def _build_export_params():
        path = selected_sql_path_full.get()
        if not path:
            messagebox.showerror("Error", "Nie wybrano pliku SQL.")
            return None
        if not os.path.isfile(path):
            messagebox.showerror("Error", "Wybrany plik SQL nie istnieje.")
            return None

        engine = engine_holder.get("engine")
        current_connection = get_connection_by_name(selected_connection_var.get())
        if engine is None or current_connection is None:
            messagebox.showerror(
                "Brak połączenia",
                "Utwórz połączenie z bazą danych przed uruchomieniem raportu.",
            )
            return None

        with open(path, "rb") as f:
            content = f.read()
        sql_query = remove_bom(content).strip()
        if current_connection.get("type") == "mssql_odbc" and sql_query:
            sql_query = (
                "SET ARITHABORT ON;\nSET NOCOUNT ON;\nSET ANSI_WARNINGS OFF;\n"
                + sql_query
            )

        output_format = format_var.get()
        use_template = use_template_var.get()
        base_name = os.path.basename(path)

        csv_profile = None
        if output_format == "csv":
            csv_config = csv_profile_state["config"]
            profile_name = selected_csv_profile_var.get() or csv_config.get("default_profile")
            csv_profile = (
                get_csv_profile(csv_config, profile_name)
                or get_csv_profile(csv_config, csv_config.get("default_profile"))
                or csv_config.get("profiles", [DEFAULT_CSV_PROFILE])[0]
            )

        if use_template:
            if output_format != "xlsx":
                messagebox.showerror(
                    "Błąd",
                    "Template można użyć tylko dla formatu XLSX.",
                )
                return None
            if not template_path_var.get():
                messagebox.showerror("Błąd", "Nie wybrano pliku template.")
                return None
            if not sheet_name_var.get():
                messagebox.showerror("Błąd", "Nie wybrano arkusza z pliku template.")
                return None

            output_file_name = os.path.splitext(base_name)[0] + ".xlsx"
            output_file_path = os.path.join(output_directory, output_file_name)

            return {
                "engine": engine,
                "sql_query": sql_query,
                "output_format": output_format,
                "output_file_path": output_file_path,
                "csv_profile": csv_profile,
                "use_template": True,
                "template": {
                    "template_path": template_path_var.get(),
                    "sheet_name": sheet_name_var.get(),
                    "start_cell": (start_cell_var.get() or "A2").strip(),
                    "include_header": include_header_var.get(),
                },
            }

        output_file_name = os.path.splitext(base_name)[0] + (
            ".xlsx" if output_format == "xlsx" else ".csv"
        )
        output_file_path = os.path.join(output_directory, output_file_name)

        return {
            "engine": engine,
            "sql_query": sql_query,
            "output_format": output_format,
            "output_file_path": output_file_path,
            "csv_profile": csv_profile,
            "use_template": False,
        }

    def run_export_gui():
        sql_query = ""
        params = None

        try:
            params = _build_export_params()
            if not params:
                return

            result_info_var.set("Trwa wykonywanie zapytania i eksport. Proszę czekać...")
            btn_start.config(state=tk.DISABLED)
            root.update_idletasks()

            sql_query = params["sql_query"]
            output_format = params["output_format"]
            csv_profile = params.get("csv_profile")

            if params.get("use_template"):
                template = params["template"]
                sql_dur, export_dur, total_dur, rows_count = run_export_to_template(
                    params["engine"],
                    sql_query,
                    template_path=template["template_path"],
                    output_file_path=params["output_file_path"],
                    sheet_name=template["sheet_name"],
                    start_cell=template["start_cell"],
                    include_header=template["include_header"],
                )
            else:
                sql_dur, export_dur, total_dur, rows_count = run_export(
                    params["engine"],
                    sql_query,
                    params["output_file_path"],
                    output_format,
                    csv_profile=csv_profile,
                )

            last_output_path["path"] = params["output_file_path"]

            if rows_count > 0:
                msg = (
                    f"Zapisano: {params['output_file_path']}\n"
                    f"Wiersze: {rows_count}\n"
                    f"Czas SQL: {sql_dur:.2f} s\n"
                    f"Czas eksportu: {export_dur:.2f} s\n"
                    f"Czas łączny: {total_dur:.2f} s"
                )
                if output_format == "csv" and csv_profile:
                    msg += f"\nProfil CSV: {csv_profile.get('name', '')}"
            else:
                msg = f"Zapytanie nie zwróciło wierszy.\nCzas SQL: {sql_dur:.2f} s"
                if output_format == "csv" and csv_profile:
                    msg += f"\nProfil CSV: {csv_profile.get('name', '')}"

            result_info_var.set(msg)
            messagebox.showinfo("Gotowe", msg)
            btn_open_file.config(state=tk.NORMAL)
            btn_open_folder.config(state=tk.NORMAL)
            update_error_display("")

        except Exception as exc:  # noqa: BLE001
            ui_msg = format_error_for_ui(exc, sql_query)
            result_info_var.set("Błąd eksportu. Pełne szczegóły w logu.")
            update_error_display(ui_msg)
            show_error_popup(ui_msg)
        finally:
            btn_start.config(state=tk.NORMAL)

    def _open_path(target):
        if not target or not os.path.exists(target):
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(target)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", target], check=False)
            else:
                subprocess.run(["xdg-open", target], check=False)
        except Exception as err:  # noqa: BLE001
            messagebox.showerror("Error", str(err))

    def open_file():
        path = last_output_path.get("path")
        _open_path(path)

    def open_folder():
        path = last_output_path.get("path")
        if path and os.path.isfile(path):
            folder = os.path.dirname(path)
            _open_path(folder)

    connection_frame = tk.LabelFrame(
        root, text="Połączenie z bazą danych", padx=10, pady=10
    )
    connection_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

    status_label = tk.Label(
        connection_frame,
        textvariable=connection_status_var,
        justify="left",
        anchor="w",
    )
    status_label.grid(row=0, column=0, sticky="we")
    connection_frame.columnconfigure(0, weight=1)

    def _update_status_wrap(event=None):
        width = connection_frame.winfo_width() or status_label.winfo_width() or 0
        wrap = max(width - 20, 200)
        status_label.config(wraplength=wrap)

    connection_frame.bind("<Configure>", _update_status_wrap)
    status_label.bind("<Configure>", _update_status_wrap)

    connection_controls = tk.Frame(connection_frame)
    connection_controls.grid(row=1, column=0, sticky="we", pady=(5, 0))
    connection_controls.columnconfigure(1, weight=1)

    tk.Label(connection_controls, text="Połączenie:").grid(row=0, column=0, sticky="w")
    connection_combo = ttk.Combobox(
        connection_controls,
        textvariable=selected_connection_var,
        state="readonly",
        width=50,
    )
    connection_combo.grid(row=0, column=1, sticky="we", padx=(5, 0))
    connections_state["combobox"] = connection_combo
    connection_combo.bind("<<ComboboxSelected>>", on_connection_change)

    tk.Button(
        connection_controls,
        text="Dodaj/edytuj połączenie",
        command=lambda: open_connection_dialog_gui(
            root,
            connections_state,
            selected_connection_var,
            get_connection_by_name,
            set_selected_connection,
            persist_connections,
            refresh_connection_combobox,
            apply_selected_connection,
            handle_db_driver_error,
            create_engine_from_entry,
        ),
    ).grid(row=0, column=2, padx=(10, 0), sticky="e")
    tk.Button(
        connection_controls,
        text="Testuj połączenie",
        command=test_connection_only,
    ).grid(row=0, column=3, padx=(10, 0), sticky="e")
    tk.Button(
        connection_controls,
        text="Usuń połączenie",
        command=delete_selected_connection,
    ).grid(row=0, column=4, padx=(10, 0), sticky="e")

    secure_edit_btn = tk.Button(
        connection_controls,
        text="Edytuj secure.txt",
        command=open_secure_editor,
    )
    secure_edit_btn.grid(row=0, column=5, padx=(10, 0), sticky="e")
    secure_edit_state["button"] = secure_edit_btn

    source_frame = tk.LabelFrame(root, text="Źródło zapytania SQL", padx=10, pady=10)
    source_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

    format_frame = tk.LabelFrame(root, text="Format wyjściowy", padx=10, pady=10)
    format_frame.pack(fill=tk.X, padx=10, pady=5)

    template_frame = tk.LabelFrame(root, text="Opcje template XLSX (GUI)", padx=10, pady=10)
    template_frame.pack(fill=tk.X, padx=10, pady=5)

    result_frame = tk.LabelFrame(root, text="Wynik i akcje", padx=10, pady=10)
    result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))

    source_frame.columnconfigure(1, weight=1)
    source_frame.columnconfigure(2, weight=0)
    template_frame.columnconfigure(1, weight=1)
    result_frame.columnconfigure(1, weight=1)
    result_frame.rowconfigure(3, weight=1)

    tk.Label(source_frame, text="Wybrany plik SQL:").grid(row=0, column=0, sticky="nw")
    tk.Label(source_frame, textvariable=sql_label_var, wraplength=600, justify="left").grid(
        row=0, column=1, columnspan=3, sticky="we"
    )

    tk.Button(source_frame, text="Wybierz plik SQL", command=choose_sql_file).grid(
        row=1, column=0, pady=5, sticky="w"
    )
    tk.Button(source_frame, text="Wybierz z listy raportów", command=choose_from_list).grid(
        row=1, column=1, pady=5, sticky="w"
    )
    tk.Button(source_frame, text="Edytuj queries.txt", command=open_queries_manager).grid(
        row=1, column=2, pady=5, sticky="w"
    )

    tk.Radiobutton(
        format_frame, text="XLSX", variable=format_var, value="xlsx", command=on_format_change
    ).grid(row=0, column=0, sticky="w")
    tk.Radiobutton(
        format_frame, text="CSV", variable=format_var, value="csv", command=on_format_change
    ).grid(row=0, column=1, sticky="w")

    on_format_change()

    tk.Label(format_frame, text="Profil CSV:").grid(
        row=1, column=0, sticky="w", pady=(5, 0)
    )
    csv_profile_combo = ttk.Combobox(
        format_frame,
        textvariable=selected_csv_profile_var,
        state="readonly",
        width=25,
    )
    csv_profile_combo.grid(row=1, column=1, sticky="w", pady=(5, 0))
    csv_profile_state["combobox"] = csv_profile_combo

    csv_profile_manage_btn = tk.Button(
        format_frame,
        text="Zarządzaj profilami CSV",
        command=lambda: open_csv_profiles_manager_gui(
            root,
            csv_profile_state,
            selected_csv_profile_var,
            refresh_csv_profile_controls,
        ),
    )
    csv_profile_manage_btn.grid(row=1, column=2, padx=(10, 0), pady=(5, 0), sticky="w")
    csv_profile_state["manage_button"] = csv_profile_manage_btn

    tk.Label(format_frame, textvariable=default_csv_label_var, justify="left", wraplength=600).grid(
        row=2, column=0, columnspan=3, sticky="w", pady=(5, 0)
    )

    refresh_csv_profile_controls(csv_profile_state["config"].get("default_profile"))

    tk.Checkbutton(
        template_frame,
        text="Użyj pliku template (tylko dla XLSX, tylko w GUI)",
        variable=use_template_var,
        command=on_toggle_template,
    ).grid(row=0, column=0, columnspan=2, sticky="w")

    tk.Label(template_frame, text="Plik template:").grid(
        row=1, column=0, sticky="w", pady=(5, 0)
    )
    choose_template_btn = tk.Button(
        template_frame, text="Wybierz template", command=choose_template_file
    )
    choose_template_btn.grid(row=1, column=1, sticky="w", pady=(5, 0))
    template_state["choose_button"] = choose_template_btn
    tk.Label(
        template_frame,
        textvariable=template_label_var,
        wraplength=600,
        justify="left",
    ).grid(row=2, column=0, columnspan=2, sticky="we")

    tk.Label(template_frame, text="Arkusz:").grid(
        row=3, column=0, sticky="w", pady=(5, 0)
    )
    sheet_combobox = ttk.Combobox(
        template_frame,
        textvariable=sheet_name_var,
        state="readonly",
        width=30,
    )
    sheet_combobox.grid(row=3, column=1, sticky="w", pady=(5, 0))
    template_state["sheet_combobox"] = sheet_combobox

    tk.Label(template_frame, text="Startowa komórka:").grid(
        row=4, column=0, sticky="w", pady=(5, 0)
    )
    start_cell_entry = tk.Entry(template_frame, textvariable=start_cell_var, width=10)
    start_cell_entry.grid(row=4, column=1, sticky="w", pady=(5, 0))
    template_state["start_cell_entry"] = start_cell_entry

    include_header_check = tk.Checkbutton(
        template_frame,
        text="Zapisz nagłówki (nazwy kolumn) w arkuszu",
        variable=include_header_var,
    )
    include_header_check.grid(row=5, column=0, columnspan=2, sticky="w", pady=(5, 0))
    template_state["include_header_check"] = include_header_check

    update_template_controls_state()
    update_csv_profile_controls_state()

    btn_start = tk.Button(result_frame, text="Start", command=run_export_gui)
    btn_start.grid(row=0, column=0, pady=(0, 10), sticky="w")
    start_button_holder["widget"] = btn_start

    refresh_connection_combobox()
    refresh_secure_edit_button()
    if selected_connection_var.get():
        apply_selected_connection(show_success=False)
    else:
        set_connection_status("Brak połączenia. Utwórz nowe połączenie.", False)

    tk.Label(result_frame, text="Informacje o eksporcie:").grid(row=1, column=0, sticky="nw")
    tk.Label(result_frame, textvariable=result_info_var, justify="left", wraplength=600).grid(
        row=1, column=1, columnspan=3, sticky="w"
    )

    btn_open_file = tk.Button(result_frame, text="Otwórz plik", command=open_file)
    btn_open_file.grid(row=2, column=0, pady=5, sticky="w")
    btn_open_folder = tk.Button(result_frame, text="Otwórz katalog", command=open_folder)
    btn_open_folder.grid(row=2, column=1, pady=5, sticky="w")

    tk.Label(result_frame, text="Błędy (skrót):").grid(
        row=3, column=0, sticky="nw", pady=(10, 0)
    )
    error_frame = tk.Frame(result_frame)
    error_frame.grid(row=3, column=1, columnspan=3, sticky="nsew", pady=(10, 0))

    error_text = tk.Text(
        error_frame,
        width=120,
        height=6,
        wrap="none",
        state="disabled",
        font=("Consolas", 9),
    )
    error_y_scroll = tk.Scrollbar(error_frame, orient="vertical", command=error_text.yview)
    error_x_scroll = tk.Scrollbar(error_frame, orient="horizontal", command=error_text.xview)
    error_text.configure(yscrollcommand=error_y_scroll.set, xscrollcommand=error_x_scroll.set)

    error_y_scroll.pack(side="right", fill="y")
    error_x_scroll.pack(side="bottom", fill="x")
    error_text.pack(side="left", fill="both", expand=True)

    error_display["widget"] = error_text

    btn_open_file.config(state=tk.DISABLED)
    btn_open_folder.config(state=tk.DISABLED)

    _center_window(root)

    root.mainloop()


if __name__ == "__main__":
    connection_store = load_connections()
    selected_name = connection_store.get("last_selected")
    selected_connection = None
    for conn in connection_store.get("connections", []):
        if conn.get("name") == selected_name:
            selected_connection = conn
            break
    if selected_connection is None and connection_store.get("connections"):
        selected_connection = connection_store["connections"][0]

    output_directory = r"generated_reports"
    ensure_directories([output_directory, "templates", "queries"])

    if len(sys.argv) > 1 and sys.argv[1] == "-c":
        if not selected_connection:
            print(
                "Brak zapisanych połączeń. Utwórz połączenie w trybie GUI, aby uruchomić konsolę."
            )
            sys.exit(1)

        engine_kwargs = {}
        if selected_connection.get("type") == "mssql_odbc":
            engine_kwargs["isolation_level"] = "AUTOCOMMIT"
        try:
            engine = create_engine(selected_connection.get("url"), **engine_kwargs)
        except Exception as exc:  # noqa: BLE001
            handled = handle_db_driver_error(
                exc,
                selected_connection.get("type"),
                selected_connection.get("name"),
                show_message=lambda *args: print(args[-1] if args else exc),
            )
            if not handled:
                LOGGER.exception(
                    "Failed to initialize console engine for %s (%s)",
                    selected_connection.get("name"),
                    selected_connection.get("type"),
                    exc_info=exc,
                )
                print("Nie udało się utworzyć połączenia. Pełne szczegóły w logu.")
            sys.exit(1)

        run_console(engine, output_directory, selected_connection)
    else:
        run_gui(connection_store, output_directory)
