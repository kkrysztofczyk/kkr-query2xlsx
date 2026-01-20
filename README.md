# kkr-query2xlsx

Run SQL queries from files against multiple databases and export results to **Excel (XLSX)** or **CSV**.  
Includes **GUI and CLI modes**, retry handling, CSV profiles, XLSX templates, and a ready-to-use **SQLite demo**.

---

## Screenshot

> Recommended file location in the repo: `docs/gui.png`

![KKr SQL to XLSX/CSV GUI](docs/gui.png)

---

## What this tool is for

- Running SQL queries stored in `.sql` files
- Exporting query results to XLSX or CSV
- Reusing the same queries across different databases
- Non-developers running reports without touching SQL editors
- Lightweight, local alternative to BI tools for ad-hoc reporting

---

## Features

- GUI (Tkinter) and CLI modes
- Supports **SQLite**, **SQL Server (ODBC)**, **PostgreSQL**
- File-based SQL queries
- Excel (XLSX) and CSV export
- CSV profiles (delimiter, encoding, decimal separator, date format)
- XLSX template support (paste results into existing sheets) — **GUI only**
- Retry logic for deadlocks / serialization errors
- Rotating logs
- Demo SQLite database + example queries included

---

## Requirements

- Python **3.9+** recommended
- Dependencies are listed in `requirements.txt`
- Notes:
  - **SQL Server (ODBC)** requires Microsoft ODBC Driver (e.g. *ODBC Driver 17/18 for SQL Server*) and `pyodbc`
  - **PostgreSQL** requires a PostgreSQL driver (e.g. `psycopg2-binary`)
  - SQLite works out-of-the-box (file-based)

---

## Repository structure

```text
.
├─ main.pyw
├─ secure.sample.json
├─ queries.sample.txt
├─ examples/
│  ├─ db/
│  │  └─ demo.sqlite
│  └─ queries/
│     ├─ 01_simple_select.sql
│     ├─ 02_join.sql
│     └─ 03_aggregation.sql
├─ docs/
│  └─ gui.png
├─ templates/
├─ generated_reports/   (created at runtime)
├─ logs/                (created at runtime)
├─ requirements.txt
├─ LICENSE
└─ README.md
```

---

## Quickstart (Demo – no setup required)

### 1. Install dependencies

```bash
python -m pip install -r requirements.txt
```

### 2. Prepare local config

Copy the sample connection file:

**macOS / Linux**
```bash
cp secure.sample.json secure.txt
```

**Windows (PowerShell)**
```powershell
Copy-Item secure.sample.json secure.txt
```

**Windows (CMD)**
```bat
copy secure.sample.json secure.txt
```

(Optional) Copy example query list:

**macOS / Linux**
```bash
cp queries.sample.txt queries.txt
```

**Windows (PowerShell)**
```powershell
Copy-Item queries.sample.txt queries.txt
```

**Windows (CMD)**
```bat
copy queries.sample.txt queries.txt
```

> `secure.txt` is **gitignored** and must never be committed.

---

### 3. Run the app

GUI mode (default):

```bash
python main.pyw
```

Console mode:

```bash
python main.pyw -c
```

---

### 4. Run a demo query (GUI)

1. Select connection: **Demo SQLite**
2. Choose a query from `examples/queries`
3. Select output format: XLSX or CSV
4. Click **Start**

Results are saved to:

```text
generated_reports/
```

---

## Usage

### GUI usage

- Configure a connection (or use **Demo SQLite**)
- Pick a `.sql` file (or use `queries.txt` list)
- Choose export format (XLSX/CSV)
- Click **Start**
- Use “Open file” / “Open folder” buttons after export

### CLI usage (`-c`)

```bash
python main.pyw -c
```

Notes:
- CLI mode requires at least one saved connection in `secure.txt`
- You will be prompted to choose a query file and output format
- Output is written to `generated_reports/`

---

## Configuration

### Connections (`secure.txt`)

Connections are stored locally in `secure.txt` (JSON).

Supported types:
- SQLite (file-based)
- SQL Server (ODBC)
- PostgreSQL

Use `secure.sample.json` as a template and rename it to `secure.txt`.

### Query list (`queries.txt`)

Optional file listing paths to `.sql` files, one per line.

Example:

```text
examples/queries/01_simple_select.sql
examples/queries/02_join.sql
```

You can also pick any `.sql` file manually from the GUI.

### CSV profiles

CSV export can be customized via profiles:
- encoding (UTF-8, windows-1250, etc.)
- delimiter and decimal separator
- quoting strategy
- date formatting

Profiles can be managed directly from the GUI.

⚠️ Note: `delimiter_replacement` intentionally **modifies string values** to avoid escaping issues. Use with care.

### XLSX templates (GUI only)

You can export query results directly into an existing Excel file:
- select template `.xlsx`
- choose target sheet
- define start cell
- optionally include column headers

The template file is copied before writing.

---

## Output & logs

- Exports are written to: `generated_reports/`
- Logs are written to: `logs/kkr_query2sheet.log`
  - Rotating log files (max ~1 MB, 3 backups)
  - Logs may include SQL text and error details

---

## Troubleshooting

- **SQL Server: ODBC driver missing / pyodbc missing**
  - Install Microsoft ODBC Driver 17/18 and ensure `pyodbc` is installed.
- **PostgreSQL: psycopg2 missing**
  - Install `psycopg2-binary` (or another PostgreSQL driver).
- **PermissionError when exporting**
  - The output file may be open in Excel. Close it and retry.
- **Nothing happens / no rows returned**
  - The query executed but returned 0 rows — check your SQL and filters.
- **Where are details of an error?**
  - See `logs/kkr_query2sheet.log`

---

## Security notes

- This tool executes **arbitrary SQL** — run only queries you trust.
- `secure.txt` may contain database credentials — **never commit it**.
- Logs can contain SQL fragments and connection/driver error details — treat logs as sensitive.

---

## License

MIT License (see `LICENSE`)

---

## Disclaimer

This tool is provided as-is.  
Use at your own risk, especially when connecting to production databases.  
This project uses third-party libraries licensed under their respective licenses (MIT, BSD, LGPL).
