"""
Local dashboard to run your ERP sync scripts with a button click instead of
the command line, and download the resulting database afterwards.

SETUP (one time):
    pip install flask

HOW TO RUN:
    1. Save this file as app.py in the SAME folder as all your scripts
       (Process_Master.py, Production_Entry.py, Sales_Register.py, etc.)
       and their .db files - e.g. D:\\Unitech Dashboard Database\\
    2. Open a terminal in that folder and run:
           python app.py
    3. Open your browser to:
           http://localhost:5000

Click a card's "Run" button to execute that script. Output/log shows on the
page as it happens. Once it finishes, a "Download DB" button appears so you
can grab the latest database file.
"""

import subprocess
import sys
import os
import threading
import queue
import uuid
import sqlite3
import io
import openpyxl.styles
import openpyxl.utils
import calendar
import hashlib
import time
from datetime import datetime, timedelta

import openpyxl

from flask import Flask, render_template, jsonify, request, send_file, Response

app = Flask(__name__)


def format_number_with_commas(value):
    """Formats a number with comma thousands-separators for display in
    editable fields (e.g. 784710 -> '784,710', 1234.5 -> '1,234.50').
    Blank/None stays blank rather than showing '0' or an error."""
    if value is None or value == "":
        return ""
    try:
        value = float(value)
    except (TypeError, ValueError):
        return str(value)
    if value == int(value):
        return f"{int(value):,}"
    return f"{value:,.2f}"


app.jinja_env.filters["commafy"] = format_number_with_commas

# This file (ProjectDashboard.py) lives in the DASHBOARD folder.
# Scripts live in a sibling SCRIPTS folder, databases in a sibling DATABASE folder:
#   D:\Unitech Dashboard Database\SCRIPTS\
#   D:\Unitech Dashboard Database\DATABASE\
#   D:\Unitech Dashboard Database\DASHBOARD\ProjectDashboard.py   <- this file
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPTS_DIR = os.path.join(BASE_DIR, "..", "SCRIPTS")
DATABASE_DIR = os.path.join(BASE_DIR, "..", "DATABASE")
ANALYTICS_DB = os.path.join(DATABASE_DIR, "Analytics.db")

# Sidebar navigation for the analytics pages, in display order.
# "ready" pages have a real screen built; others show a placeholder.
NAV_ITEMS = [
    {"key": "overview", "label": "Overview", "ready": True},
    {"key": "part_margin", "label": "Part-wise margin", "ready": True},
    {"key": "tree", "label": "BOM tree", "ready": True},
    {"key": "part_cost", "label": "Part cost calc", "ready": True},
    {"key": "month_posting", "label": "Month posting", "ready": True},
    {"key": "weighted_price", "label": "Weighted price", "ready": True},
    {"key": "cycle_time", "label": "Cycle time", "ready": True},
    {"key": "other_cost", "label": "Other cost", "ready": True},
    {"key": "data_sync", "label": "Data sync", "ready": True},
]

# ---- Script registry: edit filenames here if you rename anything ----
SCRIPTS = {
    "process_master": {
        "label": "Process Master",
        "description": "Download missing stores, fill in their process, upload back",
        "script": "Process_Master.py",
        "db": "Production_Entry.db",
        "needs_month": False,
        "has_store_download": True,
    },
    "mould_master": {
        "label": "Mould Master",
        "description": "Download Mould Master (Runner Wt lookup)",
        "script": "Mould_Master.py",
        "needs_login": True,   # uses the ERP login typed when the dashboard started
        "db": "Production_Entry.db",
        "needs_month": False,
    },
    "item_master": {
        "label": "Item Master",
        "description": "Download Item Master (monthly Item Type snapshot)",
        "script": "Item_Master.py",
        "needs_login": True,   # uses the ERP login typed when the dashboard started
        "db": "Production_Entry.db",
        "needs_month": False,
    },
    "vendor_price_master": {
        "label": "Vendor Price Master",
        "description": "Monthly snapshot of the ERP's mapped vendor purchase prices (supplier item mapping)",
        "script": "Vendor_Price_Master.py",
        "needs_login": True,   # uses the ERP login typed when the dashboard started
        "db": "Vendor_Price_Master.db",
        "needs_month": False,
    },
    "production_entry": {
        "label": "Production Entry",
        "description": "Production Entry + Sub Assembly Entry, both companies",
        "script": "Production_Entry.py",
        "needs_login": True,   # uses the ERP login typed when the dashboard started
        "db": "Production_Entry.db",
        "needs_month": False,
    },
    "sales_register": {
        "label": "Sales Register",
        "description": "Incremental Sales Register sync",
        "script": "Sales_Register.py",
        "needs_login": True,   # uses the ERP login typed when the dashboard started
        "db": "Sales_Register.db",
        "needs_month": False,
    },
    "grn_register": {
        "label": "GRN Register",
        "description": "Incremental GRN Register sync",
        "script": "GRN_Register.py",
        "needs_login": True,   # uses the ERP login typed when the dashboard started
        "db": "GRN_Register.db",
        "needs_month": False,
    },
    "issuance_report": {
        "label": "Material Issuance",
        "description": "RMI Material Issue Report sync",
        "script": "Issuance_Report.py",
        "needs_login": True,   # uses the ERP login typed when the dashboard started
        "db": "Material_Issue_Report.db",
        "needs_month": False,
    },
    "transfer_register": {
        "label": "Material Transfer",
        "description": "RMI Material Transfer Register sync",
        "script": "Transfer_Register.py",
        "needs_login": True,   # uses the ERP login typed when the dashboard started
        "db": "Material_Transfer_Register.db",
        "needs_month": False,
    },
    "return_register": {
        "label": "Material Return",
        "description": "RMI Material Return Register sync",
        "script": "Return_Register.py",
        "needs_login": True,   # uses the ERP login typed when the dashboard started
        "db": "Material_Return_Register.db",
        "needs_month": False,
    },
    "opncls_stock": {
        "label": "Opening/Closing Stock",
        "description": "Monthly RMI opening/closing stock (pick a month below)",
        "script": "OpnCls_Stock.py",
        "needs_login": True,   # uses the ERP login typed when the dashboard started
        "db": "Opening_Closing_Stock.db",
        "needs_month": True,
    },
    "weighted_price_master": {
        "label": "Weighted Price Master",
        "description": "Monthly weighted-average purchase price per item, from GRN Register",
        "script": "Weighted_Price_Master.py",
        "db": "Analytics.db",
        "needs_month": False,
    },
    "weighted_selling_price": {
        "label": "Weighted Selling Price",
        "description": "Monthly weighted-average selling price per FG, from Sales Register - powers Part-wise Margin",
        "script": "Weighted_Selling_Price.py",
        "db": "Analytics.db",
        "needs_month": False,
    },
    "bom_register": {
        "label": "BOM Register",
        "description": "Builds BOM trees for the FG codes sold that month (run Sales Register and Item Master first)",
        "script": "BOM_Register.py",
        "needs_login": True,   # uses the ERP login typed when the dashboard started
        "db": "Analytics.db",
        "needs_month": True,
    },
}

# JOBS[job_id] = {
#     "log_lines": [...],      full history so far - lets a reconnecting
#                               browser tab catch up instantly
#     "listeners": [Queue, ...]  one queue per currently-open live viewer,
#                               so multiple tabs/viewers can each tail live
#     "finished": bool,
#     "returncode": int or None,
#     "script_key": str,
#     "lock": threading.Lock(),
# }
JOBS = {}

# ---- ERP login: asked once when the dashboard starts, kept only in memory ----
# Every script started from the dashboard uses it until the dashboard window is
# closed. It is never written to a file, a database or the log.
ERP_LOGIN = {}
ERP_LOGIN_CHECK_URL = "https://unitech.tarka.in/app.php"


def check_erp_login(username, password):
    """True = ERP accepted it, False = ERP rejected it, None = couldn't reach the ERP."""
    try:
        import requests
        session = requests.Session()
        session.get("https://unitech.tarka.in", timeout=20)
        resp = session.post(
            ERP_LOGIN_CHECK_URL,
            params={"a": "validateUsernameAndPassword.demo"},
            data={"usr": username, "password": password, "client_company": "9"},
            headers={"X-Requested-With": "XMLHttpRequest",
                     "Referer": "https://unitech.tarka.in/MainPage.html"},
            timeout=20,
        )
        return resp.json().get("status") == "success"
    except Exception:
        return None


def ask_erp_login_at_startup():
    import getpass
    if not (sys.stdin and sys.stdin.isatty()):
        print("⚠️  No window to type the ERP login into - the Data Sync page will ask for it instead.")
        return

    print()
    print("ERP login - used for every script until you close this window (nothing is saved)")
    for attempt in range(3):
        username = input("  Username: ").strip()
        password = getpass.getpass("  Password (hidden while typing): ")
        if not username or not password:
            print("  Username and password can't be blank - try again.\n")
            continue

        print("  Checking with the ERP...")
        result = check_erp_login(username, password)
        if result is True:
            print("  ✅ Login OK\n")
            ERP_LOGIN.update(username=username, password=password)
            return
        if result is None:
            print("  ⚠️  Couldn't reach the ERP to check it (internet down?) - using it anyway.\n")
            ERP_LOGIN.update(username=username, password=password)
            return
        print("  ❌ ERP rejected that username/password - try again.\n")

    print("  Too many wrong tries - the Data Sync page will ask for the login when you click Run.\n")

# Which job_id is the (most recent) one for each script key - lets the
# browser ask "is anything running for this card?" after a page reload.
ACTIVE_JOB_BY_KEY = {}


def run_script_in_background(job_id, script_path, args, extra_env=None):
    """Runs the script as a subprocess. Every output line is stored in the
    job's permanent history AND pushed to every currently-listening browser
    tab, so a tab that reconnects later (e.g. after navigating away and
    back) can catch up on everything it missed."""
    job = JOBS[job_id]

    def emit(line):
        with job["lock"]:
            job["log_lines"].append(line)
            for listener_q in job["listeners"]:
                listener_q.put(line)

    def finish(returncode):
        with job["lock"]:
            job["returncode"] = returncode
            job["finished"] = True
            invalidate_cost_cache()   # any script could have changed a part's cost
            for listener_q in job["listeners"]:
                listener_q.put("__DONE__")

    try:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"   # avoids crashes on emoji (✅ ❌) in piped output on Windows
        env["PYTHONUTF8"] = "1"
        if extra_env:
            env.update(extra_env)            # ERP login for this one run - only in memory,
                                             # handed straight to the script, never saved
        env["PYTHONUNBUFFERED"] = "1"       # otherwise Python buffers ALL output until the
                                             # buffer fills or the script exits, since it's not
                                             # writing to a real terminal - looks like it's
                                             # "hung" even though it's actually working fine

        process = subprocess.Popen(
            [sys.executable, "-u", script_path] + args,  # -u: belt-and-braces, same fix as the env var above
            cwd=SCRIPTS_DIR,
            stdin=subprocess.DEVNULL,   # a script must never sit waiting for typing in the dashboard's window
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=env,
        )
        for line in process.stdout:
            emit(line.rstrip("\n"))
        process.wait()
        finish(process.returncode)
    except Exception as e:
        emit(f"❌ Failed to run script: {e}")
        finish(1)


def start_job(script_key, script_path, args, extra_env=None):
    """Kicks off a background job and returns its job_id. extra_env (the
    ERP login) goes only to the script's process - it is not kept in JOBS."""
    job_id = str(uuid.uuid4())
    JOBS[job_id] = {
        "log_lines": [],
        "listeners": [],
        "finished": False,
        "returncode": None,
        "script_key": script_key,
        "lock": threading.Lock(),
    }
    ACTIVE_JOB_BY_KEY[script_key] = job_id

    thread = threading.Thread(target=run_script_in_background, args=(job_id, script_path, args, extra_env))
    thread.daemon = True
    thread.start()

    return job_id


def get_available_months():
    """List every distinct month present in weighted_price_master, newest first."""
    if not os.path.exists(ANALYTICS_DB):
        return []
    conn = sqlite3.connect(ANALYTICS_DB)
    cur = conn.cursor()
    try:
        cur.execute("SELECT DISTINCT month FROM weighted_price_master ORDER BY month DESC")
        months = [row[0] for row in cur.fetchall()]
    except sqlite3.OperationalError:
        months = []
    conn.close()
    return months


BASELINE_YEAR_MONTH = (2026, 6)   # matches the overall data baseline (1 Jun 2026)


def get_all_relevant_months():
    """
    Every month from the data baseline through the current month, newest
    first. Used for pages like Other Cost where you need to be able to pick
    ANY month to enter data for - not just months that already have data
    (unlike get_available_months, which only lists months already present
    in weighted_price_master).
    """
    y, m = BASELINE_YEAR_MONTH
    now = datetime.now()
    months = []
    while (y, m) <= (now.year, now.month):
        months.append(f"{y}_{m:02d}")
        m += 1
        if m > 12:
            m = 1
            y += 1
    return list(reversed(months))


def month_key_to_label(month_key):
    """'2026_09' -> 'September 2026'"""
    try:
        year, month = month_key.split("_")
        return datetime(int(year), int(month), 1).strftime("%B %Y")
    except (ValueError, AttributeError):
        return month_key


def render_analytics_page(template_name, page_key, months_override=None, **extra_context):
    months = months_override if months_override is not None else get_available_months()
    selected_month = request.args.get("month", months[0] if months else None)
    return render_template(
        template_name,
        nav_items=NAV_ITEMS,
        active_page=page_key,
        months=[{"key": m, "label": month_key_to_label(m)} for m in months],
        selected_month=selected_month,
        page_label=next(n["label"] for n in NAV_ITEMS if n["key"] == page_key),
        **extra_context,
    )


DOWNLOAD_RANGE_CONFIG = {
    "production_entry": {"db": "Production_Entry.db", "table": "production_register", "date_column": "date"},
    "sales_register": {"db": "Sales_Register.db", "table": "sales_register", "date_column": "invoice_date"},
    "grn_register": {"db": "GRN_Register.db", "table": "grn_register", "date_column": "grn_date"},
    "issuance_report": {"db": "Material_Issue_Report.db", "table": "material_issue_report", "date_column": "issue_date"},
    "transfer_register": {"db": "Material_Transfer_Register.db", "table": "material_transfer_register", "date_column": "transfer_date"},
    "return_register": {"db": "Material_Return_Register.db", "table": "material_return_register", "date_column": "date"},
}


def ddmmyyyy_to_sortable(date_str):
    """'15/09/2026' -> '20260915' - a plain string that sorts/compares correctly."""
    day, month, year = date_str.split("/")
    return f"{year}{month}{day}"


@app.route("/data_sync/download_range/<key>")
def data_sync_download_range(key):
    if key not in DOWNLOAD_RANGE_CONFIG:
        return jsonify({"error": "This script doesn't support date-range downloads"}), 400

    from_date = request.args.get("from_date", "").strip()
    to_date = request.args.get("to_date", "").strip()
    if not from_date or not to_date:
        return jsonify({"error": "Both From and To dates are required"}), 400

    try:
        from_sortable = ddmmyyyy_to_sortable(from_date)
        to_sortable = ddmmyyyy_to_sortable(to_date)
    except ValueError:
        return jsonify({"error": "Dates must be in dd/mm/yyyy format"}), 400

    config = DOWNLOAD_RANGE_CONFIG[key]
    db_path = os.path.join(DATABASE_DIR, config["db"])
    if not os.path.exists(db_path):
        return jsonify({"error": f"{config['db']} not found - run this script at least once first"}), 404

    table = config["table"]
    date_col = config["date_column"]

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table})")
    columns = [row[1] for row in cur.fetchall()]

    cur.execute(f"""
        SELECT {", ".join(columns)} FROM {table}
        WHERE substr({date_col},7,4) || substr({date_col},4,2) || substr({date_col},1,2)
              BETWEEN ? AND ?
    """, (from_sortable, to_sortable))
    rows = cur.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(columns)
    for row in rows:
        ws.append(row)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    safe_from = from_date.replace("/", "-")
    safe_to = to_date.replace("/", "-")

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{key}_{safe_from}_to_{safe_to}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def parse_ddmmyyyy(date_str):
    return datetime.strptime(date_str, "%d/%m/%Y")


def get_months_in_range(from_date, to_date):
    """List of 'YYYY_MM' month keys covered by [from_date, to_date], inclusive."""
    start = parse_ddmmyyyy(from_date)
    end = parse_ddmmyyyy(to_date)
    months = []
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        months.append(f"{y}_{m:02d}")
        m += 1
        if m > 12:
            m = 1
            y += 1
    return months


def get_previous_period_dates(from_date_str, to_date_str):
    """The immediately preceding period of the SAME length, ending right
    before the current period starts - the standard 'vs previous period'
    comparison window used by most BI dashboards."""
    from_dt = parse_ddmmyyyy(from_date_str)
    to_dt = parse_ddmmyyyy(to_date_str)
    duration_days = (to_dt - from_dt).days + 1
    prev_to_dt = from_dt - timedelta(days=1)
    prev_from_dt = prev_to_dt - timedelta(days=duration_days - 1)
    return prev_from_dt.strftime("%d/%m/%Y"), prev_to_dt.strftime("%d/%m/%Y")


def pct_change(current, previous):
    """Relative percent change; None (shown as '—') if there's nothing to
    compare against, rather than a divide-by-zero or a misleadingly huge number."""
    if not previous:
        return None
    return round(((current - previous) / abs(previous)) * 100, 2)


def compute_margin_rows(from_date, to_date, units=None, include_missing=True):
    """
    The core dataset behind Overview: one row per (month, unit, customer,
    fg_code), with Revenue, Cost, Margin computed STRICTLY per month - each
    month's sales are costed using THAT month's BOM Tree data, and nothing
    is blended or averaged across months. A range spanning June-August
    costs June's sales from June's tree, August's from August's tree, then
    the caller sums the monthly results together.

    Revenue is each sale's own actual invoice value (Taxable Value) - never
    a weighted average here, that's what Part-wise Margin is for.
    """
    months = get_months_in_range(from_date, to_date)
    rows = []

    sales_db = os.path.join(DATABASE_DIR, "Sales_Register.db")
    if not os.path.exists(sales_db) and not os.path.exists(OTHER_SALES_DB):
        return rows

    range_start_dt = parse_ddmmyyyy(from_date)
    range_end_dt = parse_ddmmyyyy(to_date)

    for month in months:
        year, mon = month.split("_")
        year, mon = int(year), int(mon)

        month_start = datetime(year, mon, 1)
        last_day = calendar.monthrange(year, mon)[1]
        month_end = datetime(year, mon, last_day)

        # This month's own overlap with the requested range - so a range
        # starting mid-June only counts June sales from that start date on
        window_start = max(month_start, range_start_dt)
        window_end = min(month_end, range_end_dt)
        if window_start > window_end:
            continue

        fg_costs = get_all_fg_total_costs(month)

        # ERP sales + hand-uploaded Other Sales, same filters on both
        sales = []
        for db_path, table in ((sales_db, "sales_register"), (OTHER_SALES_DB, "other_sales")):
            if not os.path.exists(db_path):
                continue
            conn = sqlite3.connect(db_path) if table == "sales_register" else get_other_sales_connection()
            if table == "sales_register":
                ensure_sales_register_year_month(conn)
            query = f"""
                SELECT unit, customer, item_code, quantity, taxable_value
                FROM {table}
                WHERE year_month = ?
                  AND item_code IS NOT NULL AND TRIM(item_code) != ''
                  AND substr(invoice_date,7,4) || substr(invoice_date,4,2) || substr(invoice_date,1,2)
                      BETWEEN ? AND ?
            """
            params = [
                month,
                f"{window_start.year}{window_start.month:02d}{window_start.day:02d}",
                f"{window_end.year}{window_end.month:02d}{window_end.day:02d}",
            ]
            if units:
                placeholders = ", ".join("?" for _ in units)
                query += f" AND unit IN ({placeholders})"
                params.extend(units)
            try:
                sales.extend(conn.execute(query, params).fetchall())
            except sqlite3.OperationalError:
                pass
            conn.close()

        for unit, customer, item_code, quantity, taxable_value in sales:
            try:
                qty = float(quantity) if quantity else 0
                revenue = float(taxable_value) if taxable_value else 0
            except (TypeError, ValueError):
                continue

            cost_info = fg_costs.get(item_code)
            if cost_info is None:
                # No BOM tree at all for this FG this month (BOM Register
                # never run for it, or genuinely untraceable) - treat as
                # fully missing rather than silently costing it at 0
                is_missing = True
                cost = 0
            else:
                is_missing = cost_info["is_missing"]
                cost = cost_info["total_cost"] * qty

            if is_missing and not include_missing:
                continue

            rows.append({
                "month": month, "unit": unit, "customer": customer or "(no customer)",
                "fg_code": item_code, "qty": qty, "revenue": revenue, "cost": cost,
                "margin": revenue - cost, "is_missing": is_missing,
            })

    return rows


def get_available_sales_units():
    units = set()
    for db_path, table in ((os.path.join(DATABASE_DIR, "Sales_Register.db"), "sales_register"),
                           (OTHER_SALES_DB, "other_sales")):
        if not os.path.exists(db_path):
            continue
        conn = sqlite3.connect(db_path)
        try:
            units.update(r[0] for r in conn.execute(f"SELECT DISTINCT unit FROM {table} WHERE unit IS NOT NULL"))
        except sqlite3.OperationalError:
            pass
        conn.close()
    return sorted(units)


# ---------- Other Sales (uploaded by hand - sales not made through this ERP) ----------
#
# Kept in their OWN database file, deliberately NOT inside Sales_Register.db:
# if Sales_Register.db is ever wiped and re-synced from the ERP, these
# hand-uploaded rows must survive. Everything that reads sales (Overview,
# Weighted Selling Price, BOM Register's list of sold FGs) reads both.

OTHER_SALES_DB = os.path.join(DATABASE_DIR, "Other_Sales.db")

# (column header in the Excel file, internal name, required?)
OTHER_SALES_COLUMNS = [
    ("Invoice Date", "invoice_date", True),
    ("Invoice No", "invoice_no", False),
    ("Unit", "unit", True),
    ("Customer", "customer", True),
    ("FG Code", "item_code", True),
    ("Quantity", "quantity", True),
    ("Taxable Value", "taxable_value", True),
]


def get_other_sales_connection():
    conn = sqlite3.connect(OTHER_SALES_DB)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS other_sales (
            row_key TEXT PRIMARY KEY,
            invoice_date TEXT,
            invoice_no TEXT,
            unit TEXT,
            customer TEXT,
            item_code TEXT,
            quantity TEXT,
            taxable_value TEXT,
            year_month TEXT,
            uploaded_at TEXT
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_other_sales_month ON other_sales(year_month, unit)")
    conn.commit()
    return conn


def parse_other_sales_date(value):
    """Accepts a real Excel date cell, or text like 15/08/2026 or 15-08-2026.
    Returns dd/mm/yyyy text (the format Sales Register uses) or None."""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    text = str(value).strip() if value is not None else ""
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return None


@app.route("/other_sales/template")
def other_sales_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Other Sales"
    ws.append([header for header, _, _ in OTHER_SALES_COLUMNS])
    for col_letter, width in zip("ABCDEFG", (14, 16, 8, 36, 14, 12, 16)):
        ws.column_dimensions[col_letter].width = width

    notes = wb.create_sheet("How to fill")
    for line in [
        ["Fill the 'Other Sales' sheet: one row per invoice line. Row 1 (the headers) must stay as it is."],
        [""],
        ["Invoice Date", "Required. dd/mm/yyyy, e.g. 15/08/2026 (an Excel date cell also works)."],
        ["Invoice No", "Optional. For your reference only."],
        ["Unit", "Required. The unit that invoiced the sale, e.g. U1."],
        ["Customer", "Required. Customer name."],
        ["FG Code", "Required. YOUR FG code (e.g. FG5558), NOT the customer's part number."],
        ["Quantity", "Required. Number greater than 0."],
        ["Taxable Value", "Required. Rupees BEFORE GST, for the whole line (not per piece)."],
        [""],
        ["Uploading the same row twice is safe: exact duplicates are skipped."],
        ["If any row has a problem, nothing is saved - fix the rows listed and upload again."],
    ]:
        notes.append(line)
    notes.column_dimensions["A"].width = 16
    notes.column_dimensions["B"].width = 80

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="Other_Sales_Template.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/other_sales/upload", methods=["POST"])
def other_sales_upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    try:
        wb = openpyxl.load_workbook(request.files["file"], data_only=True)
    except Exception as e:
        return jsonify({"error": f"Could not read the file as Excel (.xlsx): {e}"}), 400

    ws = wb["Other Sales"] if "Other Sales" in wb.sheetnames else wb.worksheets[0]
    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        return jsonify({"error": "The file is empty"}), 400

    # Match headers by name (case/spacing-insensitive), so column order doesn't matter
    header_pos = {}
    for i, h in enumerate(all_rows[0]):
        if h is not None:
            header_pos[" ".join(str(h).split()).lower()] = i
    missing = [header for header, _, required in OTHER_SALES_COLUMNS
               if required and header.lower() not in header_pos]
    if missing:
        return jsonify({"error": "Missing column(s): " + ", ".join(missing)
                        + ". Please use the template's headers."}), 400

    def cell(row, header):
        i = header_pos.get(header.lower())
        return row[i] if i is not None and i < len(row) else None

    records, errors = [], []
    for row_no, row in enumerate(all_rows[1:], start=2):
        if row is None or all(v is None or str(v).strip() == "" for v in row):
            continue   # blank line

        problems = []
        date = parse_other_sales_date(cell(row, "Invoice Date"))
        if date is None:
            problems.append("Invoice Date is not a valid dd/mm/yyyy date")
        unit = str(cell(row, "Unit") or "").strip().upper()
        if not unit:
            problems.append("Unit is blank")
        customer = str(cell(row, "Customer") or "").strip()
        if not customer:
            problems.append("Customer is blank")
        fg = str(cell(row, "FG Code") or "").strip().upper()
        if not fg:
            problems.append("FG Code is blank")
        try:
            qty = float(cell(row, "Quantity"))
            if qty <= 0:
                problems.append("Quantity must be greater than 0")
        except (TypeError, ValueError):
            problems.append("Quantity is not a number")
            qty = None
        try:
            value = float(cell(row, "Taxable Value"))
            if value < 0:
                problems.append("Taxable Value can't be negative")
        except (TypeError, ValueError):
            problems.append("Taxable Value is not a number")
            value = None

        if problems:
            errors.append(f"Row {row_no}: " + "; ".join(problems))
            continue

        invoice_no = str(cell(row, "Invoice No") or "").strip()
        qty_text, value_text = f"{qty:g}", f"{round(value, 2):.2f}"
        row_key = hashlib.sha1("|".join(
            [date, invoice_no, unit, customer.upper(), fg, qty_text, value_text]).encode()).hexdigest()
        records.append((row_key, date, invoice_no, unit, customer, fg, qty_text, value_text,
                        f"{date[6:10]}_{date[3:5]}"))

    if errors:
        shown = errors[:20]
        more = f" (and {len(errors) - 20} more)" if len(errors) > 20 else ""
        return jsonify({"error": f"Nothing was saved - {len(errors)} row(s) need fixing{more}:",
                        "details": shown}), 400
    if not records:
        return jsonify({"error": "No data rows found under the headers"}), 400

    # Non-blocking check: codes that don't exist in Item Master are often a
    # customer part number typed by mistake instead of our FG code
    unknown = []
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    if os.path.exists(prod_db):
        pconn = sqlite3.connect(prod_db)
        try:
            for code in sorted({r[5] for r in records}):
                if pconn.execute("SELECT 1 FROM item_master WHERE erp_code = ?", (code,)).fetchone() is None:
                    unknown.append(code)
        except sqlite3.OperationalError:
            pass
        pconn.close()

    uploaded_at = datetime.now().strftime("%d/%m/%Y %H:%M")
    conn = get_other_sales_connection()
    before = conn.total_changes
    conn.executemany(
        "INSERT OR IGNORE INTO other_sales (row_key, invoice_date, invoice_no, unit, customer, item_code, "
        "quantity, taxable_value, year_month, uploaded_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        [r + (uploaded_at,) for r in records],
    )
    conn.commit()
    added = conn.total_changes - before
    conn.close()

    months = sorted({r[8] for r in records})
    month_labels = [datetime(int(m[:4]), int(m[5:]), 1).strftime("%B %Y") for m in months]
    return jsonify({
        "added": added,
        "skipped_duplicates": len(records) - added,
        "months": month_labels,
        "unknown_codes": unknown[:15],
        "unknown_count": len(unknown),
    })


@app.route("/other_sales/download")
def other_sales_download():
    """Everything uploaded so far, in the same layout as the template."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Other Sales"
    ws.append([header for header, _, _ in OTHER_SALES_COLUMNS] + ["Uploaded At"])
    if os.path.exists(OTHER_SALES_DB):
        conn = get_other_sales_connection()
        for r in conn.execute(
            "SELECT invoice_date, invoice_no, unit, customer, item_code, quantity, taxable_value, uploaded_at "
            "FROM other_sales ORDER BY year_month, substr(invoice_date,1,2), item_code"
        ):
            ws.append([r[0], r[1], r[2], r[3], r[4], float(r[5]), float(r[6]), r[7]])
        conn.close()
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="Other_Sales_Uploaded.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/overview/fg_detail")
def overview_fg_detail():
    """
    Breakdown behind one FG in Overview's negative-margin list: since
    Overview can span multiple months, customers and units, this doesn't
    show a single BOM Tree cost breakdown (Part-wise Margin's popup does
    that, for one specific month/customer/unit) - instead it shows every
    (month, customer, unit) combination that contributed to this FG's
    total within the selected range, so it's clear WHERE a negative margin
    is actually coming from.
    """
    fg_code = request.args.get("fg_code")
    from_date = request.args.get("from_date")
    to_date = request.args.get("to_date")
    units_param = request.args.getlist("units")
    include_missing = request.args.get("include_missing", "1") == "1"

    if not fg_code or not from_date or not to_date:
        return jsonify({"error": "Missing required parameters"}), 400

    try:
        all_rows = compute_margin_rows(from_date, to_date, units=units_param or None, include_missing=include_missing)

        fg_rows = [r for r in all_rows if r["fg_code"] == fg_code]
        if not fg_rows:
            return jsonify({"error": f"{fg_code} has no sales in this range"}), 404

        total_revenue = sum(r["revenue"] for r in fg_rows)
        total_cost = sum(r["cost"] for r in fg_rows)
        total_margin = total_revenue - total_cost
        total_qty = sum(r["qty"] for r in fg_rows)

        breakdown = sorted(fg_rows, key=lambda r: r["margin"])

        return jsonify({
            "fg_code": fg_code,
            "total_revenue": round(total_revenue, 2),
            "total_cost": round(total_cost, 2),
            "total_margin": round(total_margin, 2),
            "total_qty": round(total_qty, 4),
            "breakdown": [
                {
                    "month": r["month"], "unit": r["unit"], "customer": r["customer"],
                    "qty": round(r["qty"], 4), "revenue": round(r["revenue"], 2),
                    "cost": round(r["cost"], 2), "margin": round(r["margin"], 2),
                    "is_missing": r["is_missing"],
                }
                for r in breakdown
            ],
        })
    except Exception as e:
        return jsonify({"error": f"Could not load detail: {e}"}), 500


@app.route("/")
def overview():
    from_date = request.args.get("from_date", "").strip()
    to_date = request.args.get("to_date", "").strip()
    selected_units = request.args.getlist("units")

    # Checkboxes only submit a value when CHECKED - when unchecked, the param
    # is absent from the URL entirely, so a plain .get(..., "1") default would
    # make unticking the box do nothing (it'd always fall back to "1"). Using
    # "from_date" (a plain text input, always present once the form has been
    # submitted at all) to tell "fresh page load" apart from "form submitted
    # with the box unticked" fixes that.
    form_submitted = "from_date" in request.args
    include_missing = (request.args.get("include_missing") == "1") if form_submitted else True

    if not from_date or not to_date:
        now = datetime.now()
        from_date = f"01/{now.month:02d}/{now.year}"
        to_date = now.strftime("%d/%m/%Y")

    try:
        rows = compute_margin_rows(from_date, to_date, units=selected_units or None, include_missing=include_missing)
    except ValueError:
        rows = []   # bad date input - show empty rather than crash

    total_revenue = sum(r["revenue"] for r in rows)
    total_cost = sum(r["cost"] for r in rows)
    total_margin = total_revenue - total_cost
    margin_pct = round((total_margin / total_cost * 100), 2) if total_cost else 0
    missing_lines = sum(1 for r in rows if r["is_missing"])

    def group_by(key):
        groups = {}
        for r in rows:
            k = r[key]
            g = groups.setdefault(k, {"revenue": 0, "cost": 0, "qty": 0})
            g["revenue"] += r["revenue"]
            g["cost"] += r["cost"]
            g["qty"] += r["qty"]
        result = []
        for k, v in groups.items():
            margin = v["revenue"] - v["cost"]
            result.append({
                "key": k, "revenue": round(v["revenue"], 2), "cost": round(v["cost"], 2),
                "qty": round(v["qty"], 2), "margin": round(margin, 2),
                "margin_pct": round((margin / v["cost"] * 100), 2) if v["cost"] else 0,
            })
        result.sort(key=lambda r: r["margin"])
        return result

    by_customer = group_by("customer")
    by_unit = group_by("unit")
    by_month = sorted(group_by("month"), key=lambda r: r["key"])   # chronological, not by margin

    by_fg = group_by("fg_code")
    negative_fgs = [r for r in by_fg if r["margin"] < 0]
    negative_fgs.sort(key=lambda r: r["margin"])

    # ---- vs Previous Period ----
    prev_from, prev_to = get_previous_period_dates(from_date, to_date)
    try:
        prev_rows = compute_margin_rows(prev_from, prev_to, units=selected_units or None, include_missing=include_missing)
    except ValueError:
        prev_rows = []

    prev_total_revenue = sum(r["revenue"] for r in prev_rows)
    prev_total_cost = sum(r["cost"] for r in prev_rows)
    prev_total_margin = prev_total_revenue - prev_total_cost
    prev_margin_pct = round((prev_total_margin / prev_total_cost * 100), 2) if prev_total_cost else 0
    prev_missing_lines = sum(1 for r in prev_rows if r["is_missing"])

    trends = {
        "revenue": pct_change(total_revenue, prev_total_revenue),
        "cost": pct_change(total_cost, prev_total_cost),
        "margin": pct_change(total_margin, prev_total_margin),
        # Margin % is already a percentage, so its own "change" is shown as a
        # plain point difference (26.60% vs 24.28% -> "+2.32"), not a percent
        # change of a percent, which would read as confusing.
        "margin_pct": round(margin_pct - prev_margin_pct, 2),
        "missing_lines": pct_change(missing_lines, prev_missing_lines),
    }

    return render_template(
        "overview.html",
        nav_items=NAV_ITEMS, active_page="overview", page_label="Overview",
        months=[],
        from_date=from_date, to_date=to_date,
        available_units=get_available_sales_units(), selected_units=selected_units,
        include_missing=include_missing,
        total_revenue=round(total_revenue, 2), total_cost=round(total_cost, 2),
        total_margin=round(total_margin, 2), margin_pct=margin_pct,
        missing_lines=missing_lines, total_lines=len(rows),
        by_customer=by_customer, by_unit=by_unit, by_month=by_month,
        negative_fgs=negative_fgs,
        trends=trends, prev_from=prev_from, prev_to=prev_to,
    )


_part_cost_lookup_cache = {}   # {month: (computed_dict, cached_at_timestamp)}
_PART_COST_CACHE_TTL_SECONDS = 600   # 10 min - safety net in case an invalidation trigger is ever missed


def invalidate_cost_cache():
    """Clears the Part Cost Calc lookup cache. Call this after anything
    that could change a part's cost: a Data Sync script finishing, Other
    Cost being saved, Cycle Time being saved/uploaded, or Item Type being
    filled/uploaded. Without this, BOM Tree/Part-wise Margin/Overview would
    keep showing yesterday's costs even right after fresh data has been
    synced - the TTL alone would eventually catch up, but a deliberate
    change should be reflected immediately, not after a wait."""
    _part_cost_lookup_cache.clear()


def get_part_costs_for_lookup(month):
    """
    All Part Cost Calc totals for a month, keyed for tree lookups:
        {(item_code, process, unit): total_cost_per_part}
    Uses the same rules as the Part Cost Calc screen (purchased materials in
    three buckets + allocated overhead), just unpaginated and unfiltered.

    This is the single most expensive computation in the app, and BOM Tree,
    Part-wise Margin, and Overview all need it - so unlike find_sfg_cost's
    cost_cache (which only lasts one request), this result is cached
    ACROSS requests. Without it, navigating between those three pages for
    the same month reruns the full Part Cost Calc aggregation from scratch
    every single time. Invalidated explicitly via invalidate_cost_cache()
    whenever underlying data actually changes, with a TTL as a safety net.

    Returns {} for months where Item Master was never run, so the tree's
    month-by-month fallback can simply skip past them instead of crashing.
    """
    cached = _part_cost_lookup_cache.get(month)
    if cached is not None:
        result, cached_at = cached
        if time.time() - cached_at < _PART_COST_CACHE_TTL_SECONDS:
            return result

    if not get_item_type_month_column_exists(month):
        return {}

    rows, _total = build_part_cost_calc(month, filters={}, page=1, page_size_override=None)
    result = {
        (r["item_code"], r["process"], r["unit"]): r["total_cost_per_part"]
        for r in rows
    }
    _part_cost_lookup_cache[month] = (result, time.time())
    return result


def find_sfg_cost(item_code, process, unit, month, cost_cache, months_desc):
    """
    Finds a Production-type SFG's cost using the agreed fallback chain.
    Process must ALWAYS match - only the unit is relaxed, and only after the
    same unit fails:
        1. this month, same process, same unit
        2. this month, same process, any unit
        3. earlier months, same rules, most recent first

    Note the process/unit passed in are a HINT (they come from the BOM edge,
    which describes where the item was CONSUMED, not where it was made). If
    no match is found on that hint, we fall back to the item's own
    manufacturing process - an SFG consumed in ASSEMBLY may well have been
    made in INJECTION MOULDING, and its cost lives under the latter.

    Returns (cost, source_month, source_process, source_unit) or (None, None, None, None).
    """
    for m in months_desc:
        if m > month:
            continue   # never look into the future
        costs = cost_cache.get(m)
        if costs is None:
            costs = get_part_costs_for_lookup(m)
            cost_cache[m] = costs

        # 1. hinted process, same unit
        exact = costs.get((item_code, process, unit))
        if exact is not None:
            return exact, m, process, unit

        # 2. hinted process, any other unit
        for (code, proc, u), cost in costs.items():
            if code == item_code and proc == process:
                return cost, m, proc, u

        # 3. the item's OWN manufacturing process - prefer the same unit
        own = [(c, p, u) for (c, p, u) in costs if c == item_code]
        same_unit = [k for k in own if k[2] == unit]
        for key in (same_unit + own):
            return costs[key], m, key[1], key[2]

    return None, None, None, None


def get_fg_root_cost(root_fg, process, unit, month, item_types, weighted_prices, cost_cache, months_desc):
    """
    Costs a tree's ROOT FG itself. Real production evidence always wins over
    Item Master's classification: if Part Cost Calc has data for this exact
    FG this month, that means it was genuinely manufactured then, REGARDLESS
    of what Item Master says its type is - a code can be nominally tagged
    "Purchase" while concrete production entries exist for it that month
    (Item Master data can be stale, wrong, or simply not kept in sync with
    what actually got produced). Only when there's no production evidence
    anywhere does the Purchase classification + weighted price kick in as
    the fallback - matching how a genuinely outsourced FG (which never has
    its own production entries) should be priced.

    Returns (cost, source_label); cost is None if genuinely not found.
    """
    d = get_fg_root_cost_detail(root_fg, process, unit, month, item_types, weighted_prices, cost_cache, months_desc)
    return d["cost"], d["label"]


def get_fg_root_cost_detail(root_fg, process, unit, month, item_types, weighted_prices, cost_cache, months_desc):
    """
    Same costing as get_fg_root_cost (which is now a thin wrapper around
    this, so every existing caller gets exactly the same result), but also
    reports WHERE the root's own cost came from - the process and unit it
    was actually found under, and whether it was treated as made in-house
    or bought - so the summary popup can show the FG itself as a proper
    first row, not just a "cost source" note.
    """
    cost, src_month, src_process, src_unit = find_sfg_cost(root_fg, process, unit, month, cost_cache, months_desc)
    if cost is not None:
        return {"cost": cost, "label": f"Part cost ({src_month}, {src_process}, {src_unit})",
                "process": src_process, "unit": src_unit, "type": "Production"}

    if item_types.get(root_fg) == "Purchase":
        price = weighted_prices.get(root_fg)
        # A literal 0 here means "never actually purchased" (Weighted Price
        # Master's own carry-forward default when there's no purchase
        # history at all) - treat that the same as no price found, not as a
        # genuine free/zero-cost item.
        if price:
            return {"cost": price, "label": "Weighted price (outsourced FG)",
                    "process": process, "unit": unit, "type": "Purchase"}
        return {"cost": None, "label": "No cost found", "process": process, "unit": unit, "type": "Purchase"}

    return {"cost": None, "label": "No cost found", "process": process, "unit": unit,
            "type": item_types.get(root_fg)}


def flatten_tree_nodes(nodes):
    """Walks a nested tree_nodes structure into one flat list, every level."""
    flat = []
    for n in nodes:
        flat.append(n)
        flat.extend(flatten_tree_nodes(n["children"]))
    return flat


def build_tree_nodes(root_fg, month, bom_rows, item_types, weighted_prices, cost_cache, months_desc, stats):
    """
    Turns the stored flat parent->child edges into a nested tree, attaching
    each node's cost as we go.

    Costing rules:
      - Purchase-type SFG  -> 0, greyed out (its price is already counted in
                              the parent's Part Cost Calc as BOP material)
      - Purchase-type FG   -> its Weighted Price (outsourced finished good,
                              no parent stage absorbed its cost)
      - Production SFG     -> its Part Cost Calc, with the unit/month fallback
      - Not found anywhere -> flagged missing, counted on the root
    """
    children_by_parent = {}
    for r in bom_rows:
        children_by_parent.setdefault(r["parent_code"], []).append(r)

    def build(code, process, unit, level, path):
        if code in path:
            return []   # cycle guard

        nodes = []
        for edge in children_by_parent.get(code, []):
            child_code = edge["child_code"]
            child_type = item_types.get(child_code)
            child_process = edge["process"]
            child_unit = edge["unit"]

            is_fg = "FG" in str(child_code).upper() and "SFG" not in str(child_code).upper()
            is_sfg = "SFG" in str(child_code).upper()

            if not is_fg and not is_sfg:
                # A pure material (RMI/RMP/PACK/ASP/CONS/etc.) - the tree only
                # ever shows FG/SFG codes. Its cost is already inside the
                # parent's Part Cost Calc either way, and it never has
                # children of its own in practice (only Production-type
                # FG/SFG codes get expanded when the tree is built), so
                # skipping it here doesn't hide any structure.
                continue

            cost = 0
            cost_source = ""
            is_purchase_sfg = False
            is_missing = False

            if child_type == "Purchase":
                if is_fg:
                    # Outsourced FG - priced at what we pay for it. A literal
                    # 0 here means "never actually purchased" (Weighted Price
                    # Master's own carry-forward default), not a genuine
                    # free/zero-cost item, so treat it the same as missing.
                    price = weighted_prices.get(child_code)
                    if price:
                        cost = price
                        cost_source = "Weighted price"
                    else:
                        is_missing = True
                        stats["missing"] += 1
                        stats["missing_codes"].append(child_code)
                        cost_source = "No cost found"
                else:
                    # Purchased SFG/material - already inside the parent's Part Cost Calc
                    cost = 0
                    is_purchase_sfg = True
                    cost_source = "Already in parent's cost"
            elif child_type == "Production" and is_sfg:
                found, src_month, src_process, src_unit = find_sfg_cost(
                    child_code, child_process, child_unit, month, cost_cache, months_desc
                )
                if found is None:
                    is_missing = True
                    stats["missing"] += 1
                    stats["missing_codes"].append(child_code)
                    cost_source = "No cost found"
                else:
                    cost = found
                    cost_source = f"Part cost ({src_month}, {src_process}, {src_unit})"
                    child_process = src_process   # show the process the price actually came from, not just the consuming stage's process
                    child_unit = src_unit
            else:
                # Unclassified, or a non-FG/SFG code that slipped in
                cost = 0
                cost_source = "No item type" if child_type is None else ""
                if child_type is None:
                    is_missing = True
                    stats["missing"] += 1
                    stats["missing_codes"].append(child_code)

            sub_nodes = build(child_code, child_process, child_unit, level + 1, path + [code])

            # Each node shows ONLY its own individual cost - no nested rollup
            # from its children. The tree's grand total (computed separately,
            # flattened across every node) is the only place costs get added
            # together, so a code like SFGA2253 always reads as exactly its
            # own Part Cost Calc figure, not "itself + everything under it".
            norm_qty = edge["norm_qty"] or 0
            extended = cost * norm_qty

            nodes.append({
                "code": child_code,
                "item": edge["child_item"],
                "item_type": child_type,
                "process": child_process,
                "unit": child_unit,
                "norm_qty": round(norm_qty, 6),
                "unit_cost": round(cost, 4),
                "extended_cost": round(extended, 4),
                "cost_source": cost_source,
                "is_purchase_sfg": is_purchase_sfg,
                "is_missing": is_missing,
                "level": level,
                "children": sub_nodes,
            })

        return nodes

    return build(root_fg, None, None, 1, [])


def build_production_only_paths(month, root_fgs=None):
    """
    Flattens stored BOM trees into the NXT1...NXTn shape - one row per
    branch, from the root FG down through PRODUCTION-type codes only.
    Purchase-type children (whether SFG or FG) end the chain and are NOT
    written into the sheet at all - their cost already lives elsewhere
    (Part Cost Calc / weighted price), this view is just the manufacturing
    chain itself.
    """
    conn = sqlite3.connect(ANALYTICS_DB)
    cur = conn.cursor()
    try:
        query = "SELECT root_fg, parent_code, child_code, child_item_type FROM bom_tree WHERE month = ?"
        params = [month]
        if root_fgs:
            placeholders = ", ".join("?" for _ in root_fgs)
            query += f" AND root_fg IN ({placeholders})"
            params.extend(root_fgs)
        cur.execute(query, params)
        edges = cur.fetchall()
    except sqlite3.OperationalError:
        edges = []
    conn.close()

    # Group Production-type children by (root_fg, parent_code)
    children_by_root_parent = {}
    for root_fg, parent_code, child_code, child_item_type in edges:
        if child_item_type != "Production":
            continue
        children_by_root_parent.setdefault((root_fg, parent_code), []).append(child_code)

    rows = []
    max_depth = 0

    def walk(root_fg, code, chain):
        children = children_by_root_parent.get((root_fg, code), [])
        if not children:
            if chain:   # only emit a row once we've gone at least one level deep
                rows.append([root_fg] + chain)
            return
        for child in children:
            walk(root_fg, child, chain + [child])

    roots = sorted({r for r, _, _, _ in edges} if not root_fgs else root_fgs)
    for root_fg in roots:
        walk(root_fg, root_fg, [])

    max_depth = max((len(r) - 1 for r in rows), default=0)
    return rows, max_depth


@app.route("/tree/download_sheet")
def tree_download_sheet():
    month = request.args.get("month")
    fg = request.args.get("fg", "").strip()

    rows, max_depth = build_production_only_paths(month, root_fgs=[fg] if fg else None)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["FG Code"] + [f"NXT{i+1}" for i in range(max_depth)])
    for row in rows:
        padded = row + [None] * (max_depth - (len(row) - 1))
        ws.append(padded)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    suffix = fg if fg else "all"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"bom_tree_{suffix}_{month}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


ALL_FG_TABLE_PAGE_SIZE = 200
ALL_FG_TABLE_FILTER_FIELDS = ["root_fg", "code", "item", "item_type", "process", "unit"]


def build_all_fgs_cost_table(month, filters=None):
    """
    Flattens EVERY sold FG's tree for the month into one big table - one row
    per code (the FG root itself, plus every sub-item under it), each
    showing only its OWN individual cost - exactly matching how a single
    node reads in the per-FG tree view. This is the "whole month at a
    glance" view, versus picking one FG at a time.
    """
    conn = sqlite3.connect(ANALYTICS_DB)
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT root_fg FROM bom_tree WHERE month = ? ORDER BY root_fg", (month,))
    root_fgs = [r[0] for r in cur.fetchall()]

    item_types = get_item_types_for_month(month)
    weighted_prices = get_weighted_prices_for_month(month)
    months = get_all_relevant_months()
    months_desc = sorted([m for m in months if m <= month], reverse=True)
    cost_cache = {}   # shared across every FG in this call for efficiency

    all_rows = []

    for root_fg in root_fgs:
        cur.execute("""
            SELECT root_fg, parent_code, child_code, child_item, norm_qty, level,
                   company, unit, process, wip_store, child_item_type
            FROM bom_tree WHERE month = ? AND root_fg = ?
        """, (month, root_fg))
        bom_rows = [
            {"root_fg": r[0], "parent_code": r[1], "child_code": r[2], "child_item": r[3],
             "norm_qty": r[4], "level": r[5], "company": r[6], "unit": r[7],
             "process": r[8], "wip_store": r[9], "child_item_type": r[10]}
            for r in cur.fetchall()
        ]
        if not bom_rows:
            continue

        stats = {"missing": 0, "missing_codes": []}
        tree_nodes = build_tree_nodes(root_fg, month, bom_rows, item_types, weighted_prices, cost_cache, months_desc, stats)

        root_edges = [r for r in bom_rows if r["parent_code"] == root_fg]
        root_process = root_edges[0]["process"] if root_edges else None
        root_unit_hint = root_edges[0]["unit"] if root_edges else None
        root_cost, root_source_label = get_fg_root_cost(
            root_fg, root_process, root_unit_hint, month, item_types, weighted_prices, cost_cache, months_desc
        )
        root_type = item_types.get(root_fg) or "Production"
        all_rows.append({
            "root_fg": root_fg, "code": root_fg, "item": root_fg, "level": 0,
            "item_type": root_type, "process": root_process, "unit": root_unit_hint,
            "norm_qty": 1, "unit_cost": round(root_cost or 0, 4),
            "extended_cost": round(root_cost or 0, 4),
            "cost_source": root_source_label,
            "is_missing": root_cost is None,
        })

        for n in flatten_tree_nodes(tree_nodes):
            all_rows.append({
                "root_fg": root_fg, "code": n["code"], "item": n["item"], "level": n["level"],
                "item_type": n["item_type"], "process": n["process"], "unit": n["unit"],
                "norm_qty": n["norm_qty"], "unit_cost": n["unit_cost"],
                "extended_cost": n["extended_cost"], "cost_source": n["cost_source"],
                "is_missing": n["is_missing"],
            })

    conn.close()

    filters = filters or {}
    for field, value in filters.items():
        value = (value or "").strip().lower()
        if value:
            all_rows = [r for r in all_rows if value in str(r.get(field, "")).lower()]

    return all_rows


@app.route("/tree/download_all")
def tree_download_all():
    month = request.args.get("month")
    rows = build_all_fgs_cost_table(month)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["FG Code", "Code", "Item", "Level", "Type", "Process", "Unit", "Norm Qty", "Unit Cost", "Extended Cost", "Source", "Missing"])
    for r in rows:
        ws.append([
            r["root_fg"], r["code"], r["item"], r["level"], r["item_type"],
            r["process"], r["unit"], r["norm_qty"], r["unit_cost"], r["extended_cost"],
            r["cost_source"], "YES" if r["is_missing"] else "",
        ])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"bom_tree_all_fgs_priced_{month}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/tree")
def tree():
    months = get_all_relevant_months()
    selected_month = request.args.get("month", months[0] if months else None)
    selected_fg = request.args.get("fg", "").strip()

    # Which FGs have a stored tree for this month
    available_fgs = []
    conn = sqlite3.connect(ANALYTICS_DB)
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT DISTINCT root_fg FROM bom_tree WHERE month = ? ORDER BY root_fg
        """, (selected_month,))
        available_fgs = [r[0] for r in cur.fetchall()]
    except sqlite3.OperationalError:
        available_fgs = []

    tree_nodes = []
    root_info = None
    stats = {"missing": 0, "missing_codes": []}

    if selected_fg and available_fgs:
        cur.execute("""
            SELECT root_fg, parent_code, child_code, child_item, norm_qty, level,
                   company, unit, process, wip_store, child_item_type
            FROM bom_tree WHERE month = ? AND root_fg = ?
        """, (selected_month, selected_fg))
        bom_rows = [
            {"root_fg": r[0], "parent_code": r[1], "child_code": r[2], "child_item": r[3],
             "norm_qty": r[4], "level": r[5], "company": r[6], "unit": r[7],
             "process": r[8], "wip_store": r[9], "child_item_type": r[10]}
            for r in cur.fetchall()
        ]

        if bom_rows:
            item_types = get_item_types_for_month(selected_month)
            weighted_prices = get_weighted_prices_for_month(selected_month)
            cost_cache = {}
            months_desc = sorted([m for m in months if m <= selected_month], reverse=True)

            tree_nodes = build_tree_nodes(
                selected_fg, selected_month, bom_rows, item_types,
                weighted_prices, cost_cache, months_desc, stats
            )

            # The root FG's own stage cost - Purchase-type (outsourced) FGs use
            # their weighted price, Production-type use Part Cost Calc, same
            # branching every child node in the tree gets.
            root_edges = [r for r in bom_rows if r["parent_code"] == selected_fg]
            root_process = root_edges[0]["process"] if root_edges else None
            root_unit_hint = root_edges[0]["unit"] if root_edges else None
            root_cost, root_source_label = get_fg_root_cost(
                selected_fg, root_process, root_unit_hint, selected_month,
                item_types, weighted_prices, cost_cache, months_desc
            )

            # The grand total sums every node's OWN extended cost across the
            # WHOLE tree (every level, not just direct children) - this is
            # the only place costs get added together; each individual node
            # in the tree itself always shows just its own figure.
            descendants_total = sum(n["extended_cost"] for n in flatten_tree_nodes(tree_nodes))
            root_info = {
                "code": selected_fg,
                "own_cost": round(root_cost or 0, 4),
                "own_cost_found": root_cost is not None,
                "own_cost_source": root_source_label,
                "descendants_total": round(descendants_total, 4),
                "total_cost": round((root_cost or 0) + descendants_total, 4),
                "missing_count": stats["missing"],
            }
            if root_cost is None:
                stats["missing"] += 1
                root_info["missing_count"] = stats["missing"]

    conn.close()

    sheet_rows, sheet_max_depth = ([], 0)
    if selected_fg and available_fgs:
        sheet_rows, sheet_max_depth = build_production_only_paths(selected_month, root_fgs=[selected_fg])

    # When no specific FG is picked, show every FG's costed tree flattened
    # into one table - the "whole month" view - instead of a bare "pick an
    # FG" placeholder.
    all_fg_rows = []
    all_fg_total_count = 0
    all_fg_total_pages = 1
    all_fg_page = 1
    all_fg_filters = {}
    if not selected_fg:
        all_fg_filters = {field: request.args.get(field, "") for field in ALL_FG_TABLE_FILTER_FIELDS}
        all_fg_page = max(1, request.args.get("page", 1, type=int))
        full_table = build_all_fgs_cost_table(selected_month, filters=all_fg_filters)
        all_fg_total_count = len(full_table)
        all_fg_total_pages = max(1, -(-all_fg_total_count // ALL_FG_TABLE_PAGE_SIZE))
        start = (all_fg_page - 1) * ALL_FG_TABLE_PAGE_SIZE
        all_fg_rows = full_table[start:start + ALL_FG_TABLE_PAGE_SIZE]

    return render_analytics_page(
        "tree.html", "tree",
        months_override=months,
        available_fgs=available_fgs,
        selected_fg=selected_fg,
        tree_nodes=tree_nodes,
        root_info=root_info,
        missing_codes=sorted(set(stats["missing_codes"])),
        sheet_rows=sheet_rows,
        sheet_max_depth=sheet_max_depth,
        all_fg_rows=all_fg_rows,
        all_fg_total_count=all_fg_total_count,
        all_fg_total_pages=all_fg_total_pages,
        all_fg_page=all_fg_page,
        all_fg_filters=all_fg_filters,
    )


def get_item_types_for_month(month):
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    if not os.path.exists(prod_db):
        return {}
    item_type_col = f"item_type_{month}"
    conn = sqlite3.connect(prod_db)
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(item_master)")
    if item_type_col not in {r[1] for r in cur.fetchall()}:
        conn.close()
        return {}
    cur.execute(f"SELECT erp_code, {item_type_col} FROM item_master")
    result = {c: t for c, t in cur.fetchall() if t}
    conn.close()
    return result


def ensure_weighted_price_indexes():
    """month-alone indexes on weighted_price_master and
    weighted_selling_price_master. Their PRIMARY KEY is (item_code, month),
    so a plain "WHERE month = ?" query - used constantly, e.g. by
    get_weighted_prices_for_month - can't use that index efficiently since
    month isn't its leading column. Safe/cheap to call defensively -
    CREATE INDEX IF NOT EXISTS is a fast metadata check when the index
    already exists, not a rebuild."""
    if not os.path.exists(ANALYTICS_DB):
        return
    conn = sqlite3.connect(ANALYTICS_DB)
    cur = conn.cursor()
    for table in ("weighted_price_master", "weighted_selling_price_master"):
        try:
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{table}_month ON {table}(month)")
        except sqlite3.OperationalError:
            pass
    conn.commit()
    conn.close()


def get_weighted_prices_for_month(month):
    """{item_code: final_price}. Deliberately does NOT coerce a missing row
    to 0 here - Weighted_Price_Master.py itself already stores a literal 0
    for "never actually purchased" (its own carry-forward default when there's
    no purchase history at all), so a real 0 and a genuinely-absent row both
    need to reach callers as "no real price", not get flattened into the same
    number as a real zero-cost item. Callers that need a plain fallback for
    display purposes can still do `.get(code, 0)` themselves; callers deciding
    whether a PRICE genuinely exists should check truthiness (0 is falsy),
    not just `is not None`.
    """
    if not os.path.exists(ANALYTICS_DB):
        return {}
    conn = sqlite3.connect(ANALYTICS_DB)
    cur = conn.cursor()
    try:
        cur.execute("SELECT item_code, final_price FROM weighted_price_master WHERE month = ?", (month,))
        result = dict(cur.fetchall())
    except sqlite3.OperationalError:
        result = {}
    conn.close()
    return result


def get_item_name(item_code):
    """An item's name, for the ONE FG shown in the summary popup. Only ever
    call this for a single item - never inside a loop over every FG.
    Item Master first (keyed by ERP code, so it's an instant lookup);
    the production register only as a fallback, since it has no index on
    item_code and scanning it is slow on real data (~0.1 s each - fine
    once, but ~2 minutes if done for every FG in a month)."""
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    if not os.path.exists(prod_db):
        return ""
    conn = sqlite3.connect(prod_db)
    name = ""
    try:
        row = conn.execute("SELECT description FROM item_master WHERE erp_code = ?", (item_code,)).fetchone()
        if row and row[0]:
            name = row[0]
    except sqlite3.OperationalError:
        pass
    if not name:
        try:
            row = conn.execute(
                "SELECT item_name FROM production_register WHERE item_code = ? AND item_name IS NOT NULL LIMIT 1",
                (item_code,),
            ).fetchone()
            if row:
                name = row[0]
        except sqlite3.OperationalError:
            pass
    conn.close()
    return name


def get_single_fg_tree_detail(root_fg, month, item_types=None, weighted_prices=None, months_desc=None, cost_cache=None):
    """
    Full cost breakdown for ONE FG: its own cost, every node in its BOM tree
    flattened (each with its own individual cost, per build_tree_nodes), and
    the aggregate total - the same computation get_all_fg_total_costs does
    per FG in its loop, extracted here so the Part-wise Margin summary
    popup can get just one FG's detail on demand without paying the cost of
    computing every FG in the month.
    """
    if item_types is None:
        item_types = get_item_types_for_month(month)
    if weighted_prices is None:
        weighted_prices = get_weighted_prices_for_month(month)
    if months_desc is None:
        months = get_all_relevant_months()
        months_desc = sorted([m for m in months if m <= month], reverse=True)
    if cost_cache is None:
        cost_cache = {}

    conn = sqlite3.connect(ANALYTICS_DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT root_fg, parent_code, child_code, child_item, norm_qty, level,
               company, unit, process, wip_store, child_item_type
        FROM bom_tree WHERE month = ? AND root_fg = ?
    """, (month, root_fg))
    bom_rows = [
        {"root_fg": r[0], "parent_code": r[1], "child_code": r[2], "child_item": r[3],
         "norm_qty": r[4], "level": r[5], "company": r[6], "unit": r[7],
         "process": r[8], "wip_store": r[9], "child_item_type": r[10]}
        for r in cur.fetchall()
    ]
    conn.close()

    stats = {"missing": 0, "missing_codes": []}
    tree_nodes = build_tree_nodes(
        root_fg, month, bom_rows, item_types, weighted_prices, cost_cache, months_desc, stats
    ) if bom_rows else []

    root_edges = [r for r in bom_rows if r["parent_code"] == root_fg]
    root_process = root_edges[0]["process"] if root_edges else None
    root_unit_hint = root_edges[0]["unit"] if root_edges else None
    root_detail = get_fg_root_cost_detail(
        root_fg, root_process, root_unit_hint, month, item_types, weighted_prices, cost_cache, months_desc
    )
    root_cost, root_source = root_detail["cost"], root_detail["label"]

    flat_nodes = flatten_tree_nodes(tree_nodes)
    descendants_total = sum(n["extended_cost"] for n in flat_nodes)
    total_cost = (root_cost or 0) + descendants_total
    missing_count = stats["missing"] + (1 if root_cost is None else 0)

    return {
        "root_fg": root_fg,
        "root_cost": round(root_cost, 4) if root_cost is not None else None,
        "root_source": root_source,
        "root_process": root_process,
        "root_unit": root_unit_hint,
        # Where the FG's own stage cost was actually found (can differ from
        # the hint above), for showing the FG itself as the first row
        "root_cost_process": root_detail["process"],
        "root_cost_unit": root_detail["unit"],
        "root_type": root_detail["type"],
        "sub_items": flat_nodes,
        "total_cost": round(total_cost, 4),
        "is_missing": missing_count > 0,
        "missing_count": missing_count,
        "root_cost_unknown": root_cost is None,
    }


def get_all_fg_total_costs(month):
    """
    For every FG with a stored BOM tree in this month, the FULLY ROLLED UP
    cost - its own Part Cost Calc/weighted price plus every Production
    descendant's own cost, summed across the whole tree. This is the same
    "Total cost/part" figure shown at the top of the per-FG tree view, just
    computed for every FG in the month at once. Both Overview and Part-wise
    Margin need "the cost of this FG this month" as a single number, so they
    share this one function.

    Returns {fg_code: {"total_cost": X, "is_missing": bool, "missing_count": N}}
    """
    conn = sqlite3.connect(ANALYTICS_DB)
    cur = conn.cursor()
    try:
        cur.execute("SELECT DISTINCT root_fg FROM bom_tree WHERE month = ?", (month,))
        root_fgs = [r[0] for r in cur.fetchall()]
    except sqlite3.OperationalError:
        conn.close()
        return {}
    conn.close()

    item_types = get_item_types_for_month(month)
    weighted_prices = get_weighted_prices_for_month(month)
    months = get_all_relevant_months()
    months_desc = sorted([m for m in months if m <= month], reverse=True)
    cost_cache = {}

    result = {}
    for root_fg in root_fgs:
        detail = get_single_fg_tree_detail(root_fg, month, item_types, weighted_prices, months_desc, cost_cache)
        result[root_fg] = {
            "total_cost": detail["total_cost"],
            "is_missing": detail["is_missing"],
            "missing_count": detail["missing_count"],
            # Distinct from is_missing: True only when the ROOT's own cost is
            # fundamentally unknown (e.g. Purchase-type with no weighted
            # price found anywhere) - as opposed to the root being known but
            # some descendant sub-item being untraceable (a genuine partial
            # total, still meaningfully usable). Consumers that would
            # otherwise treat "cost unknown" and "cost is exactly zero" as
            # the same number (0) should check this flag first.
            "root_cost_unknown": detail["root_cost_unknown"],
        }

    return result


PART_COST_PAGE_SIZE = 200

PART_COST_FILTER_COLUMNS = {
    "company": "final.company",
    "unit": "final.unit",
    "store": "final.store",
    "process": "final.process",
    "item_code": "final.item_code",
    "item_name": "final.item_name",
    "uom": "final.uom",
}


def get_store_time_totals(month):
    """
    For each (Unit, Store) in the month: the total 'time' consumed, where a
    part's time = its Cycle Time x its Prod Qty, summed across every part
    produced in that store. Parts with no cycle time captured yet contribute
    0 time (so they simply get no overhead allocated to them).

    Returns:
        store_totals: {(unit, store): total_time}
        part_times:   {(unit, store, process, item_code): part_time}
    """
    combos = get_all_cycle_time_combos(month)

    store_totals = {}
    part_times = {}

    for c in combos:
        cycle_time = c["cycle_time"] or 0      # missing cycle time -> 0 time
        part_time = cycle_time * (c["total_prod_qty"] or 0)

        key = (c["unit"], c["store"], c["process"], c["item_code"])
        part_times[key] = part_time

        store_key = (c["unit"], c["store"])
        store_totals[store_key] = store_totals.get(store_key, 0) + part_time

    return store_totals, part_times


def get_other_costs_for_month(month):
    """
    All overhead figures per (Unit, Store) for the month. TRIMMING stores are
    auto-calculated (flat Rs.1 per piece = total Prod Qty) rather than read
    from the saved table, matching how the Other Cost page behaves.
    """
    result = {}

    conn = get_other_cost_connection()
    cur = conn.cursor()
    col_list = ", ".join(key for key, _ in OTHER_COST_FIELDS)
    cur.execute(f"SELECT unit, store, {col_list} FROM other_cost_master WHERE month = ?", (month,))
    for row in cur.fetchall():
        unit, store = row[0], row[1]
        result[(unit, store)] = dict(zip([key for key, _ in OTHER_COST_FIELDS], row[2:]))
    conn.close()

    # Overlay the auto-calculated TRIMMING values
    trimming_qty = get_trimming_prod_qty_by_unit_store(month)
    for unit, store, process in get_process_master_unit_stores():
        if (process or "").strip().upper() == "TRIMMING":
            costs = {key: None for key, _ in OTHER_COST_FIELDS}
            costs["trimming_insp"] = trimming_qty.get((unit, store), 0)
            result[(unit, store)] = costs

    return result


def get_part_cost_material_detail(month, unit, store, process, item_code):
    """
    Every material consumed for one specific (Unit, Store, Process, Item
    Code) in the given month - the drill-down behind Part Cost Calc's
    double-click popup. Always computed live (Item Type and price are
    looked up fresh, same as the rest of Part Cost Calc), so a correction
    made afterward via the fill-missing-types tool is reflected immediately,
    even if it changes which bucket a material falls into compared to what's
    already baked into that row's own totals.

    Production-type materials are included too (not hidden) with a blank
    cost, so the numbers shown here visibly explain - not just approximate -
    the row's own RM/Cons/BOP totals.
    """
    item_type_col = f"item_type_{month}"
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    conn = sqlite3.connect(prod_db)
    ensure_year_month_column(conn)
    cur = conn.cursor()
    cur.execute(f"""
        SELECT pr.input_item, MAX(pr.item_description) AS description,
               (SELECT pr2.norm_qty FROM production_register pr2
                WHERE pr2.unit = pr.unit AND pr2.store = pr.store AND pr2.process = pr.process
                  AND pr2.item_code = pr.item_code AND pr2.input_item = pr.input_item
                  AND pr2.year_month = pr.year_month
                ORDER BY substr(pr2.date,7,4) || substr(pr2.date,4,2) || substr(pr2.date,1,2) DESC
                LIMIT 1) AS norm_qty,
               SUM(CAST(pr.input_qty AS REAL)) AS total_consumed,
               MAX(im.{item_type_col}) AS item_type
        FROM production_register pr
        LEFT JOIN item_master im ON pr.input_item = im.erp_code
        WHERE pr.year_month = ? AND pr.unit = ? AND pr.store = ? AND pr.process = ? AND pr.item_code = ?
          AND pr.input_item IS NOT NULL AND TRIM(pr.input_item) != ''
        GROUP BY pr.input_item
        ORDER BY pr.input_item
    """, (month, unit, store, process, item_code))
    material_rows = cur.fetchall()
    conn.close()

    weighted_prices = get_weighted_prices_for_month(month)

    result = []
    for input_item, description, norm_qty, total_consumed, item_type in material_rows:
        total_consumed = total_consumed or 0

        if item_type == "Purchase":
            code_upper = str(input_item).upper()
            if "RMI" in code_upper or "RMP" in code_upper:
                bucket = "RM"
            elif "SFG" in code_upper:
                bucket = "BOP"
            else:
                bucket = "Cons"
            price = weighted_prices.get(input_item) or 0
            cost_contribution = round(total_consumed * price, 2)
        else:
            bucket = None
            price = None
            cost_contribution = None

        result.append({
            "input_item": input_item,
            "description": description,
            "item_type": item_type,
            "norm_qty": norm_qty,
            "total_consumed": round(total_consumed, 4),
            "bucket": bucket,
            "unit_price": round(price, 4) if price is not None else None,
            "cost_contribution": cost_contribution,
        })

    return result


@app.route("/part_cost/detail")
def part_cost_detail():
    month = request.args.get("month")
    unit = request.args.get("unit")
    store = request.args.get("store")
    process = request.args.get("process")
    item_code = request.args.get("item_code")

    if not all([month, unit, store, process, item_code]):
        return jsonify({"error": "Missing required parameters"}), 400

    try:
        materials = get_part_cost_material_detail(month, unit, store, process, item_code)
    except Exception as e:
        # Never let an unexpected failure fall through to Flask's default
        # HTML error page - the frontend does fetch(...).then(res => res.json()),
        # and JSON.parse() on an HTML page fails with a confusing
        # "Unexpected token '<'" that gives no clue what actually went wrong.
        return jsonify({"error": f"Could not load material detail: {e}"}), 500

    return jsonify({"materials": materials, "unit": unit, "store": store, "process": process, "item_code": item_code})


def build_part_cost_calc(month, filters=None, page=1, page_size_override=-1):
    """
    Part Cost Calc: for each (Unit, Store, Process, Item Code) produced this
    month, sums the cost of every PURCHASE-type material consumed (Production-
    type materials are excluded here - their cost shows up at their own stage
    instead), split into three buckets by code pattern:
        RM   - contains "RMI" or "RMP"
        BOP  - contains "SFG" (a purchased semi-finished/child part)
        Cons - everything else purchased
    Cost/Part = each bucket's total ÷ Total Prod Qty.

    Prices come from Weighted Price Master (a separate database file,
    Analytics.db) via ATTACH DATABASE, joined on (item_code, month).
    """
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    item_type_col = f"item_type_{month}"
    filters = filters or {}

    conn = sqlite3.connect(prod_db)
    cur = conn.cursor()
    ensure_year_month_column(conn)
    cur.execute("ATTACH DATABASE ? AS analytics", (ANALYTICS_DB,))

    where_clauses = []
    params = [month, month]  # for the two CTEs' WHERE year_month = ?

    for field, column in PART_COST_FILTER_COLUMNS.items():
        value = filters.get(field, "").strip()
        if value:
            where_clauses.append(f"{column} LIKE ?")
            params.append(f"%{value}%")

    where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""

    # page_size_override=None means "no pagination, return everything" (used
    # by the BOM tree's cost lookups); -1 means "use the normal page size".
    if page_size_override is None:
        page_size = -1     # SQLite treats LIMIT -1 as no limit
        offset = 0
    else:
        page_size = PART_COST_PAGE_SIZE if page_size_override == -1 else page_size_override
        offset = (page - 1) * page_size

    # Order must match the ?'s as they appear in the query text: qty's month,
    # materials' month, priced's wp.month join condition, THEN the filter
    # values, THEN LIMIT/OFFSET at the very end.
    final_params = [month, month, month] + params[2:] + [page_size, offset]

    query = f"""
        WITH qty AS (
            SELECT unit, store, process, item_code,
                   MAX(company) AS company, MAX(item_name) AS item_name, MAX(uom) AS uom,
                   SUM(CAST(acc_qty AS REAL)) AS total_acc_qty,
                   SUM(CAST(prod_qty AS REAL)) AS total_prod_qty
            FROM (
                SELECT DISTINCT source_entry_id, company, unit, store, process, item_code, item_name, uom, acc_qty, prod_qty
                FROM production_register
                WHERE year_month = ?
                  AND item_code IS NOT NULL AND TRIM(item_code) != ''
                  AND unit IS NOT NULL AND TRIM(unit) != ''
                  AND store IS NOT NULL AND TRIM(store) != ''
                  AND process IS NOT NULL AND TRIM(process) != ''
            )
            GROUP BY unit, store, process, item_code
        ),
        materials AS (
            SELECT pr.unit, pr.store, pr.process, pr.item_code, pr.input_item,
                   SUM(CAST(pr.input_qty AS REAL)) AS total_consumed,
                   MAX(im.{item_type_col}) AS item_type
            FROM production_register pr
            LEFT JOIN item_master im ON pr.input_item = im.erp_code
            WHERE pr.year_month = ? AND pr.input_item IS NOT NULL AND TRIM(pr.input_item) != ''
              AND pr.item_code IS NOT NULL AND TRIM(pr.item_code) != ''
              AND pr.unit IS NOT NULL AND TRIM(pr.unit) != ''
              AND pr.store IS NOT NULL AND TRIM(pr.store) != ''
              AND pr.process IS NOT NULL AND TRIM(pr.process) != ''
            GROUP BY pr.unit, pr.store, pr.process, pr.item_code, pr.input_item
            HAVING item_type = 'Purchase'
        ),
        priced AS (
            SELECT m.unit, m.store, m.process, m.item_code,
                   m.total_consumed * COALESCE(wp.final_price, 0) AS cost,
                   CASE
                       WHEN m.input_item LIKE '%RMI%' OR m.input_item LIKE '%RMP%' THEN 'rm'
                       WHEN m.input_item LIKE '%SFG%' THEN 'bop'
                       ELSE 'cons'
                   END AS bucket
            FROM materials m
            LEFT JOIN analytics.weighted_price_master wp ON wp.item_code = m.input_item AND wp.month = ?
        ),
        final AS (
            SELECT qty.company, qty.unit, qty.store, qty.process, qty.item_code, qty.item_name, qty.uom,
                   qty.total_acc_qty, qty.total_prod_qty,
                   COALESCE(SUM(CASE WHEN priced.bucket = 'rm' THEN priced.cost END), 0) AS rm_cost,
                   COALESCE(SUM(CASE WHEN priced.bucket = 'cons' THEN priced.cost END), 0) AS cons_cost,
                   COALESCE(SUM(CASE WHEN priced.bucket = 'bop' THEN priced.cost END), 0) AS bop_cost
            FROM qty
            LEFT JOIN priced ON priced.unit = qty.unit AND priced.store = qty.store
                            AND priced.process = qty.process AND priced.item_code = qty.item_code
            GROUP BY qty.unit, qty.store, qty.process, qty.item_code
        )
        SELECT final.*, COUNT(*) OVER() AS total_count
        FROM final
        {where_sql}
        ORDER BY unit, store, item_code
        LIMIT ? OFFSET ?
    """

    cur.execute(query, final_params)
    results = cur.fetchall()
    conn.close()

    rows = []
    total_count = results[0][-1] if results else 0

    # Overhead allocation inputs - computed once, then applied per row
    store_totals, part_times = get_store_time_totals(month)
    other_costs = get_other_costs_for_month(month)
    overhead_keys = [key for key, _ in OTHER_COST_FIELDS]

    for r in results:
        company, unit, store, process, item_code, item_name, uom, acc_qty, prod_qty, rm_cost, cons_cost, bop_cost, _total = r
        prod_qty = prod_qty or 0

        # This part's share of its store's total machine/bay time that month
        part_time = part_times.get((unit, store, process, item_code), 0)
        store_total_time = store_totals.get((unit, store), 0)
        utilization = (part_time / store_total_time) if store_total_time else 0

        store_costs = other_costs.get((unit, store), {})
        overhead = {}
        for key in overhead_keys:
            store_amount = store_costs.get(key) or 0
            allocated = store_amount * utilization
            overhead[key] = round(allocated, 2)
            overhead[f"{key}_per_part"] = round(allocated / prod_qty, 4) if prod_qty else 0

        row = {
            "company": company, "unit": unit, "store": store, "process": process, "item_code": item_code,
            "item_name": item_name, "uom": uom,
            "acc_qty": round(acc_qty or 0, 4), "prod_qty": round(prod_qty, 4),
            "rm_cost": round(rm_cost, 2), "cons_cost": round(cons_cost, 2), "bop_cost": round(bop_cost, 2),
            "rm_cost_per_part": round(rm_cost / prod_qty, 4) if prod_qty else 0,
            "cons_cost_per_part": round(cons_cost / prod_qty, 4) if prod_qty else 0,
            "bop_cost_per_part": round(bop_cost / prod_qty, 4) if prod_qty else 0,
            "utilization_pct": round(utilization * 100, 4),
        }
        row.update(overhead)

        # Grand total per part: the three material buckets plus every
        # allocated overhead category
        row["total_cost_per_part"] = round(
            row["rm_cost_per_part"] + row["cons_cost_per_part"] + row["bop_cost_per_part"]
            + sum(overhead[f"{key}_per_part"] for key in overhead_keys),
            4,
        )

        rows.append(row)

    return rows, total_count


@app.route("/part_cost")
def part_cost():
    months = get_all_relevant_months()
    selected_month = request.args.get("month", months[0] if months else None)

    if not selected_month:
        return render_analytics_page("part_cost.html", "part_cost",
                                      months_override=months, item_master_ready=False,
                                      rows=[], missing_item_types=[], page=1, total_pages=1,
                                      total_count=0, filters={}, page_size=PART_COST_PAGE_SIZE,
                                      overhead_fields=OTHER_COST_FIELDS)

    item_master_ready = get_item_type_month_column_exists(selected_month)

    filters = {field: request.args.get(field, "") for field in PART_COST_FILTER_COLUMNS}
    page = max(1, request.args.get("page", 1, type=int))

    rows = []
    missing_item_types = []
    total_count = 0
    total_pages = 1

    if item_master_ready:
        rows, total_count = build_part_cost_calc(selected_month, filters=filters, page=page)
        total_pages = max(1, -(-total_count // PART_COST_PAGE_SIZE))
        missing_item_types = [code for code, _desc in get_missing_item_type_codes(selected_month)]

    return render_analytics_page(
        "part_cost.html", "part_cost",
        months_override=months,
        item_master_ready=item_master_ready,
        rows=rows,
        missing_item_types=missing_item_types,
        page=page,
        total_pages=total_pages,
        total_count=total_count,
        filters=filters,
        page_size=PART_COST_PAGE_SIZE,
        overhead_fields=OTHER_COST_FIELDS,
    )


PART_MARGIN_PAGE_SIZE = 200


def build_part_margin_table(month, item_code_filter="", customer_filter="", unit_filter="",
                              include_missing=True, sort_by="item_code", sort_dir="asc"):
    """
    Margin for a single month, one row per FG actually sold that month.
    Customer and Unit are reference fields only - the customer and unit
    from that FG's single MOST RECENT invoice that month (from Sales
    Register; Unit here is the sales-side unit the invoice was made from,
    not the production unit BOM Tree associates with the FG). An FG can
    legitimately be sold to several customers from several units within
    one month; these two fields aren't an attempt to represent all of
    that, just a quick "who/where was this last sold to" reference
    alongside the aggregate weighted price. customer_filter/unit_filter
    filter on these same reference fields.

    Cost is the per-FG-per-month figure from BOM Tree (production-side,
    tied to wherever the part was actually made).

    Margin is for the FULL quantity sold that month, not per unit.
    """
    fg_costs = get_all_fg_total_costs(month)

    conn = sqlite3.connect(ANALYTICS_DB)
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT item_code, customer, unit, qty_sold, final_price "
            "FROM weighted_selling_price_master WHERE month = ? AND qty_sold IS NOT NULL AND qty_sold > 0",
            (month,),
        )
        selling_data = cur.fetchall()
    except sqlite3.OperationalError:
        selling_data = []
    conn.close()

    rows = []
    for code, customer, unit, qty_sold, price in selling_data:
        cost_info = fg_costs.get(code)
        cost_missing = cost_info["is_missing"] if cost_info else True
        # Fundamentally unknown (no cost_info at all, or the root's own cost
        # was never found) means "price NA" - the 0-vs-unknown distinction
        # matters a lot here, since treating an unknown cost as a real zero
        # would silently invent a margin that isn't real.
        cost_unknown = (cost_info is None) or cost_info["root_cost_unknown"]
        cost = None if cost_unknown else cost_info["total_cost"]

        if cost is not None and price is not None:
            margin_per_unit = price - cost
            margin_total = round(margin_per_unit * qty_sold, 2)
            margin_pct = round((margin_per_unit / cost * 100), 2) if cost else None
        else:
            margin_total = None
            margin_pct = None

        rows.append({
            "item_code": code,
            "customer": customer,
            "unit": unit,
            "qty_sold": round(qty_sold, 4),
            "selling_price": round(price, 4) if price is not None else None,
            "cost": round(cost, 4) if cost is not None else None,
            "cost_missing": cost_missing,
            "margin": margin_total,
            "margin_pct": margin_pct,
        })

    if item_code_filter.strip():
        needle = item_code_filter.strip().lower()
        rows = [r for r in rows if needle in r["item_code"].lower()]

    if customer_filter.strip():
        needle = customer_filter.strip().lower()
        rows = [r for r in rows if needle in (r["customer"] or "").lower()]

    if unit_filter.strip():
        needle = unit_filter.strip().lower()
        rows = [r for r in rows if needle in (r["unit"] or "").lower()]

    if not include_missing:
        rows = [r for r in rows if not r["cost_missing"]]

    sort_key_map = {
        "item_code": lambda r: r["item_code"],
        "customer": lambda r: (r["customer"] or ""),
        "unit": lambda r: (r["unit"] or ""),
        "qty_sold": lambda r: (r["qty_sold"] is None, r["qty_sold"] or 0),
        "selling_price": lambda r: (r["selling_price"] is None, r["selling_price"] or 0),
        "cost": lambda r: (r["cost"] is None, r["cost"] or 0),
        "margin": lambda r: (r["margin"] is None, r["margin"] or 0),
        "margin_pct": lambda r: (r["margin_pct"] is None, r["margin_pct"] or 0),
    }
    key_fn = sort_key_map.get(sort_by, sort_key_map["item_code"])
    rows.sort(key=key_fn, reverse=(sort_dir == "desc"))

    return rows


def weighted_selling_price_has_run(month):
    """Whether Weighted Selling Price has any data at all for this month -
    distinguishes 'never run yet' from 'ran but this FG has no sales', and
    also catches an OLDER schema (before the FG-Customer-Unit restructure)
    that still exists on disk but is missing the customer/unit columns -
    without this check, a table left over from before that change would
    report count > 0 here, but build_part_margin_table's query against it
    would silently return zero rows, leaving the page confusingly blank
    with no explanation instead of the clear 'please re-run' message."""
    if not os.path.exists(ANALYTICS_DB):
        return False
    conn = sqlite3.connect(ANALYTICS_DB)
    cur = conn.cursor()
    try:
        cur.execute("PRAGMA table_info(weighted_selling_price_master)")
        columns = {row[1] for row in cur.fetchall()}
        if "customer" not in columns or "unit" not in columns:
            conn.close()
            return False
        cur.execute("SELECT COUNT(*) FROM weighted_selling_price_master WHERE month = ?", (month,))
        count = cur.fetchone()[0]
    except sqlite3.OperationalError:
        count = 0
    conn.close()
    return count > 0


@app.route("/part_margin/detail")
def part_margin_detail():
    """
    Full summary for one FG in Part-wise Margin's double-click popup: its
    sales figures (weighted selling price, qty sold, and the customer/unit
    of its most recent invoice that month), cost breakdown (root cost plus
    every sub-item in its BOM tree, each with its own individual cost), and
    margin - computed the SAME way build_part_margin_table computes the row
    itself, so the popup's numbers always agree with the row that was
    clicked.
    """
    month = request.args.get("month")
    item_code = request.args.get("item_code")

    if not month or not item_code:
        return jsonify({"error": "Missing required parameters"}), 400

    try:
        conn = sqlite3.connect(ANALYTICS_DB)
        cur = conn.cursor()

        cur.execute(
            "SELECT qty_sold, final_price, customer, unit FROM weighted_selling_price_master "
            "WHERE item_code = ? AND month = ?",
            (item_code, month),
        )
        sales_row = cur.fetchone()

        cur.execute("SELECT 1 FROM bom_tree WHERE month = ? AND root_fg = ? LIMIT 1", (month, item_code))
        has_bom_tree = cur.fetchone() is not None
        conn.close()

        if sales_row is None:
            return jsonify({"error": f"{item_code} was not sold in this month"}), 404

        qty_sold, price, customer, unit = sales_row
        qty_sold = qty_sold or 0

        if has_bom_tree:
            detail = get_single_fg_tree_detail(item_code, month)
            cost = None if detail["root_cost_unknown"] else detail["total_cost"]
            cost_source = detail["root_source"]
            sub_items = detail["sub_items"]
            root_row = {
                "code": item_code,
                "item": get_item_name(item_code),
                "item_type": detail["root_type"],
                "process": detail["root_cost_process"],
                "unit": detail["root_cost_unit"],
                "norm_qty": 1,
                "unit_cost": detail["root_cost"],
                "extended_cost": detail["root_cost"],
                "is_missing": detail["root_cost"] is None,
            }
        else:
            cost = None
            cost_source = "No BOM Tree entry for this FG this month"
            sub_items = []
            root_row = None

        if cost is not None and price is not None:
            margin_per_unit = price - cost
            margin_total = round(margin_per_unit * qty_sold, 2)
            margin_pct = round((margin_per_unit / cost * 100), 2) if cost else None
        else:
            margin_total = None
            margin_pct = None

        return jsonify({
            "item_code": item_code,
            "customer": customer,
            "unit": unit,
            "month": month,
            "qty_sold": round(qty_sold, 4),
            "selling_price": round(price, 4) if price is not None else None,
            "cost": round(cost, 4) if cost is not None else None,
            "cost_source": cost_source,
            "margin": margin_total,
            "margin_pct": margin_pct,
            "root_row": root_row,
            "sub_items": sub_items,
        })
    except Exception as e:
        return jsonify({"error": f"Could not load detail: {e}"}), 500


@app.route("/part_margin")
def part_margin():
    months = get_all_relevant_months()
    selected_month = request.args.get("month", months[0] if months else None)
    item_code_filter = request.args.get("item_code", "")
    customer_filter = request.args.get("customer", "")
    unit_filter = request.args.get("unit", "")
    page = max(1, request.args.get("page", 1, type=int))
    sort_by = request.args.get("sort", "item_code")
    sort_dir = request.args.get("dir", "asc")

    # Checkboxes only submit a value when CHECKED - detect a real form
    # submission (item_code is always present once the form posts, blank or
    # not) so unchecking the box is correctly read as False, not silently
    # ignored (this bit us on Overview's identical checkbox earlier).
    form_submitted = "item_code" in request.args
    include_missing = (request.args.get("include_missing") == "1") if form_submitted else True

    selling_price_ready = weighted_selling_price_has_run(selected_month) if selected_month else False

    all_rows = build_part_margin_table(
        selected_month, item_code_filter, customer_filter, unit_filter,
        include_missing=include_missing, sort_by=sort_by, sort_dir=sort_dir,
    ) if (selected_month and selling_price_ready) else []
    total_count = len(all_rows)
    total_pages = max(1, -(-total_count // PART_MARGIN_PAGE_SIZE))
    start = (page - 1) * PART_MARGIN_PAGE_SIZE
    rows = all_rows[start:start + PART_MARGIN_PAGE_SIZE]

    return render_analytics_page(
        "part_margin.html", "part_margin",
        months_override=months,
        selling_price_ready=selling_price_ready,
        rows=rows,
        total_count=total_count,
        total_pages=total_pages,
        page=page,
        item_code_filter=item_code_filter,
        customer_filter=customer_filter,
        unit_filter=unit_filter,
        include_missing=include_missing,
        sort_by=sort_by,
        sort_dir=sort_dir,
    )


def get_item_type_month_column_exists(month):
    """Checks whether Item Master has EVER been run for this month at all
    (i.e. the item_type_<month> column exists), as opposed to it existing
    but a specific item just not being found in it."""
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    if not os.path.exists(prod_db):
        return False
    conn = sqlite3.connect(prod_db)
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(item_master)")
    columns = {row[1] for row in cur.fetchall()}
    conn.close()
    return f"item_type_{month}" in columns


def ensure_sales_register_year_month(conn):
    """Self-heals older Sales Register databases that predate the
    year_month column/index - mirrors the exact fix already proven on
    production_register. Without this, Overview's date-range queries do a
    full table scan, computing a string concatenation on-the-fly for every
    single row just to compare dates - which gets very slow once Sales
    Register has thousands of rows (it easily can: ~9,000 rows per unit
    per month)."""
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(sales_register)")
    columns = {row[1] for row in cur.fetchall()}

    if "year_month" not in columns:
        cur.execute("ALTER TABLE sales_register ADD COLUMN year_month TEXT")
        conn.commit()

    cur.execute("SELECT COUNT(*) FROM sales_register WHERE year_month IS NULL AND invoice_date IS NOT NULL")
    if cur.fetchone()[0] > 0:
        cur.execute("""
            UPDATE sales_register
            SET year_month = substr(invoice_date,7,4) || '_' || substr(invoice_date,4,2)
            WHERE year_month IS NULL AND invoice_date IS NOT NULL
        """)
        conn.commit()

    cur.execute("CREATE INDEX IF NOT EXISTS idx_sales_year_month ON sales_register(year_month, unit)")
    conn.commit()


def ensure_year_month_column(conn):
    """Self-heals older databases that predate the year_month column/index -
    adds it and backfills it automatically, so you don't have to remember
    to re-run Production_Entry.py just for this."""
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(production_register)")
    columns = {row[1] for row in cur.fetchall()}

    if "year_month" not in columns:
        cur.execute("ALTER TABLE production_register ADD COLUMN year_month TEXT")
        conn.commit()

    cur.execute("SELECT COUNT(*) FROM production_register WHERE year_month IS NULL AND date IS NOT NULL")
    if cur.fetchone()[0] > 0:
        cur.execute("""
            UPDATE production_register
            SET year_month = substr(date,7,4) || '_' || substr(date,4,2)
            WHERE year_month IS NULL AND date IS NOT NULL
        """)
        conn.commit()

    cur.execute("CREATE INDEX IF NOT EXISTS idx_year_month ON production_register(year_month)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_year_month_group ON production_register(year_month, unit, store, process, item_code)")
    conn.commit()


MONTH_POSTING_PAGE_SIZE = 200

# Which template filter field maps to which underlying SQL column
MONTH_POSTING_FILTER_COLUMNS = {
    "company": "final.company",
    "unit": "final.unit",
    "store": "final.store",
    "process": "final.process",
    "item_code": "final.item_code",
    "item_name": "final.item_name",
    "uom": "final.uom",
    "input_item": "final.input_item",
    "item_description": "final.item_description",
    "item_type": "final.item_type",
}


def build_month_posting(month, filters=None, page=1):
    """
    Pivots production_register into a clean monthly summary: one row per
    (Unit, Store, Process, Item Code, Input Item), with quantities properly
    de-duplicated (Acc/Prod Qty are repeated on every input-item row of the
    same entry in the raw data, so we collapse to distinct entries FIRST,
    then sum - otherwise multi-ingredient entries would inflate the totals).

    Entirely SQL-side (via CTEs) so it stays fast even with huge datasets:
    filtering, grouping, AND pagination all happen in the database, so we
    only ever pull one page's worth of rows into Python/the browser - never
    the whole month's data at once (which is what was making the page slow).

    Item Type is looked up LIVE against item_master's month-specific column,
    not the value baked into production_register at extraction time - so
    running Item Master after Production Entry still works correctly.
    """
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    item_type_col = f"item_type_{month}"
    filters = filters or {}

    conn = sqlite3.connect(prod_db)
    cur = conn.cursor()
    ensure_year_month_column(conn)

    where_clauses = []
    params = [month, month]  # for the two CTEs' WHERE year_month = ?

    for field, column in MONTH_POSTING_FILTER_COLUMNS.items():
        value = filters.get(field, "").strip()
        if value:
            where_clauses.append(f"{column} LIKE ?")
            params.append(f"%{value}%")

    where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""

    offset = (page - 1) * MONTH_POSTING_PAGE_SIZE
    params_with_paging = params + [MONTH_POSTING_PAGE_SIZE, offset]

    query = f"""
        WITH qty AS (
            SELECT unit, store, process, item_code,
                   MAX(company) AS company, MAX(item_name) AS item_name, MAX(uom) AS uom,
                   SUM(CAST(acc_qty AS REAL)) AS total_acc_qty,
                   SUM(CAST(prod_qty AS REAL)) AS total_prod_qty
            FROM (
                SELECT DISTINCT source_entry_id, company, unit, store, process, item_code, item_name, uom, acc_qty, prod_qty
                FROM production_register
                WHERE year_month = ?
                  AND item_code IS NOT NULL AND TRIM(item_code) != ''
                  AND unit IS NOT NULL AND TRIM(unit) != ''
                  AND store IS NOT NULL AND TRIM(store) != ''
                  AND process IS NOT NULL AND TRIM(process) != ''
            )
            GROUP BY unit, store, process, item_code
        ),
        materials AS (
            SELECT pr.unit, pr.store, pr.process, pr.item_code, pr.input_item,
                   MAX(pr.item_description) AS item_description,
                   (SELECT pr2.norm_qty FROM production_register pr2
                    WHERE pr2.unit = pr.unit AND pr2.store = pr.store AND pr2.process = pr.process
                      AND pr2.item_code = pr.item_code AND pr2.input_item = pr.input_item
                      AND pr2.year_month = pr.year_month
                    ORDER BY substr(pr2.date,7,4) || substr(pr2.date,4,2) || substr(pr2.date,1,2) DESC
                    LIMIT 1) AS norm_qty,
                   SUM(CAST(pr.input_qty AS REAL)) AS total_consumed,
                   MAX(im.{item_type_col}) AS item_type
            FROM production_register pr
            LEFT JOIN item_master im ON pr.input_item = im.erp_code
            WHERE pr.year_month = ? AND pr.input_item IS NOT NULL AND TRIM(pr.input_item) != ''
              AND pr.item_code IS NOT NULL AND TRIM(pr.item_code) != ''
              AND pr.unit IS NOT NULL AND TRIM(pr.unit) != ''
              AND pr.store IS NOT NULL AND TRIM(pr.store) != ''
              AND pr.process IS NOT NULL AND TRIM(pr.process) != ''
            GROUP BY pr.unit, pr.store, pr.process, pr.item_code, pr.input_item
        ),
        final AS (
            SELECT qty.company, materials.unit, materials.store, materials.process, materials.item_code,
                   qty.item_name, qty.uom, qty.total_acc_qty, qty.total_prod_qty,
                   materials.input_item, materials.item_description, materials.item_type,
                   materials.norm_qty, materials.total_consumed
            FROM materials
            JOIN qty ON materials.unit = qty.unit AND materials.store = qty.store
                    AND materials.process = qty.process AND materials.item_code = qty.item_code
        )
        SELECT final.*, COUNT(*) OVER() AS total_count
        FROM final
        {where_sql}
        ORDER BY unit, store, item_code, input_item
        LIMIT ? OFFSET ?
    """

    cur.execute(query, params_with_paging)
    results = cur.fetchall()
    conn.close()

    rows = []
    total_count = results[0][-1] if results else 0

    for r in results:
        rows.append({
            "company": r[0], "unit": r[1], "store": r[2], "process": r[3], "item_code": r[4],
            "item_name": r[5], "uom": r[6],
            "total_acc_qty": round(r[7] or 0, 4), "total_prod_qty": round(r[8] or 0, 4),
            "input_item": r[9], "item_description": r[10], "item_type": r[11],
            "norm_qty": r[12], "total_consumed": round(r[13] or 0, 4),
        })

    return rows, total_count


def get_missing_item_type_codes(month):
    """Lightweight version of the missing-item-type check (used by the
    download-missing-types button) - doesn't materialize the full posting
    table, just the distinct input items with no Item Type this month."""
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    item_type_col = f"item_type_{month}"

    conn = sqlite3.connect(prod_db)
    cur = conn.cursor()
    ensure_year_month_column(conn)
    cur.execute(f"""
        SELECT DISTINCT pr.input_item, pr.item_description
        FROM production_register pr
        LEFT JOIN item_master im ON pr.input_item = im.erp_code
        WHERE pr.year_month = ? AND pr.input_item IS NOT NULL AND TRIM(pr.input_item) != ''
          AND im.{item_type_col} IS NULL
        ORDER BY pr.input_item
    """, (month,))
    result = cur.fetchall()
    conn.close()
    return result


@app.route("/month_posting/download_missing_types")
def month_posting_download_missing_types():
    month = request.args.get("month")
    missing = get_missing_item_type_codes(month)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Item Code", "Description", "Item Type"])
    for item_code, description in missing:
        ws.append([item_code, description, None])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"item_type_missing_{month}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/month_posting/upload_missing_types", methods=["POST"])
def month_posting_upload_missing_types():
    month = request.form.get("month")
    if not month:
        return jsonify({"error": "Missing month"}), 400

    if "file" not in request.files:
        return jsonify({"error": "No file was sent"}), 400

    uploaded_file = request.files["file"]
    if uploaded_file.filename == "":
        return jsonify({"error": "No file was selected"}), 400

    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        ws = wb.active
    except Exception as e:
        return jsonify({"error": f"Could not read the file: {e}"}), 400

    rows = list(ws.iter_rows(min_row=2, values_only=True))  # skip header row
    item_type_col = f"item_type_{month}"

    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    conn = sqlite3.connect(prod_db)
    cur = conn.cursor()

    # Make sure this month's column exists before we try to write into it
    cur.execute("PRAGMA table_info(item_master)")
    existing_cols = {row[1] for row in cur.fetchall()}
    if item_type_col not in existing_cols:
        cur.execute(f"ALTER TABLE item_master ADD COLUMN {item_type_col} TEXT")
        existing_cols.add(item_type_col)

    # Every item_type_* month column that already exists - a filled-in value
    # applies to ALL of them wherever they're currently blank, not just the
    # month you happened to be viewing. This is what a "missing entirely from
    # Item Master" code needs: one answer that closes the gap everywhere,
    # while never touching a month that already has its own real value from
    # an official Item Master import.
    all_month_cols = sorted(c for c in existing_cols if c.startswith("item_type_"))

    updated = 0
    skipped_blank = 0

    for row in rows:
        if not row or len(row) < 3:
            continue
        item_code, description, item_type = row[0], row[1], row[2]

        if not item_code:
            continue
        if item_type is None or str(item_type).strip() == "":
            skipped_blank += 1
            continue

        item_type = str(item_type).strip()

        # Make sure a row exists for this code at all (a code that was never
        # in Item Master's CSV export won't have one yet)
        cur.execute("INSERT OR IGNORE INTO item_master (erp_code) VALUES (?)", (item_code,))

        set_clause = ", ".join(f"{c} = COALESCE({c}, ?)" for c in all_month_cols)
        cur.execute(
            f"UPDATE item_master SET {set_clause} WHERE erp_code = ?",
            [item_type] * len(all_month_cols) + [item_code],
        )
        updated += 1

    conn.commit()
    conn.close()

    invalidate_cost_cache()
    return jsonify({"updated": updated, "skipped_blank": skipped_blank})


@app.route("/month_posting")
def month_posting():
    months = get_all_relevant_months()
    selected_month = request.args.get("month", months[0] if months else None)

    if not selected_month:
        return render_analytics_page("month_posting.html", "month_posting",
                                      months_override=months, item_master_ready=False,
                                      rows=[], missing_item_types=[], page=1, total_pages=1,
                                      total_count=0, filters={}, page_size=MONTH_POSTING_PAGE_SIZE)

    item_master_ready = get_item_type_month_column_exists(selected_month)

    filters = {field: request.args.get(field, "") for field in MONTH_POSTING_FILTER_COLUMNS}
    page = max(1, request.args.get("page", 1, type=int))

    rows = []
    missing_item_types = []
    total_count = 0
    total_pages = 1

    if item_master_ready:
        rows, total_count = build_month_posting(selected_month, filters=filters, page=page)
        total_pages = max(1, -(-total_count // MONTH_POSTING_PAGE_SIZE))  # ceiling division
        missing_item_types = [code for code, _desc in get_missing_item_type_codes(selected_month)]

    return render_analytics_page(
        "month_posting.html", "month_posting",
        months_override=months,
        item_master_ready=item_master_ready,
        rows=rows,
        missing_item_types=missing_item_types,
        page=page,
        total_pages=total_pages,
        total_count=total_count,
        filters=filters,
        page_size=MONTH_POSTING_PAGE_SIZE,
    )


@app.route("/cycle_time")
def cycle_time():
    months = get_all_relevant_months()
    selected_month = request.args.get("month", months[0] if months else None)
    scope = request.args.get("scope", "all")   # "all" or "missing"

    all_combos = get_all_cycle_time_combos(selected_month) if selected_month else []
    missing_count = sum(1 for c in all_combos if c["is_missing"])

    rows = [c for c in all_combos if c["is_missing"]] if scope == "missing" else all_combos

    return render_analytics_page(
        "cycle_time.html", "cycle_time",
        months_override=months,
        missing_count=missing_count,
        total_count=len(all_combos),
        rows=rows,
        scope=scope,
    )


@app.route("/cycle_time/update", methods=["POST"])
def cycle_time_update():
    """Saves Cycle Time values edited directly in the table. Works for both
    brand-new combos (not captured yet) and corrections to existing ones -
    the value entered here becomes what's used from now on. Fixed-cycle-time
    processes (ASM/BUFFING/TRIMMING/PVD) are always 1s and can't be changed."""
    body = request.get_json(silent=True) or {}
    updates = body.get("updates", [])

    if not updates:
        return jsonify({"error": "No updates provided"}), 400

    conn = get_cycle_time_connection()
    cur = conn.cursor()

    saved = 0
    skipped_blank = 0
    skipped_fixed = 0

    for u in updates:
        unit, store, process, item_code = u.get("unit"), u.get("store"), u.get("process"), u.get("item_code")
        cycle_time = u.get("cycle_time")

        if not all([unit, store, process, item_code]):
            continue

        if str(process).strip().upper() in FIXED_CYCLE_TIME_PROCESSES:
            skipped_fixed += 1
            continue

        if cycle_time is None or str(cycle_time).strip() == "":
            skipped_blank += 1
            continue

        try:
            cycle_time = float(cycle_time)
        except (TypeError, ValueError):
            skipped_blank += 1
            continue

        cur.execute("""
            INSERT INTO cycle_time_master (unit, store, process, item_code, cycle_time)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(unit, store, process, item_code) DO UPDATE SET cycle_time = excluded.cycle_time
        """, (unit, store, process, item_code, cycle_time))
        saved += 1

    conn.commit()
    conn.close()

    invalidate_cost_cache()
    return jsonify({"saved": saved, "skipped_blank": skipped_blank, "skipped_fixed": skipped_fixed})


def get_cycle_time_connection():
    conn = sqlite3.connect(ANALYTICS_DB)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS cycle_time_master (
            unit TEXT,
            store TEXT,
            process TEXT,
            item_code TEXT,
            cycle_time REAL,
            PRIMARY KEY (unit, store, process, item_code)
        )
    """)
    conn.commit()
    return conn


# Processes where cycle time is fixed at 1 second by definition - never
# editable, never uploadable, never counted as "missing".
FIXED_CYCLE_TIME_PROCESSES = {"ASM", "BUFFING", "TRIMMING", "PVD"}
FIXED_CYCLE_TIME_VALUE = 1.0


def get_all_cycle_time_combos(month):
    """
    Every distinct (Unit, Store, Process, Item Code) actually produced in
    the given month, with its cycle time if one exists (None if not yet
    captured), the item's description, and its total Prod Qty for context.

    Processes in FIXED_CYCLE_TIME_PROCESSES always report 1.0 automatically
    and are flagged is_fixed=True, regardless of what's in the table.
    """
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    if not os.path.exists(prod_db) or not month:
        return []

    conn = sqlite3.connect(prod_db)
    ensure_year_month_column(conn)
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT unit, store, process, item_code,
                   MAX(item_name) AS item_name,
                   SUM(CAST(prod_qty AS REAL)) AS total_prod_qty
            FROM (
                SELECT DISTINCT source_entry_id, unit, store, process, item_code, item_name, prod_qty
                FROM production_register
                WHERE year_month = ?
                  AND unit IS NOT NULL AND TRIM(unit) != ''
                  AND store IS NOT NULL AND TRIM(store) != ''
                  AND process IS NOT NULL AND TRIM(process) != ''
                  AND item_code IS NOT NULL AND TRIM(item_code) != ''
            )
            GROUP BY unit, store, process, item_code
            ORDER BY unit, store, process, item_code
        """, (month,))
        produced = cur.fetchall()
    except sqlite3.OperationalError:
        produced = []
    conn.close()

    conn2 = get_cycle_time_connection()
    cur2 = conn2.cursor()
    cur2.execute("SELECT unit, store, process, item_code, cycle_time FROM cycle_time_master")
    stored = {(u, s, p, i): ct for u, s, p, i, ct in cur2.fetchall()}
    conn2.close()

    result = []
    for unit, store, process, item_code, item_name, total_prod_qty in produced:
        is_fixed = (process or "").strip().upper() in FIXED_CYCLE_TIME_PROCESSES
        cycle_time = FIXED_CYCLE_TIME_VALUE if is_fixed else stored.get((unit, store, process, item_code))
        result.append({
            "unit": unit, "store": store, "process": process, "item_code": item_code,
            "item_name": item_name,
            "total_prod_qty": round(total_prod_qty or 0, 4),
            "cycle_time": cycle_time,
            "is_fixed": is_fixed,
            "is_missing": (not is_fixed) and cycle_time is None,
        })
    return result


def get_missing_cycle_time_combos(month):
    """Just the combos still needing a cycle time (fixed-process ones are
    never considered missing)."""
    return [c for c in get_all_cycle_time_combos(month) if c["is_missing"]]


@app.route("/cycle_time/download_missing")
def cycle_time_download_missing():
    month = request.args.get("month")
    scope = request.args.get("scope", "missing")   # "all" or "missing"

    combos = get_all_cycle_time_combos(month)
    if scope == "missing":
        combos = [c for c in combos if c["is_missing"]]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Unit", "Store", "Process", "Item Code", "Item Description", "Total Prod Qty (this month)", "Cycle Time"])
    for c in combos:
        ws.append([c["unit"], c["store"], c["process"], c["item_code"],
                   c["item_name"], c["total_prod_qty"], c["cycle_time"]])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"cycle_time_{scope}_{month}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/cycle_time/upload", methods=["POST"])
def cycle_time_upload():
    if "file" not in request.files:
        return jsonify({"error": "No file was sent"}), 400

    uploaded_file = request.files["file"]
    if uploaded_file.filename == "":
        return jsonify({"error": "No file was selected"}), 400

    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        ws = wb.active
    except Exception as e:
        return jsonify({"error": f"Could not read the file: {e}"}), 400

    rows = list(ws.iter_rows(min_row=2, values_only=True))  # skip header row

    conn = get_cycle_time_connection()
    cur = conn.cursor()

    saved = 0
    skipped_blank = 0
    skipped_fixed = 0

    for row in rows:
        if not row or len(row) < 7:
            continue
        unit, store, process, item_code = row[0], row[1], row[2], row[3]
        cycle_time = row[6]   # Item Description (4) and Total Prod Qty (5) are just context columns

        if not unit or not store or not process or not item_code:
            continue

        # Fixed-cycle-time processes are always 1s and can't be changed by upload
        if str(process).strip().upper() in FIXED_CYCLE_TIME_PROCESSES:
            skipped_fixed += 1
            continue

        if cycle_time is None or str(cycle_time).strip() == "":
            skipped_blank += 1
            continue

        try:
            cycle_time = float(cycle_time)
        except (TypeError, ValueError):
            skipped_blank += 1
            continue

        # Upload now OVERWRITES an existing value rather than skipping it
        cur.execute(
            "INSERT INTO cycle_time_master (unit, store, process, item_code, cycle_time) VALUES (?, ?, ?, ?, ?) "
            "ON CONFLICT(unit, store, process, item_code) DO UPDATE SET cycle_time = excluded.cycle_time",
            (unit, store, process, item_code, cycle_time),
        )
        saved += 1

    conn.commit()
    conn.close()

    invalidate_cost_cache()
    return jsonify({"saved": saved, "skipped_blank": skipped_blank, "skipped_fixed": skipped_fixed})


OTHER_COST_FIELDS = [
    ("eb_cost", "EB Cost"),
    ("machine_cost_amortized", "Machine Cost (amortized)"),
    ("space_rent", "Space Rent"),
    ("manpower", "Manpower"),
    ("consumables", "Consumables"),
    ("stationery", "Stationery"),
    ("trimming_insp", "Trimming/Insp."),
    ("jig_booth_cleaning_cost", "Jig & Booth Cleaning"),
    ("mold_part_cost", "Mold Part Cost"),
]


def get_other_cost_connection():
    conn = sqlite3.connect(ANALYTICS_DB)
    cols_sql = ", ".join(f"{key} REAL" for key, _ in OTHER_COST_FIELDS)
    col_names = [key for key, _ in OTHER_COST_FIELDS]
    col_list = ", ".join(col_names)

    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='other_cost_master'")
    exists = cur.fetchone()

    if not exists:
        conn.execute(f"""
            CREATE TABLE other_cost_master (
                unit TEXT,
                store TEXT,
                month TEXT,
                {cols_sql},
                PRIMARY KEY (unit, store, month)
            )
        """)
        conn.commit()
        return conn

    cur.execute("PRAGMA table_info(other_cost_master)")
    col_info = cur.fetchall()   # (cid, name, type, notnull, dflt_value, pk)
    columns = {row[1] for row in col_info}
    pk_columns = {row[1] for row in col_info if row[5] > 0}

    # Migration is needed unless the primary key is EXACTLY (unit, store, month) -
    # checking column existence alone isn't enough, since an earlier buggy fix
    # could have added the "unit" column without fixing the primary key itself.
    if pk_columns != {"unit", "store", "month"}:
        # A simple ALTER TABLE ADD COLUMN can't change the PRIMARY KEY, so the
        # table has to be recreated properly to actually enforce (unit, store,
        # month) uniqueness - just adding the column (the old, buggy fix)
        # left the old (store, month) constraint in place, which made every
        # save fail with "ON CONFLICT clause does not match any PRIMARY KEY".
        print("  Migrating other_cost_master to be Unit-aware (one-time)...")

        # Only select cost columns that actually exist in the old table -
        # protects against crashing if an even older version had fewer columns.
        # Also carry forward the "unit" column's own value if it already
        # exists (from the earlier buggy fix) rather than always blanking it.
        existing_cost_cols = [c for c in col_names if c in columns]
        has_old_unit_col = "unit" in columns
        select_cols = (["unit"] if has_old_unit_col else []) + ["store", "month"] + existing_cost_cols
        cur.execute(f"SELECT {', '.join(select_cols)} FROM other_cost_master")
        old_rows = cur.fetchall()

        cur.execute("ALTER TABLE other_cost_master RENAME TO other_cost_master_old")
        conn.execute(f"""
            CREATE TABLE other_cost_master (
                unit TEXT,
                store TEXT,
                month TEXT,
                {cols_sql},
                PRIMARY KEY (unit, store, month)
            )
        """)

        # Old rows may not have tracked Unit at all - carry forward whatever
        # was there (or blank) so nothing is silently lost.
        placeholders = ", ".join("?" for _ in col_names)
        for row in old_rows:
            if has_old_unit_col:
                unit_val, store, month = row[0], row[1], row[2]
                cost_values_raw = row[3:]
            else:
                unit_val = ""
                store, month = row[0], row[1]
                cost_values_raw = row[2:]
            unit_val = unit_val or ""

            old_values = dict(zip(existing_cost_cols, cost_values_raw))
            values = [old_values.get(c) for c in col_names]   # None for any column that didn't exist before
            cur.execute(
                f"INSERT OR IGNORE INTO other_cost_master (unit, store, month, {col_list}) "
                f"VALUES (?, ?, ?, {placeholders})",
                [unit_val, store, month] + values,
            )

        cur.execute("DROP TABLE other_cost_master_old")
        conn.commit()

    return conn


def get_trimming_prod_qty_by_unit_store(month):
    """
    For TRIMMING-process stores, the Trimming/Insp cost is a flat Rs.1 per
    piece - so the cost simply equals the total Prod Qty produced in that
    (Unit, Store) for the month. Calculated live, never stored.
    """
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    if not os.path.exists(prod_db) or not month:
        return {}

    conn = sqlite3.connect(prod_db)
    ensure_year_month_column(conn)
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT unit, store, SUM(CAST(prod_qty AS REAL)) AS total_prod_qty
            FROM (
                SELECT DISTINCT source_entry_id, unit, store, item_code, prod_qty
                FROM production_register
                WHERE year_month = ?
                  AND unit IS NOT NULL AND TRIM(unit) != ''
                  AND store IS NOT NULL AND TRIM(store) != ''
                  AND item_code IS NOT NULL AND TRIM(item_code) != ''
            )
            GROUP BY unit, store
        """, (month,))
        result = {(unit, store): (qty or 0) for unit, store, qty in cur.fetchall()}
    except sqlite3.OperationalError:
        result = {}
    conn.close()
    return result


def get_process_master_unit_stores():
    """Distinct (Unit, Store, Process) from the Process Master (lives in Production_Entry.db)."""
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    if not os.path.exists(prod_db):
        return []
    conn = sqlite3.connect(prod_db)
    cur = conn.cursor()
    try:
        cur.execute("SELECT DISTINCT unit, wip_store, process FROM process_master ORDER BY unit, wip_store")
        combos = [(row[0], row[1], row[2]) for row in cur.fetchall() if row[0] and row[1]]
    except sqlite3.OperationalError:
        combos = []
    conn.close()
    return combos


@app.route("/other_cost")
def other_cost():
    months = get_all_relevant_months()
    selected_month = request.args.get("month", months[0] if months else None)

    unit_stores = get_process_master_unit_stores()
    trimming_qty = get_trimming_prod_qty_by_unit_store(selected_month) if selected_month else {}

    existing = {}
    if selected_month and unit_stores:
        conn = get_other_cost_connection()
        cur = conn.cursor()
        col_list = ", ".join(key for key, _ in OTHER_COST_FIELDS)
        cur.execute(f"""
            SELECT unit, store, {col_list} FROM other_cost_master WHERE month = ?
        """, (selected_month,))
        for row in cur.fetchall():
            unit, store = row[0], row[1]
            values = dict(zip([key for key, _ in OTHER_COST_FIELDS], row[2:]))
            existing[(unit, store)] = values
        conn.close()

    rows = []
    for unit, store, process in unit_stores:
        is_trimming = (process or "").strip().upper() == "TRIMMING"

        if is_trimming:
            # Flat Rs.1 per piece - the whole row is auto-calculated and locked,
            # with everything blank except Trimming/Insp.
            costs = {key: None for key, _ in OTHER_COST_FIELDS}
            costs["trimming_insp"] = trimming_qty.get((unit, store), 0)
        else:
            row_values = existing.get((unit, store), {})
            costs = {key: row_values.get(key) for key, _ in OTHER_COST_FIELDS}

        rows.append({
            "unit": unit,
            "store": store,
            "process": process,
            "is_trimming": is_trimming,
            "costs": costs,
        })

    return render_analytics_page(
        "other_cost.html", "other_cost",
        months_override=months,
        rows=rows,
        fields=OTHER_COST_FIELDS,
    )


@app.route("/other_cost/save", methods=["POST"])
def other_cost_save():
    body = request.get_json(silent=True) or {}
    month = body.get("month")
    rows = body.get("rows", [])

    if not month or not rows:
        return jsonify({"error": "Missing month or rows"}), 400

    conn = get_other_cost_connection()
    cur = conn.cursor()

    col_names = [key for key, _ in OTHER_COST_FIELDS]
    col_list = ", ".join(col_names)
    placeholders = ", ".join("?" for _ in col_names)
    update_clause = ", ".join(f"{c} = excluded.{c}" for c in col_names)

    saved = 0
    for row in rows:
        unit = row.get("unit")
        store = row.get("store")
        if not unit or not store:
            continue
        values = []
        for key, _ in OTHER_COST_FIELDS:
            raw = row.get(key, "")
            try:
                values.append(float(raw) if str(raw).strip() != "" else None)
            except (TypeError, ValueError):
                values.append(None)

        cur.execute(f"""
            INSERT INTO other_cost_master (unit, store, month, {col_list})
            VALUES (?, ?, ?, {placeholders})
            ON CONFLICT(unit, store, month) DO UPDATE SET {update_clause}
        """, [unit, store, month] + values)
        saved += 1

    conn.commit()
    conn.close()

    invalidate_cost_cache()
    return jsonify({"saved": saved})


@app.route("/weighted_price")
def weighted_price():
    months = get_available_months()
    selected_month = request.args.get("month", months[0] if months else None)
    search = request.args.get("search", "").strip()

    rows = []
    if selected_month and os.path.exists(ANALYTICS_DB):
        conn = sqlite3.connect(ANALYTICS_DB)
        cur = conn.cursor()
        query = """
            SELECT item_code, qty_purchased, value_purchased, weighted_price,
                   final_price, price_source_month
            FROM weighted_price_master
            WHERE month = ?
        """
        params = [selected_month]
        if search:
            query += " AND item_code LIKE ?"
            params.append(f"%{search}%")
        query += " ORDER BY item_code"

        cur.execute(query, params)
        for item_code, qty, value, weighted, final, source_month in cur.fetchall():
            rows.append({
                "item_code": item_code,
                "qty_purchased": qty,
                "value_purchased": value,
                "weighted_price": weighted,
                "final_price": final,
                "price_source_month": source_month,
                "is_stale": source_month != selected_month,
            })
        conn.close()

    return render_analytics_page(
        "weighted_price.html", "weighted_price",
        rows=rows, search=search,
    )


def get_item_descriptions():
    """{erp_code: description} from Item Master, for the Excel downloads.
    Empty if Item Master hasn't been run or has no description column."""
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    if not os.path.exists(prod_db):
        return {}
    conn = sqlite3.connect(prod_db)
    try:
        result = {c: d for c, d in conn.execute("SELECT erp_code, description FROM item_master")}
    except sqlite3.OperationalError:
        result = {}
    conn.close()
    return result


def month_label(month_key):
    try:
        return datetime(int(month_key[:4]), int(month_key[5:7]), 1).strftime("%B %Y")
    except (ValueError, TypeError):
        return month_key


def make_excel(title, headers, rows, money_cols=(), qty_cols=(), widths=None):
    """A simple formatted sheet: bold header row, frozen header, number formats."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title[:31]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True)
    for r in rows:
        ws.append(r)
    for col_idx, header in enumerate(headers, start=1):
        letter = openpyxl.utils.get_column_letter(col_idx)
        fmt = "#,##0.00" if header in money_cols else ("#,##0.####" if header in qty_cols else None)
        if fmt:
            for cell in ws[letter][1:]:
                cell.number_format = fmt
        ws.column_dimensions[letter].width = (widths or {}).get(header, max(12, len(header) + 2))
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.route("/weighted_price/download")
def weighted_price_download():
    """Excel of one month's weighted price - kind=purchase (from GRN) or
    kind=selling (from Sales Register + uploaded Other Sales)."""
    month = request.args.get("month", "")
    kind = request.args.get("kind", "purchase")
    if not month or not os.path.exists(ANALYTICS_DB):
        return "No month selected, or no data yet.", 400
    names = get_item_descriptions()
    conn = sqlite3.connect(ANALYTICS_DB)

    if kind == "selling":
        try:
            data = conn.execute("""
                SELECT item_code, qty_sold, value_sold, final_price, customer, unit
                FROM weighted_selling_price_master WHERE month = ? ORDER BY item_code
            """, (month,)).fetchall()
        except sqlite3.OperationalError:
            conn.close()
            return "No selling prices yet - run Weighted Selling Price from Data Sync first.", 404
        conn.close()
        headers = ["FG Code", "Description", "Qty Sold", "Value Sold (Rs)", "Weighted Selling Price (Rs)",
                   "Last Invoiced To", "Last Invoiced From (Unit)"]
        rows = [[c, names.get(c, ""), q, v, p, cust, u] for c, q, v, p, cust, u in data]
        buf = make_excel("Selling Price", headers, rows,
                         money_cols=("Value Sold (Rs)", "Weighted Selling Price (Rs)"), qty_cols=("Qty Sold",),
                         widths={"Description": 45, "Last Invoiced To": 40})
        filename = f"Weighted_Selling_Price_{month}.xlsx"
    else:
        try:
            data = conn.execute("""
                SELECT item_code, qty_purchased, value_purchased, weighted_price, final_price, price_source_month
                FROM weighted_price_master WHERE month = ? ORDER BY item_code
            """, (month,)).fetchall()
        except sqlite3.OperationalError:
            conn.close()
            return "No purchase prices yet - run Weighted Price Master from Data Sync first.", 404
        conn.close()
        headers = ["Item Code", "Description", "Qty Purchased", "Value Purchased (Rs)",
                   "Weighted Price This Month (Rs)", "Final Price Used (Rs)", "Price Source Month", "Carried Forward"]
        rows = [[c, names.get(c, ""), q, v, w, f, month_label(src) if src else "",
                 "Yes" if src and src != month else ("No price" if not src else "No")]
                for c, q, v, w, f, src in data]
        buf = make_excel("Purchase Price", headers, rows,
                         money_cols=("Value Purchased (Rs)", "Weighted Price This Month (Rs)", "Final Price Used (Rs)"),
                         qty_cols=("Qty Purchased",), widths={"Description": 45})
        filename = f"Weighted_Purchase_Price_{month}.xlsx"

    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def ensure_process_master_migrated(conn):
    """
    Migrates process_master from (Store) keyed to (Unit, Store) keyed.
    Since old rows only knew Store (not Unit), each old row gets expanded to
    apply to every (Unit, Store) combination currently seen in
    production_register that matches that Store name - so existing work
    isn't lost, it just gets applied per-unit going forward.
    """
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='process_master'")
    if not cur.fetchone():
        cur.execute("""
            CREATE TABLE process_master (
                unit TEXT,
                wip_store TEXT,
                process TEXT,
                PRIMARY KEY (unit, wip_store)
            )
        """)
        conn.commit()
        return

    cur.execute("PRAGMA table_info(process_master)")
    columns = {row[1] for row in cur.fetchall()}
    if "unit" in columns:
        return   # already migrated

    print("  Migrating process_master to be Unit-aware (one-time)...")

    cur.execute("SELECT wip_store, process FROM process_master")
    old_rows = cur.fetchall()   # [(store, process), ...]

    cur.execute("SELECT DISTINCT unit, store FROM production_register WHERE unit IS NOT NULL AND store IS NOT NULL")
    known_unit_store = cur.fetchall()

    cur.execute("ALTER TABLE process_master RENAME TO process_master_old")
    cur.execute("""
        CREATE TABLE process_master (
            unit TEXT,
            wip_store TEXT,
            process TEXT,
            PRIMARY KEY (unit, wip_store)
        )
    """)

    old_lookup = {store: process for store, process in old_rows}
    for unit, store in known_unit_store:
        if store in old_lookup:
            cur.execute(
                "INSERT OR IGNORE INTO process_master (unit, wip_store, process) VALUES (?, ?, ?)",
                (unit, store, old_lookup[store]),
            )

    cur.execute("DROP TABLE process_master_old")
    conn.commit()


def get_missing_process_master_stores():
    """
    Every distinct (Unit, Store) that's ever appeared in production_register,
    where that exact (Unit, Store) combination doesn't have a Process
    mapping yet in process_master.
    """
    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    if not os.path.exists(prod_db):
        return []

    conn = sqlite3.connect(prod_db)
    ensure_process_master_migrated(conn)
    cur = conn.cursor()
    cur.execute("""
        SELECT DISTINCT unit, store
        FROM production_register
        WHERE unit IS NOT NULL AND store IS NOT NULL
        ORDER BY unit, store
    """)
    all_combos = cur.fetchall()

    cur.execute("SELECT unit, wip_store FROM process_master")
    already_have = set(cur.fetchall())
    conn.close()

    return [(unit, store) for unit, store in all_combos if (unit, store) not in already_have]


@app.route("/process_master/download_missing")
def process_master_download_missing():
    missing = get_missing_process_master_stores()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Unit", "Store", "Process"])
    for unit, store in missing:
        ws.append([unit, store, None])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="process_master_missing_stores.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/process_master/upload_missing", methods=["POST"])
def process_master_upload_missing():
    if "file" not in request.files:
        return jsonify({"error": "No file was sent"}), 400

    uploaded_file = request.files["file"]
    if uploaded_file.filename == "":
        return jsonify({"error": "No file was selected"}), 400

    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        ws = wb.active
    except Exception as e:
        return jsonify({"error": f"Could not read the file: {e}"}), 400

    rows = list(ws.iter_rows(min_row=2, values_only=True))  # skip header row

    prod_db = os.path.join(DATABASE_DIR, "Production_Entry.db")
    conn = sqlite3.connect(prod_db)
    ensure_process_master_migrated(conn)
    cur = conn.cursor()

    updated = 0
    skipped_blank = 0

    for row in rows:
        if not row or len(row) < 3:
            continue
        unit, store, process = row[0], row[1], row[2]

        if not unit or not store:
            continue
        if process is None or str(process).strip() == "":
            skipped_blank += 1
            continue

        cur.execute(
            "INSERT INTO process_master (unit, wip_store, process) VALUES (?, ?, ?) "
            "ON CONFLICT(unit, wip_store) DO UPDATE SET process = excluded.process",
            (unit, store, str(process).strip()),
        )
        updated += 1

    conn.commit()
    conn.close()

    return jsonify({"updated": updated, "skipped_blank": skipped_blank})


@app.route("/data_sync")
def data_sync():
    return render_template("data_sync.html", nav_items=NAV_ITEMS, active_page="data_sync",
                            page_label="Data sync", scripts=SCRIPTS, months=[],
                            range_download_keys=set(DOWNLOAD_RANGE_CONFIG.keys()),
                            erp_login_ready=bool(ERP_LOGIN))


@app.route("/start/<key>", methods=["POST"])
def start(key):
    if key not in SCRIPTS:
        return jsonify({"error": "Unknown script"}), 404

    info = SCRIPTS[key]
    script_path = os.path.join(SCRIPTS_DIR, info["script"])

    if not os.path.exists(script_path):
        return jsonify({"error": f"Script not found: {info['script']} (expected in {SCRIPTS_DIR})"}), 404

    body = request.get_json(silent=True) or {}

    # ERP login: asked once when the dashboard starts (see ERP_LOGIN), passed
    # to the script in memory only. Never written to a file or the log.
    extra_env = None
    if info.get("needs_login"):
        if ERP_LOGIN:
            username, password = ERP_LOGIN["username"], ERP_LOGIN["password"]
        else:
            # Fallback when the login wasn't given at startup: the browser box.
            username = str(body.get("username", "")).strip()
            password = str(body.get("password", ""))
            if not username or not password:
                return jsonify({"error": "ERP login missing - close and restart the dashboard, "
                                         "then type your username and password in its window"}), 400
            ERP_LOGIN.update(username=username, password=password)   # remember until the dashboard closes
        extra_env = {"ERP_USERNAME": username, "ERP_PASSWORD": password}

    args = []
    if info["needs_month"]:
        year = str(body.get("year", "")).strip()
        month = str(body.get("month", "")).strip()
        if not year or not month:
            return jsonify({"error": "Year and month are required for this script"}), 400
        args = [year, month]

    job_id = start_job(key, script_path, args, extra_env)
    return jsonify({"job_id": job_id})


@app.route("/upload/<key>", methods=["POST"])
def upload(key):
    """Accepts a file upload (e.g. a new PROCESS_MASTER.csv), saves it into
    the SCRIPTS folder (overwriting the old one), then runs the script."""
    if key not in SCRIPTS:
        return jsonify({"error": "Unknown script"}), 404

    info = SCRIPTS[key]
    if not info.get("has_upload"):
        return jsonify({"error": "This script doesn't accept file uploads"}), 400

    if "file" not in request.files:
        return jsonify({"error": "No file was sent"}), 400

    uploaded_file = request.files["file"]
    if uploaded_file.filename == "":
        return jsonify({"error": "No file was selected"}), 400

    dest_path = os.path.join(SCRIPTS_DIR, info["upload_filename"])
    uploaded_file.save(dest_path)

    script_path = os.path.join(SCRIPTS_DIR, info["script"])
    if not os.path.exists(script_path):
        return jsonify({"error": f"File saved, but script not found: {info['script']}"}), 404

    job_id = start_job(key, script_path, [])
    return jsonify({"job_id": job_id})


@app.route("/status/<key>")
def status(key):
    """Lets a freshly-loaded page ask 'is anything running (or recently
    finished) for this card?' so it can reconnect and show the right state
    instead of looking blank/stopped after navigating away and back."""
    job_id = ACTIVE_JOB_BY_KEY.get(key)
    if not job_id or job_id not in JOBS:
        return jsonify({"job_id": None})

    job = JOBS[job_id]
    return jsonify({
        "job_id": job_id,
        "finished": job["finished"],
        "returncode": job["returncode"],
    })


@app.route("/stream/<job_id>")
def stream(job_id):
    """Server-Sent Events endpoint. Immediately replays everything that's
    happened so far (so a reconnecting tab catches up instantly), then
    keeps streaming new lines live. Each connection gets its own listener
    queue, so multiple tabs/viewers can watch the same job at once."""
    if job_id not in JOBS:
        return "Unknown job", 404

    job = JOBS[job_id]
    my_queue = queue.Queue()

    with job["lock"]:
        history_snapshot = list(job["log_lines"])
        already_finished = job["finished"]
        returncode_snapshot = job["returncode"]
        if not already_finished:
            job["listeners"].append(my_queue)

    def generate():
        for line in history_snapshot:
            yield f"data: {line}\n\n"

        if already_finished:
            yield f"event: done\ndata: {returncode_snapshot}\n\n"
            return

        while True:
            line = my_queue.get()
            if line == "__DONE__":
                yield f"event: done\ndata: {job['returncode']}\n\n"
                break
            yield f"data: {line}\n\n"

    return Response(generate(), mimetype="text/event-stream")


@app.route("/download/<key>")
def download(key):
    if key not in SCRIPTS:
        return "Unknown script", 404
    db_path = os.path.join(DATABASE_DIR, SCRIPTS[key]["db"])
    if not os.path.exists(db_path):
        return f"Database not found yet: {SCRIPTS[key]['db']}", 404
    return send_file(db_path, as_attachment=True)


if __name__ == "__main__":
    print(f"Scripts folder: {SCRIPTS_DIR}")
    print(f"Database folder: {DATABASE_DIR}")
    ensure_weighted_price_indexes()
    ask_erp_login_at_startup()
    print("Dashboard running - open http://localhost:5000 in your browser")
    app.run(debug=False, port=5000)
