"""
Vendor Price Master - monthly snapshot of the ERP's supplier item mapping
(the "Vendor Price Master" screen: the purchase price currently mapped to
each item/supplier).

Each run downloads the full screen and stores it as THIS month's snapshot
in DATABASE\\Vendor_Price_Master.db, table vendor_price_snapshot:

  * month      - e.g. 2026_10 (same format as the rest of the dashboard)
  * backfilled - 0 = really downloaded in that month
                 1 = copied from the first-ever download, because the ERP
                     screen only shows today's prices and there's no older
                     snapshot for months before this script existed
  * every column from the ERP file, as-is

Running it again in the same month replaces that month's snapshot with a
fresh copy. Earlier months are never touched, so history builds up month by
month. The very first run also fills every month from June 2026 up to now
(marked backfilled = 1), so earlier months have something to fall back on.

Run: python Vendor_Price_Master.py
"""

import io
import os
import re
import sqlite3
from datetime import datetime

import pandas as pd
import requests

# ERP login comes from the dashboard (or is asked when run by hand) - nothing is saved in this file
from erp_credentials import get_credentials

BASE_URL = "https://unitech.tarka.in"
LOGIN_URL = f"{BASE_URL}/app.php"
DATA_URL = f"{BASE_URL}/app.php"
CLIENT_COMPANY = "9"   # Unitech login - same as Item Master

HEADERS = {
    "X-Requested-With": "XMLHttpRequest",
    "Referer": f"{BASE_URL}/MainPage.html",
}

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_DIR = os.path.join(SCRIPT_DIR, "..", "DATABASE")
os.makedirs(DATABASE_DIR, exist_ok=True)
DB_FILE = os.path.join(DATABASE_DIR, "Vendor_Price_Master.db")
TABLE = "vendor_price_snapshot"

BASELINE_YEAR_MONTH = (2026, 6)   # matches the overall data baseline (1 Jun 2026)


# ---------- ERP ----------

def login():
    username, password = get_credentials()
    session = requests.Session()
    session.get(BASE_URL)
    resp = session.post(
        LOGIN_URL,
        params={"a": "validateUsernameAndPassword.demo"},
        data={"usr": username, "password": password, "client_company": CLIENT_COMPANY},
        headers=HEADERS,
    )
    data = resp.json()
    if data.get("status") == "success":
        print("✅ Login successful")
        return session
    print("❌ Login failed - check your ERP username and password:", data)
    return None


def download_vendor_price_master(session):
    resp = session.get(
        DATA_URL,
        params={
            "a": "DownloadExcelFormatFullData",
            "schema_id": 1112,
            "system_module_table_name": "supplier_item_mapping",
            "filter_cond": "false",
            "selected_db_fields_arr": "",
            "is_filter": 0,
        },
        headers=HEADERS,
    )
    return resp.content


def parse_file(file_bytes):
    """The ERP sends this kind of screen as CSV (like Item Master); fall
    back to Excel in case this screen is different."""
    if not file_bytes:
        return pd.DataFrame()
    head = file_bytes[:8]
    if head.startswith(b"PK"):                       # .xlsx
        return pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    if head.startswith(b"\xd0\xcf\x11\xe0"):         # old .xls
        return pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    return pd.read_csv(io.BytesIO(file_bytes), dtype=str, low_memory=False)


# ---------- Database ----------

def sql_name(column):
    """'Supplier Name' -> 'supplier_name', so every ERP column gets a safe SQL name."""
    name = re.sub(r"[^0-9a-zA-Z]+", "_", str(column)).strip("_").lower()
    if not name:
        name = "col"
    if name[0].isdigit():
        name = "c_" + name
    return name


def unique_sql_names(columns):
    seen = {}
    result = []
    for col in columns:
        base = sql_name(col)
        if base in ("month", "backfilled", "row_no", "downloaded_at"):
            base = "erp_" + base
        name = base
        n = 2
        while name in seen:
            name = f"{base}_{n}"
            n += 1
        seen[name] = True
        result.append(name)
    return result


def get_connection():
    conn = sqlite3.connect(DB_FILE, timeout=60)
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS {TABLE} (
            month TEXT NOT NULL,
            backfilled INTEGER NOT NULL DEFAULT 0,
            downloaded_at TEXT,
            row_no INTEGER NOT NULL,
            PRIMARY KEY (month, row_no)
        )
    """)
    conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{TABLE}_month ON {TABLE} (month)")
    conn.commit()
    return conn


def ensure_columns(conn, sql_columns):
    existing = {row[1] for row in conn.execute(f"PRAGMA table_info({TABLE})")}
    for col in sql_columns:
        if col not in existing:
            conn.execute(f'ALTER TABLE {TABLE} ADD COLUMN "{col}" TEXT')
    conn.commit()


def stored_months(conn):
    return [r[0] for r in conn.execute(f"SELECT DISTINCT month FROM {TABLE} ORDER BY month")]


def months_between(start_year_month, end_year_month):
    y, m = start_year_month
    end_y, end_m = end_year_month
    result = []
    while (y, m) <= (end_y, end_m):
        result.append(f"{y}_{m:02d}")
        m += 1
        if m > 12:
            m = 1
            y += 1
    return result


def save_snapshot(conn, df, sql_columns, month, backfilled, downloaded_at):
    """Replaces one month's snapshot with this download (one transaction)."""
    conn.execute(f"DELETE FROM {TABLE} WHERE month = ?", (month,))
    cols = ["month", "backfilled", "downloaded_at", "row_no"] + sql_columns
    placeholders = ", ".join("?" for _ in cols)
    quoted = ", ".join(f'"{c}"' for c in cols)
    rows = []
    for i, values in enumerate(df.itertuples(index=False, name=None), start=1):
        clean = [None if (v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "") else str(v).strip()
                 for v in values]
        rows.append([month, backfilled, downloaded_at, i] + clean)
    conn.executemany(f"INSERT INTO {TABLE} ({quoted}) VALUES ({placeholders})", rows)
    conn.commit()
    return len(rows)


# ---------- Main ----------

if __name__ == "__main__":
    session = login()
    if not session:
        raise SystemExit(1)

    print("Downloading Vendor Price Master (supplier item mapping)...")
    file_bytes = download_vendor_price_master(session)
    try:
        df = parse_file(file_bytes)
    except Exception as e:
        print(f"❌ Couldn't read the downloaded file: {e}")
        print(f"   First bytes received: {file_bytes[:200]!r}")
        raise SystemExit(1)

    df = df.dropna(how="all")
    if df.empty:
        print("❌ The ERP sent back no rows - nothing saved. Check that your login can open the Vendor Price Master screen.")
        raise SystemExit(1)

    sql_columns = unique_sql_names(df.columns)
    print(f"  {len(df)} rows, {len(sql_columns)} columns in file.")
    print("  Columns: " + ", ".join(str(c) for c in df.columns))

    now = datetime.now()
    current_month = f"{now.year}_{now.month:02d}"
    downloaded_at = now.strftime("%d/%m/%Y %H:%M")

    conn = get_connection()
    ensure_columns(conn, sql_columns)
    already = stored_months(conn)

    if not already:
        # First-ever run: earlier months have no snapshot of their own, so
        # they get a copy of today's prices, clearly marked as backfilled.
        for month in months_between(BASELINE_YEAR_MONTH, (now.year, now.month)):
            backfilled = 0 if month == current_month else 1
            save_snapshot(conn, df, sql_columns, month, backfilled, downloaded_at)
            print(f"  {month}: saved{' (backfilled - copy of today' + chr(39) + 's prices)' if backfilled else ''}")
    else:
        count = save_snapshot(conn, df, sql_columns, current_month, 0, downloaded_at)
        print(f"  {current_month}: {count} rows saved (replaces any earlier download this month).")

    months_now = stored_months(conn)
    conn.close()

    print("\n  First 3 rows, for checking:")
    with pd.option_context("display.max_columns", None, "display.width", 250):
        print(df.head(3).to_string(index=False))

    print(f"\n✅ Done. Snapshot months stored: {', '.join(months_now)}")
