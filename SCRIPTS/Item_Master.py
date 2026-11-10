"""
Downloads the Item Master (covers both companies in one file) and saves it
into the item_master table in Production_Entry.db.

Static fields (Company, Description, Store, etc.) are simply kept up to
date each run. "Item Type" is different - since an item can switch between
Production/Purchase over time, we keep a SEPARATE column per month you run
this script (e.g. item_type_2026_09), so historical production entries can
always look up what an item's type was during their own month.

First-ever run: since there's no history, we backfill every month from your
data baseline (June 2026) through the current month with this run's values.
Later runs: only add/update the current month's column.

Run: python import_item_master.py
"""

import requests
import pandas as pd
import sqlite3
import io
from datetime import datetime

BASE_URL = "https://unitech.tarka.in"
LOGIN_URL = f"{BASE_URL}/app.php"
DATA_URL = f"{BASE_URL}/app.php"

# ERP login comes from the dashboard (or is asked when run by hand) - nothing is saved in this file
from erp_credentials import get_credentials
CLIENT_COMPANY = "9"   # Unitech login - the export covers both companies anyway

HEADERS = {
    "X-Requested-With": "XMLHttpRequest",
    "Referer": f"{BASE_URL}/MainPage.html",
}

import os
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
from costing_config import database_dir   # set in costing\config\settings.json
DATABASE_DIR = database_dir()
os.makedirs(DATABASE_DIR, exist_ok=True)
DB_FILE = os.path.join(DATABASE_DIR, "Production_Entry.db")
BASELINE_YEAR_MONTH = (2026, 6)   # matches the overall data baseline (1 Jun 2026)

# Static fields we keep (everything through "Item Type" in the source file,
# minus Item Type itself which is handled separately as month columns)
COLUMN_MAP = {
    "internal_code": "internal_code",
    "Company": "company",
    "ERP Code": "erp_code",
    "Description": "description",
    "Party Code": "party_code",
    "Drawing No": "drawing_no",
    "Parent Item": "parent_item",
    "Specification": "specification",
    "Revision": "revision",
    "UOM": "uom",
    "Store": "store",
    "Item Group": "item_group",
    "Item Sub Group": "item_sub_group",
    "Customer": "customer",
}

STATIC_SQL_COLUMNS = list(COLUMN_MAP.values())


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


def download_item_master(session):
    resp = session.get(
        DATA_URL,
        params={
            "a": "DownloadExcelFormatFullData",
            "schema_id": 84,
            "system_module_table_name": "item",
            "filter_cond": "false",
            "selected_db_fields_arr": "",
            "is_filter": 0,
        },
        headers=HEADERS,
    )
    return resp.content


def months_between(start_year_month, end_year_month):
    """List of (year, month) tuples from start through end, inclusive."""
    y, m = start_year_month
    end_y, end_m = end_year_month
    result = []
    while (y, m) <= (end_y, end_m):
        result.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return result


def month_column_name(year, month):
    return f"item_type_{year}_{month:02d}"


def get_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS item_master (
            {", ".join(f"{col} TEXT" for col in STATIC_SQL_COLUMNS)},
            PRIMARY KEY (erp_code)
        )
    """)
    conn.commit()
    return conn


def existing_month_columns(conn):
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(item_master)")
    return [row[1] for row in cur.fetchall() if row[1].startswith("item_type_")]


def ensure_column(conn, col_name):
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(item_master)")
    existing = {row[1] for row in cur.fetchall()}
    if col_name not in existing:
        cur.execute(f"ALTER TABLE item_master ADD COLUMN {col_name} TEXT")
        conn.commit()
        print(f"  Added column: {col_name}")


def upsert_rows(conn, df, month_columns_to_fill):
    cur = conn.cursor()
    count = 0
    for _, row in df.iterrows():
        erp_code = row.get("ERP Code")
        if pd.isna(erp_code) or not str(erp_code).strip():
            continue

        static_values = [row.get(excel_col) for excel_col in COLUMN_MAP.keys()]
        static_values = [None if pd.isna(v) else str(v) for v in static_values]

        item_type = row.get("Item Type")
        item_type = None if pd.isna(item_type) else str(item_type)

        all_columns = STATIC_SQL_COLUMNS + month_columns_to_fill
        all_values = static_values + [item_type] * len(month_columns_to_fill)

        placeholders = ", ".join("?" for _ in all_columns)
        update_clause = ", ".join(f"{c} = excluded.{c}" for c in all_columns if c != "erp_code")

        cur.execute(
            f"INSERT INTO item_master ({', '.join(all_columns)}) VALUES ({placeholders}) "
            f"ON CONFLICT(erp_code) DO UPDATE SET {update_clause}",
            all_values,
        )
        count += 1
    conn.commit()
    return count


if __name__ == "__main__":
    session = login()
    if not session:
        exit()

    print("Downloading item master...")
    file_bytes = download_item_master(session)
    df = pd.read_csv(io.BytesIO(file_bytes), low_memory=False)
    print(f"  {len(df)} rows in file.")

    conn = get_connection()

    now = datetime.now()
    current_year_month = (now.year, now.month)
    current_col = month_column_name(*current_year_month)

    existing_months = existing_month_columns(conn)

    if not existing_months:
        # First-ever run: backfill every month from baseline through current
        months_to_fill = months_between(BASELINE_YEAR_MONTH, current_year_month)
        print(f"First run - backfilling months: {[month_column_name(y, m) for y, m in months_to_fill]}")
    else:
        # Normal run: just this month
        months_to_fill = [current_year_month]

    month_columns_to_fill = [month_column_name(y, m) for y, m in months_to_fill]
    for col in month_columns_to_fill:
        ensure_column(conn, col)

    count = upsert_rows(conn, df, month_columns_to_fill)
    conn.close()

    print(f"\n✅ Done. {count} items imported/updated. Filled columns: {month_columns_to_fill}")
