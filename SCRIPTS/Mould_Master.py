"""
Downloads the Mould Master (all units, one company at a time) directly from
the ERP's export, and saves it into the mould_master table in
Production_Entry.db. Re-run this any time to refresh (updates existing
codes, adds new ones).

Run: python download_mould_master.py
"""

import requests
import pandas as pd
import sqlite3
import io

BASE_URL = "https://unitech.tarka.in"
LOGIN_URL = f"{BASE_URL}/app.php"
DATA_URL = f"{BASE_URL}/app.php"

# ERP login comes from the dashboard (or is asked when run by hand) - nothing is saved in this file
from erp_credentials import get_credentials

HEADERS = {
    "X-Requested-With": "XMLHttpRequest",
    "Referer": f"{BASE_URL}/MainPage.html",
}

COMPANIES = [
    {"name": "Unitech", "client_company": "9"},
    {"name": "Veetech", "client_company": "11"},
]

import os
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
from costing_config import database_dir   # set in costing\config\settings.json
DATABASE_DIR = database_dir()
os.makedirs(DATABASE_DIR, exist_ok=True)
DB_FILE = os.path.join(DATABASE_DIR, "Production_Entry.db")

# Excel column -> SQL column (sanitized)
COLUMN_MAP = {
    "internal_code": "internal_code",
    "Company": "company",
    "Unit": "unit",
    "Model": "model",
    "Machine": "machine",
    "Mould Owner": "mould_owner",
    "Code": "code",
    "Name": "name",
    "Cycle Time": "cycle_time",
    "Runner Wt": "runner_wt",
    "Part Wt": "part_wt",
    "Shot Wt(Kgs)": "shot_wt_kgs",
    "Shots Cutoff Date": "shots_cutoff_date",
    "PM Opening Shots": "pm_opening_shots",
    "Mould Life": "mould_life",
    "Total Opening Shots": "total_opening_shots",
    "PM Cycle Count": "pm_cycle_count",
    "PM Alert 1": "pm_alert_1",
    "PM Alert 2": "pm_alert_2",
    "PM Alert 3": "pm_alert_3",
    "MM Cycle Count": "mm_cycle_count",
    "MM Alert 1": "mm_alert_1",
    "MM Alert 2": "mm_alert_2",
    "MM Alert 3": "mm_alert_3",
    "Mould Manufacturer": "mould_manufacturer",
    "HRS Type": "hrs_type",
    "Layout Type": "layout_type",
    "MFG/Purchase Date": "mfg_purchase_date",
    "Warranty/AMC Date": "warranty_amc_date",
    "Active": "active",
}

SQL_COLUMNS = list(COLUMN_MAP.values())


def login(client_company):
    username, password = get_credentials()
    session = requests.Session()
    session.get(BASE_URL)
    resp = session.post(
        LOGIN_URL,
        params={"a": "validateUsernameAndPassword.demo"},
        data={"usr": username, "password": password, "client_company": client_company},
        headers=HEADERS,
    )
    data = resp.json()
    if data.get("status") == "success":
        print("✅ Login successful")
        return session
    print("❌ Login failed - check your ERP username and password:", data)
    return None


def download_mould_master(session):
    resp = session.get(
        DATA_URL,
        params={
            "a": "DownloadExcelFormatFullData",
            "schema_id": 904,
            "system_module_table_name": "mould_master",
            "filter_cond": "false",
            "selected_db_fields_arr": "",
            "is_filter": 0,
        },
        headers=HEADERS,
    )
    return resp.content


def get_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS mould_master (
            {", ".join(f"{col} TEXT" for col in SQL_COLUMNS)},
            PRIMARY KEY (code)
        )
    """)
    conn.commit()
    return conn


def upsert_rows(conn, df):
    cur = conn.cursor()
    count = 0
    for _, row in df.iterrows():
        code = row.get("Code")
        if pd.isna(code) or not str(code).strip():
            continue  # skip blank/header artifact rows

        values = [row.get(excel_col) for excel_col in COLUMN_MAP.keys()]
        values = [None if pd.isna(v) else str(v) for v in values]

        placeholders = ", ".join("?" for _ in SQL_COLUMNS)
        set_clause = ", ".join(f"{c} = excluded.{c}" for c in SQL_COLUMNS if c != "code")
        cur.execute(
            f"INSERT INTO mould_master ({', '.join(SQL_COLUMNS)}) VALUES ({placeholders}) "
            f"ON CONFLICT(code) DO UPDATE SET {set_clause}",
            values,
        )
        count += 1
    conn.commit()
    return count


if __name__ == "__main__":
    conn = get_connection()
    total = 0

    for company in COMPANIES:
        name = company["name"]
        client_company = company["client_company"]

        print(f"\n===== {name} =====")
        session = login(client_company)
        if not session:
            continue

        print("Downloading mould master...")
        file_bytes = download_mould_master(session)
        df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
        print(f"  {len(df)} rows in file.")

        count = upsert_rows(conn, df)
        total += count
        print(f"  ✅ {count} mould codes imported/updated for {name}.")

    conn.close()
    print(f"\n✅ Done. {total} total rows processed.")
