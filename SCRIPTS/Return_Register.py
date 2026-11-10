"""
Incremental Material Return Register sync - filtered to:
  - item Code contains "RMI"
  - From Store is an Injection Moulding store (opposite direction of
    Material Issue Report, where it was To WareHouse)

Note: this report has no "Company" column in the source data, so we add it
ourselves based on which company session we downloaded from.

Run: python sync_material_return_register.py
"""

import requests
import pandas as pd
import sqlite3
import hashlib
import time
import io
from datetime import datetime, timedelta

BASE_URL = "https://unitech.tarka.in"
LOGIN_URL = f"{BASE_URL}/app.php"
DATA_URL = f"{BASE_URL}/app.php"

# ERP login comes from the dashboard (or is asked when run by hand) - nothing is saved in this file
from erp_credentials import get_credentials

HEADERS = {
    "X-Requested-With": "XMLHttpRequest",
    "Referer": f"{BASE_URL}/MainPage.html",
}

BASELINE_DATE = "01/06/2026"

COMPANIES = [
    {
        "name": "Unitech",
        "client_company": "9",
        "plant": "9",
        "units": [
            {"name": "U1", "showroom": "21"},
            {"name": "U2", "showroom": "25"},
            {"name": "U3", "showroom": "29"},
            {"name": "U4", "showroom": "26"},
        ],
    },
    {
        "name": "Vtech",
        "client_company": "11",
        "plant": "11",
        "units": [
            {"name": "V1", "showroom": "27"},
            {"name": "V2", "showroom": "28"},
        ],
    },
]

import os
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_DIR = os.path.join(SCRIPT_DIR, "..", "DATABASE")
os.makedirs(DATABASE_DIR, exist_ok=True)
DB_FILE = os.path.join(DATABASE_DIR, "Material_Return_Register.db")
PROCESS_MASTER_DB = os.path.join(DATABASE_DIR, "Production_Entry.db")

COLUMN_MAP = {
    "Company": "company",     # added by us, not in source file
    "Unit": "unit",
    "Date": "date",
    "Entry No": "entry_no",
    "From Store": "from_store",
    "To Store": "to_store",
    "Remarks": "remarks",
    "item Code": "item_code",
    "item": "item_name",
    "UOM": "uom",
    "Quantity": "quantity",
    "Notes": "notes",
}

SQL_COLUMNS = list(COLUMN_MAP.values())


# ---------- Database ----------

def get_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS material_return_register (
            row_key TEXT PRIMARY KEY,
            {", ".join(f"{col} TEXT" for col in SQL_COLUMNS)}
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS sync_state (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            last_pulled_date TEXT
        )
    """)
    conn.commit()
    return conn


def get_injection_moulding_stores():
    conn = sqlite3.connect(PROCESS_MASTER_DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT wip_store FROM process_master
        WHERE UPPER(process) = 'INJECTION MOULDING'
    """)
    stores = {row[0] for row in cur.fetchall()}
    conn.close()
    return stores


def get_last_pulled_date(conn):
    cur = conn.cursor()
    cur.execute("SELECT last_pulled_date FROM sync_state WHERE id = 1")
    row = cur.fetchone()
    return row[0] if row else None


def set_last_pulled_date(conn, date_str):
    conn.execute(
        "INSERT INTO sync_state (id, last_pulled_date) VALUES (1, ?) "
        "ON CONFLICT(id) DO UPDATE SET last_pulled_date = excluded.last_pulled_date",
        (date_str,),
    )
    conn.commit()


def make_row_key(row):
    key_str = "|".join(str(row.get(c, "")) for c in
                        ["Entry No", "item Code", "Date", "Quantity"])
    return hashlib.md5(key_str.encode("utf-8")).hexdigest()


def insert_rows(conn, df):
    cur = conn.cursor()
    inserted = 0
    for _, row in df.iterrows():
        row_key = make_row_key(row)
        values = [row.get(excel_col) for excel_col in COLUMN_MAP.keys()]
        values = [None if pd.isna(v) else str(v) for v in values]

        placeholders = ", ".join("?" for _ in SQL_COLUMNS)
        cur.execute(
            f"INSERT OR IGNORE INTO material_return_register (row_key, {', '.join(SQL_COLUMNS)}) "
            f"VALUES (?, {placeholders})",
            [row_key] + values,
        )
        if cur.rowcount > 0:
            inserted += 1
    conn.commit()
    return inserted


# ---------- Date chunking ----------

def daterange_chunks(from_date_str, to_date_str, chunk_days=7):
    start = datetime.strptime(from_date_str, "%d/%m/%Y")
    end = datetime.strptime(to_date_str, "%d/%m/%Y")
    chunks = []
    current = start
    while current <= end:
        chunk_end = min(current + timedelta(days=chunk_days - 1), end)
        chunks.append((current.strftime("%d/%m/%Y"), chunk_end.strftime("%d/%m/%Y")))
        current = chunk_end + timedelta(days=1)
    return chunks


# ---------- ERP calls ----------

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


def download_material_return(session, plant, showroom, from_date, to_date):
    resp = session.get(
        DATA_URL,
        params={
            "a": "getDownloadReportFile",
            "report_id": 894,
            "schema_name": "Material~Return~Register",
            "system_module_table_name": "material_return_register",
            "plant": plant,
            "showroom": showroom,
            "from_date": from_date,
            "to_date": to_date,
            "entry_no": 0,
            "item": 0,
            "data_type": "Detailed",
            "report_type": "Excel",
        },
        headers=HEADERS,
    )
    return resp.content


def parse_material_return(file_bytes, valid_stores, company_name):
    raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl")

    header_row_idx = None
    for i in range(len(raw)):
        row_values = raw.iloc[i].astype(str).tolist()
        if "Entry No" in row_values:
            header_row_idx = i
            break

    if header_row_idx is None:
        return pd.DataFrame()

    columns = raw.iloc[header_row_idx].tolist()
    data = raw.iloc[header_row_idx + 1:].reset_index(drop=True)
    data.columns = columns
    data = data.dropna(how="all")

    if "item Code" in data.columns:
        data = data[data["item Code"].notna()]

    # Filters: item Code contains "RMI", and From Store is an Injection
    # Moulding store (opposite direction of Material Issue Report)
    data = data[data["item Code"].astype(str).str.contains("RMI")]
    data = data[data["From Store"].isin(valid_stores)]

    # This report has no Company column in the source - add it ourselves
    data.insert(0, "Company", company_name)

    return data


# ---------- Main ----------

if __name__ == "__main__":
    valid_stores = get_injection_moulding_stores()
    print(f"Injection Moulding stores (from process_master): {sorted(valid_stores)}\n")

    conn = get_connection()

    last_pulled = get_last_pulled_date(conn)
    from_date = last_pulled if last_pulled else BASELINE_DATE
    to_date = datetime.now().strftime("%d/%m/%Y")

    print(f"Pulling from {from_date} to {to_date} (overlap on {from_date} is expected and safe)...\n")

    total_inserted = 0

    for company in COMPANIES:
        name = company["name"]
        client_company = company["client_company"]
        plant = company["plant"]

        print(f"===== {name} =====")
        session = login(client_company)
        if not session:
            continue

        for unit in company["units"]:
            unit_name = unit["name"]
            showroom = unit["showroom"]

            chunks = daterange_chunks(from_date, to_date, chunk_days=7)
            print(f"  {name} / {unit_name}: fetching {len(chunks)} weekly chunk(s)...")

            unit_total_fetched = 0
            unit_total_new = 0

            for chunk_from, chunk_to in chunks:
                file_bytes = download_material_return(session, plant, showroom, chunk_from, chunk_to)
                df = parse_material_return(file_bytes, valid_stores, name)

                if df.empty:
                    continue

                inserted = insert_rows(conn, df)
                unit_total_fetched += len(df)
                unit_total_new += inserted
                total_inserted += inserted
                time.sleep(0.3)

            print(f"    {unit_total_fetched} matching rows fetched, {unit_total_new} genuinely new.")

    set_last_pulled_date(conn, to_date)
    conn.close()

    print(f"\n✅ Done. {total_inserted} new rows added. Next run will start from {to_date}.")
