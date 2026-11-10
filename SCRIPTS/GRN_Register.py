"""
Incremental GRN Register sync - same approach as Sales Register.

- First run: pulls everything from BASELINE_DATE (1 Jun 2026) through today.
- Later runs: re-pulls from the last recorded date through today (overlap
  catches invoices/GRNs added later that day). Duplicate-safe via a hash key.
- Downloads in weekly chunks to avoid the legacy .xls 65,536-row limit.

Run: python sync_grn_register.py
"""

import requests
import pandas as pd
import sqlite3
import hashlib
import xlrd
import time
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
DB_FILE = os.path.join(DATABASE_DIR, "GRN_Register.db")

COLUMN_MAP = {
    "UNIT": "unit",
    "STORE": "store",
    "GRN Date": "grn_date",
    "GRN Created Date": "grn_created_date",
    "GRN No": "grn_no",
    "GRN Type": "grn_type",
    "Gate Entry No": "gate_entry_no",
    "Gate Entry Date": "gate_entry_date",
    "INV/DC No": "inv_dc_no",
    "INV/DC Date": "inv_dc_date",
    "PO No": "po_no",
    "PO Date": "po_date",
    "Vehicle No": "vehicle_no",
    "Supplier": "supplier",
    "Paid By Customer": "paid_by_customer",
    "Item Group": "item_group",
    "Item Code": "item_code",
    "Item Name": "item_name",
    "Category": "category",
    "Parent Category": "parent_category",
    "UOM": "uom",
    "HSN CODE": "hsn_code",
    "Quantity": "quantity",
    "Accepted QTY": "accepted_qty",
    "Shortage QTY": "shortage_qty",
    "Rate": "rate",
    "Disc %": "disc_pct",
    "Taxable Value": "taxable_value",
    "CGST": "cgst",
    "CGST VALUE": "cgst_value",
    "SGST": "sgst",
    "SGST VALUE": "sgst_value",
    "IGST": "igst",
    "IGST VALUE": "igst_value",
    "ROUND(+)": "round_plus",
    "ROUND(-)": "round_minus",
    "TOTAL VALUE": "total_value",
    "ACCOUNT NAME": "account_name",
    "Ack_Status": "ack_status",
}

SQL_COLUMNS = ["company"] + list(COLUMN_MAP.values())

# GRN types that are not real purchases - never stored. "Sample" GRNs carry
# token/zero rates and would drag the weighted purchase price down.
EXCLUDED_GRN_TYPES = {"sample"}


# ---------- Database ----------

def get_connection():
    conn = sqlite3.connect(DB_FILE, timeout=60)
    cur = conn.cursor()

    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS grn_register (
            row_key TEXT PRIMARY KEY,
            {", ".join(f"{col} TEXT" for col in SQL_COLUMNS)}
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS sync_state (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            last_pulled_date TEXT
        )
    """)

    conn.commit()
    return conn


def delete_excluded_grn_types(conn):
    """Removes Sample GRNs that were stored before they were excluded.
    Runs every time (harmless when there's nothing to remove)."""
    placeholders = ", ".join("?" for _ in EXCLUDED_GRN_TYPES)
    cur = conn.execute(
        f"DELETE FROM grn_register WHERE LOWER(TRIM(COALESCE(grn_type, ''))) IN ({placeholders})",
        list(EXCLUDED_GRN_TYPES),
    )
    conn.commit()
    return cur.rowcount


def drop_excluded_grn_types(df):
    """Drops Sample GRN lines from a freshly downloaded file before saving."""
    if df.empty or "GRN Type" not in df.columns:
        return df, 0
    grn_type = df["GRN Type"].fillna("").astype(str).str.strip().str.lower()
    keep = ~grn_type.isin(EXCLUDED_GRN_TYPES)
    return df[keep], int((~keep).sum())


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


def make_row_key(company, row):
    key_str = "|".join(str(row.get(c, "")) for c in
                        ["GRN No", "Item Code", "GRN Date", "Quantity"])
    key_str = company + "|" + key_str
    return hashlib.md5(key_str.encode("utf-8")).hexdigest()


def insert_rows(conn, df):
    cur = conn.cursor()
    inserted = 0
    for _, row in df.iterrows():
        row_key = make_row_key(row["Company"], row)
        values = [row["Company"]] + [row.get(excel_col) for excel_col in COLUMN_MAP.keys()]
        values = [None if pd.isna(v) else str(v) for v in values]

        placeholders = ", ".join("?" for _ in SQL_COLUMNS)
        cur.execute(
            f"INSERT OR IGNORE INTO grn_register (row_key, {', '.join(SQL_COLUMNS)}) "
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


def download_grn_register(session, plant, showroom, from_date, to_date):
    resp = session.get(
        DATA_URL,
        params={
            "a": "getDownloadReportFile",
            "report_id": 191,
            "schema_name": "GRN~(Inv)~Register",
            "system_module_table_name": "grn_inv_register",
            "plant": plant,
            "showroom": showroom,
            "from_date": from_date,
            "to_date": to_date,
            "supplier": 0,
            "item": 0,
            "data_type": "Detailed",
            "barcode_no": "",
            "report_type": "Excel",
            "item_group": 0,
            "grn_type": "",
            "created_from": "",
            "created_to": "",
        },
        headers=HEADERS,
    )
    return resp.content


def parse_grn_register(file_bytes):
    book = xlrd.open_workbook(file_contents=file_bytes, ignore_workbook_corruption=True)
    sheet = book.sheet_by_index(0)
    rows = [sheet.row_values(r) for r in range(sheet.nrows)]
    raw = pd.DataFrame(rows)
    raw = raw.replace("", pd.NA)

    header_row_idx = None
    for i in range(len(raw)):
        row_values = raw.iloc[i].astype(str).tolist()
        if "GRN No" in row_values:
            header_row_idx = i
            break

    if header_row_idx is None:
        return pd.DataFrame()

    columns = raw.iloc[header_row_idx].tolist()
    data = raw.iloc[header_row_idx + 1:].reset_index(drop=True)
    data.columns = columns

    stop_idx = len(data)
    for i in range(len(data)):
        row_values = data.iloc[i].astype(str).tolist()
        if "TOTAL" in row_values:
            stop_idx = i
            break

    data = data.iloc[:stop_idx]
    data = data.dropna(how="all")
    return data


# ---------- Main ----------

if __name__ == "__main__":
    conn = get_connection()

    removed = delete_excluded_grn_types(conn)
    if removed:
        print(f"🧹 Removed {removed} Sample GRN line(s) saved by earlier runs - Sample GRNs are no longer kept.\n")

    last_pulled = get_last_pulled_date(conn)
    from_date = last_pulled if last_pulled else BASELINE_DATE
    to_date = datetime.now().strftime("%d/%m/%Y")

    print(f"Pulling from {from_date} to {to_date} (overlap on {from_date} is expected and safe)...\n")

    total_inserted = 0
    total_samples_skipped = 0

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
            print(f"  {name} / {unit_name}: fetching {len(chunks)} weekly chunk(s) from {from_date} to {to_date}...")

            unit_total_fetched = 0
            unit_total_new = 0
            unit_samples_skipped = 0

            for chunk_from, chunk_to in chunks:
                file_bytes = download_grn_register(session, plant, showroom, chunk_from, chunk_to)
                df = parse_grn_register(file_bytes)
                df, samples_skipped = drop_excluded_grn_types(df)
                unit_samples_skipped += samples_skipped
                total_samples_skipped += samples_skipped

                if df.empty:
                    continue

                df.insert(0, "Company", name)
                inserted = insert_rows(conn, df)
                unit_total_fetched += len(df)
                unit_total_new += inserted
                total_inserted += inserted
                time.sleep(0.3)

            print(f"    {unit_total_fetched} rows fetched, {unit_total_new} genuinely new"
                  f"{f', {unit_samples_skipped} Sample GRN line(s) skipped' if unit_samples_skipped else ''}.")

    set_last_pulled_date(conn, to_date)
    conn.close()

    print(f"\n✅ Done. {total_inserted} new rows added, {total_samples_skipped} Sample GRN line(s) skipped. "
          f"Next run will start from {to_date}.")
    if removed or total_samples_skipped:
        print("👉 Now run Weighted Price Master so purchase prices are recalculated without Sample GRNs.")
