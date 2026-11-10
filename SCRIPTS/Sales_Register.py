"""
Incremental Sales Register sync.

- Normal run (no arguments): pulls everything from the last recorded date
  through today (the overlap on that boundary day catches invoices added
  later, e.g. after your last run at 11am). Duplicate-safe: each row gets a
  unique key (hash of Invoice No + FG Code + Date + Qty), so re-fetching
  the same day never creates duplicate rows - only genuinely new invoices
  get added.

Run this anytime (once a day, several times a day, whatever) - it always
safely tops up just the new stuff.

Only genuine Finished Goods rows are kept - FG Code must contain "FG" but
NOT "SFG" (semi-finished goods also contain the substring "FG"). Sales of
SFG, RMI, or any other non-FG item are excluded entirely. Note: the ERP
export has a separate "Item Code" column too, which is the CUSTOMER's own
part number - not ours - and is kept only as a reference field.

Run: python Sales_Register.py
"""

import requests
import pandas as pd
import sqlite3
import io
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

BASELINE_DATE = "01/06/2026"   # dd/mm/yyyy - first-ever run starts from here

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
from costing_config import database_dir   # set in costing\config\settings.json
DATABASE_DIR = database_dir()
os.makedirs(DATABASE_DIR, exist_ok=True)
DB_FILE = os.path.join(DATABASE_DIR, "Sales_Register.db")

# Excel column name -> SQL column name (sanitized)
COLUMN_MAP = {
    "Unit": "unit",
    "Invoice Type": "invoice_type",
    "Invoice Date": "invoice_date",
    "Invoice Time": "invoice_time",
    "Invoice No": "invoice_no",
    "Vehicle No": "vehicle_no",
    "Customer": "customer",
    "GST No": "gst_no",
    "Item Code": "customer_item_code",   # the CUSTOMER's own part number, not ours - kept for reference only
    "Item Name": "item_name",
    "FG Code": "item_code",   # our real FG/SFG code - this is what every other table (production, item master, weighted price) matches on
    "HSN Code": "hsn_code",
    "UOM": "uom",
    "Quantity": "quantity",
    "Rate": "rate",
    "Price Type": "price_type",
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
    "REMARKS": "remarks",
    "CAR MODEL": "car_model",
    "E-invoice No": "einvoice_no",
    "Ewaybill No": "ewaybill_no",
}

SQL_COLUMNS = ["company"] + list(COLUMN_MAP.values())


# ---------- Database ----------

def get_connection():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS sales_register (
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
                        ["Invoice No", "FG Code", "Invoice Date", "Quantity"])
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
            f"INSERT OR IGNORE INTO sales_register (row_key, {', '.join(SQL_COLUMNS)}) "
            f"VALUES (?, {placeholders})",
            [row_key] + values,
        )
        if cur.rowcount > 0:
            inserted += 1
    conn.commit()
    return inserted


# ---------- ERP calls ----------

def daterange_chunks(from_date_str, to_date_str, chunk_days=7):
    """
    Break a date range into smaller chunks (default: weekly). The ERP's .xls
    export uses the legacy Excel format with a 65,536-row limit, and large
    date ranges for busy units can exceed that and corrupt the file. Smaller
    chunks keep each downloaded file safely small.
    """
    start = datetime.strptime(from_date_str, "%d/%m/%Y")
    end = datetime.strptime(to_date_str, "%d/%m/%Y")
    chunks = []
    current = start
    while current <= end:
        chunk_end = min(current + timedelta(days=chunk_days - 1), end)
        chunks.append((current.strftime("%d/%m/%Y"), chunk_end.strftime("%d/%m/%Y")))
        current = chunk_end + timedelta(days=1)
    return chunks


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


def download_sales_register(session, plant, showroom, from_date, to_date):
    resp = session.get(
        DATA_URL,
        params={
            "a": "getDownloadReportFile",
            "report_id": 193,
            "schema_name": "Sales~Register",
            "system_module_table_name": "sales_register",
            "plant": plant,
            "showroom": showroom,
            "from_date": from_date,
            "to_date": to_date,
            "customer": 0,
            "item": 0,
            "item_group": 0,
            "item_sub_group": 0,
            "data_type": "Detailed",
            "report_type": "Excel",
            "from_time": "",
            "to_time": "",
        },
        headers=HEADERS,
    )
    return resp.content


def parse_sales_register(file_bytes):
    # Use xlrd directly (not via pandas' wrapper) so we can pass
    # ignore_workbook_corruption=True - this ERP's .xls export is slightly
    # malformed and xlrd's default strict parsing rejects it otherwise.
    book = xlrd.open_workbook(file_contents=file_bytes, ignore_workbook_corruption=True)
    sheet = book.sheet_by_index(0)
    rows = [sheet.row_values(r) for r in range(sheet.nrows)]
    raw = pd.DataFrame(rows)
    raw = raw.replace("", pd.NA)  # xlrd gives '' for empty cells; normalize to NaN

    header_row_idx = None
    for i in range(len(raw)):
        row_values = raw.iloc[i].astype(str).tolist()
        if "Invoice No" in row_values:
            header_row_idx = i
            break

    if header_row_idx is None:
        # This used to fail SILENTLY, returning an empty result indistinguishable
        # from "genuinely no sales that week" - which is exactly what made U2/U3/U4
        # showing 0 rows across 3+ months impossible to diagnose. Now it prints
        # exactly what it actually saw, so the next run reveals the real cause.
        preview = raw.iloc[0].tolist()[:6] if len(raw) else []
        print(f"    ⚠️  Could not find 'Invoice No' header ({len(raw)} row(s) in file). First row seen: {preview}")
        return pd.DataFrame()

    columns = raw.iloc[header_row_idx].tolist()
    data = raw.iloc[header_row_idx + 1:].reset_index(drop=True)
    data.columns = columns
    rows_after_header = len(data)

    # NOTE: we used to stop at the first row containing "TOTAL" anywhere,
    # assuming it marked the file's single grand total at the end. Real
    # invoice reports commonly have a subtotal row (also containing "TOTAL")
    # after EVERY customer/date group, well before the actual end of the
    # data - so that logic was truncating the file after the first such
    # group, sometimes almost immediately. Removed: the Item Code filter
    # below already excludes summary/subtotal rows on its own, since they
    # never have a valid FG-pattern item code.
    data = data.dropna(how="all")
    rows_after_dropna = len(data)

    # Only keep genuine Finished Goods rows - "FG Code" must contain "FG"
    # but NOT "SFG" (semi-finished goods also contain the substring "FG",
    # e.g. SFGA2649, so a plain "contains FG" check would wrongly include them).
    # IMPORTANT: this ERP's export has TWO separate code columns - "Item Code"
    # is the CUSTOMER's own part number (e.g. '85873-BV000NNB'), which almost
    # never contains "FG" and is NOT what we want; "FG Code" is our own real
    # code (FG6906, SFGP1408, etc.) and is the one to filter and key on.
    if "FG Code" in data.columns:
        codes = data["FG Code"].astype(str)
        data = data[codes.str.contains("FG") & ~codes.str.contains("SFG")]
    else:
        print(f"    ⚠️  No 'FG Code' column found. Columns seen: {list(data.columns)[:12]}")

    print(f"    (diagnostic: {rows_after_header} row(s) after header, {rows_after_dropna} after removing blanks, {len(data)} after FG-code filter)")

    return data


# ---------- Main ----------

if __name__ == "__main__":
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
            print(f"  {name} / {unit_name}: fetching {len(chunks)} weekly chunk(s) from {from_date} to {to_date}...")

            unit_total_fetched = 0
            unit_total_new = 0

            for chunk_from, chunk_to in chunks:
                file_bytes = download_sales_register(session, plant, showroom, chunk_from, chunk_to)
                df = parse_sales_register(file_bytes)

                if df.empty:
                    continue

                df.insert(0, "Company", name)
                inserted = insert_rows(conn, df)
                unit_total_fetched += len(df)
                unit_total_new += inserted
                total_inserted += inserted
                time.sleep(0.3)

            print(f"    {unit_total_fetched} rows fetched, {unit_total_new} genuinely new.")

    set_last_pulled_date(conn, to_date)
    conn.close()

    print(f"\n✅ Done. {total_inserted} new rows added. Next run will start from {to_date}.")
