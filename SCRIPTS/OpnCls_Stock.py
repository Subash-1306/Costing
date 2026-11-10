"""
Monthly Opening/Closing Stock sync - pulls the Stock Register for a specific
month (you set TARGET_YEAR/TARGET_MONTH below before each run), filtered to
RMI items only, and saves into Opening_Closing_Stock.db.

Each row is keyed by (Company, Unit, Warehouse, Item Code, Month) - so
re-running the same month safely OVERWRITES that month's numbers instead of
duplicating, which is handy if you need to re-pull a month for a correction.

BEFORE RUNNING: set TARGET_YEAR and TARGET_MONTH below to the month you want.

Run: python sync_opening_closing_stock.py
"""

import requests
import pandas as pd
import sqlite3
import io
import calendar

BASE_URL = "https://unitech.tarka.in"
LOGIN_URL = f"{BASE_URL}/app.php"
DATA_URL = f"{BASE_URL}/app.php"

# ERP login comes from the dashboard (or is asked when run by hand) - nothing is saved in this file
from erp_credentials import get_credentials

HEADERS = {
    "X-Requested-With": "XMLHttpRequest",
    "Referer": f"{BASE_URL}/MainPage.html",
}

# TARGET_YEAR and TARGET_MONTH are asked interactively when you run the
# script (see the bottom of the file) - no need to edit this file each time.

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
DB_FILE = os.path.join(DATABASE_DIR, "Opening_Closing_Stock.db")

COLUMN_MAP = {
    "Company": "company",
    "UNIT": "unit",
    "W/H": "warehouse",
    "W/H Desc.": "warehouse_desc",
    "ERP CODE": "erp_code",
    "ITEM CODE": "item_code",
    "ITEM NAME": "item_name",
    "UOM": "uom",
    "ITEM GROUP": "item_group",
    "ITEM SUB GROUP": "item_sub_group",
    "Month": "month",
    "Opening Qty.": "opening_qty",
    "Closing Qty.": "closing_qty",
}

SQL_COLUMNS = list(COLUMN_MAP.values())


def get_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.execute(f"""
        CREATE TABLE IF NOT EXISTS opening_closing_stock (
            {", ".join(f"{col} TEXT" for col in SQL_COLUMNS)},
            PRIMARY KEY (company, unit, warehouse, item_code, month)
        )
    """)
    conn.commit()
    return conn


def upsert_rows(conn, df):
    cur = conn.cursor()
    count = 0
    for _, row in df.iterrows():
        values = [row.get(excel_col) for excel_col in COLUMN_MAP.keys()]
        values = [None if pd.isna(v) else str(v) for v in values]

        placeholders = ", ".join("?" for _ in SQL_COLUMNS)
        cur.execute(
            f"INSERT OR REPLACE INTO opening_closing_stock ({', '.join(SQL_COLUMNS)}) "
            f"VALUES ({placeholders})",
            values,
        )
        count += 1
    conn.commit()
    return count


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


def download_stock_register(session, plant, showroom, from_date, to_date):
    resp = session.get(
        DATA_URL,
        params={
            "a": "getDownloadReportFile",
            "report_id": 915,
            "schema_name": "Stock~Register",
            "system_module_table_name": "stock_register",
            "plant": plant,
            "showroom": showroom,
            "from_date": from_date,
            "to_date": to_date,
            "store": 0,
            "item_group": 0,   # fetch all groups, filter locally to avoid
                               # relying on a group-ID that might differ per company
            "report_type": "Excel",
        },
        headers=HEADERS,
    )
    return resp.content


def parse_stock_register(file_bytes, company_name, month_key):
    raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl")

    header_row_idx = None
    for i in range(len(raw)):
        row_values = raw.iloc[i].astype(str).tolist()
        if "Opening Qty." in row_values:
            header_row_idx = i
            break

    if header_row_idx is None:
        return pd.DataFrame()

    columns = raw.iloc[header_row_idx].tolist()
    data = raw.iloc[header_row_idx + 1:].reset_index(drop=True)
    data.columns = columns
    data = data.dropna(how="all")

    if "ITEM CODE" in data.columns:
        data = data[data["ITEM CODE"].notna()]

    # Filters: Raw Material group, RMI items only (excludes RMP codes)
    data = data[data["ITEM GROUP"].astype(str).str.upper() == "RAW MATERIAL"]
    data = data[data["ITEM CODE"].astype(str).str.contains("RMI")]

    data.insert(0, "Company", company_name)
    data["Month"] = month_key

    return data


if __name__ == "__main__":
    import sys

    if len(sys.argv) >= 3:
        # Called with arguments (e.g. from the dashboard): python OpnCls_Stock.py 2026 9
        TARGET_YEAR = int(sys.argv[1])
        TARGET_MONTH = int(sys.argv[2])
    else:
        # Run directly from the terminal - ask interactively
        print("Which month's stock register do you want to pull?")
        TARGET_YEAR = int(input("Year (e.g. 2026): ").strip())
        TARGET_MONTH = int(input("Month number (1-12): ").strip())

    from_date = f"01/{TARGET_MONTH:02d}/{TARGET_YEAR}"
    last_day = calendar.monthrange(TARGET_YEAR, TARGET_MONTH)[1]
    to_date = f"{last_day:02d}/{TARGET_MONTH:02d}/{TARGET_YEAR}"
    month_key = f"{TARGET_YEAR}_{TARGET_MONTH:02d}"

    print(f"Pulling stock register for {from_date} to {to_date} (month key: {month_key})\n")

    conn = get_connection()
    total_rows = 0

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

            print(f"  Downloading {name} / {unit_name}...")
            file_bytes = download_stock_register(session, plant, showroom, from_date, to_date)
            df = parse_stock_register(file_bytes, name, month_key)

            if df.empty:
                print(f"    No matching RMI rows for {unit_name}.")
                continue

            count = upsert_rows(conn, df)
            total_rows += count
            print(f"    ✅ {count} rows saved.")

    conn.close()
    print(f"\n✅ Done. {total_rows} total rows saved for {month_key}.")
