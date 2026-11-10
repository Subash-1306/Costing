"""
Downloads the Rejection Register from the ERP for ALL units, for one month.
Standalone - run it by hand; it doesn't touch any database.

Output (in the DOWNLOADS folder next to SCRIPTS):
    Rejection_Register_2026_08\\
        Rejection_Register_2026_08_ALL_UNITS.xlsx   <- every unit in one sheet, with a Unit column
        U1.xls, U2.xls, ...                         <- each unit's file exactly as the ERP sent it

ONE-TIME SETUP - the ERP needs this report's own ID and settings, so copy
the download link from your browser once:
    1. Open the Rejection Register report in the ERP, pick any unit and dates.
    2. Press F12 and open the "Network" tab.
    3. Click the Excel download.
    4. In the Network list, find the request containing "getDownloadReportFile",
       right-click it > Copy > Copy URL.
    5. Paste it between the quotes of REPORT_URL below and save this file.
The script keeps everything from that link and only changes the unit and
the dates, for each unit and the month you ask for.

Usage (from the SCRIPTS folder):
    python Rejection_Register.py --month 08/2026
"""

import argparse
import calendar
import io
import os
import sys
import time
from urllib.parse import urlparse, parse_qsl

import openpyxl
import openpyxl.styles
import pandas as pd

import Sales_Register as sr   # reuse the same login and unit list as the other scripts

# Paste the copied download link here (see ONE-TIME SETUP above):
REPORT_URL = ""

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(SCRIPT_DIR, "..", "DOWNLOADS")
XLS_ROW_LIMIT = 65536   # the ERP's old .xls format can't hold more rows than this


def report_params_from_url(url):
    """The copied link's settings, ready to reuse. Stops with a clear message
    if it isn't a report download link or doesn't have the expected fields."""
    query = dict(parse_qsl(urlparse(url.strip()).query, keep_blank_values=True))
    if query.get("a") != "getDownloadReportFile" or "report_id" not in query:
        sys.exit("❌ That link isn't a report download link - it should contain "
                 "'a=getDownloadReportFile' and 'report_id='. Please copy it again (see the top of this file).")
    missing = [k for k in ("plant", "showroom", "from_date", "to_date") if k not in query]
    if missing:
        sys.exit(f"❌ The link has no {', '.join(missing)} setting, so the unit/dates can't be changed safely.\n"
                 f"   Settings it does have: {', '.join(query)}\n   Send me this message and I'll adjust the script.")
    return query


def month_range(text):
    """'08/2026' or '2026_08' -> ('01/08/2026', '31/08/2026', '2026_08')"""
    text = text.strip()
    if "/" in text:
        mon, year = text.split("/")
    else:
        year, mon = text.split("_")
    year, mon = int(year), int(mon)
    last = calendar.monthrange(year, mon)[1]
    return f"01/{mon:02d}/{year}", f"{last:02d}/{mon:02d}/{year}", f"{year}_{mon:02d}"


def read_report(file_bytes):
    """Reads an ERP export (.xls or .xlsx) into a DataFrame. Returns
    (DataFrame, None) or (None, reason it couldn't be read)."""
    head = file_bytes[:8]
    try:
        if head.startswith(b"PK"):
            raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl")
        elif head.startswith(b"\xd0\xcf\x11\xe0"):
            import xlrd
            book = xlrd.open_workbook(file_contents=file_bytes, ignore_workbook_corruption=True)
            sheet = book.sheet_by_index(0)
            raw = pd.DataFrame([sheet.row_values(r) for r in range(sheet.nrows)])
        else:
            return None, "the ERP sent back a page instead of an Excel file (not logged in, or wrong settings)"
    except Exception as e:
        return None, f"couldn't read the file ({e})"

    raw = raw.replace("", pd.NA).dropna(how="all").reset_index(drop=True)
    if raw.empty:
        return pd.DataFrame(), None
    # The header is the first row as wide as the widest row - report title
    # rows above it are narrower (usually a single cell).
    widths = raw.notna().sum(axis=1)
    header_idx = int(widths[widths == widths.max()].index[0])
    data = raw.iloc[header_idx + 1:].copy()
    data.columns = [str(c).strip() if pd.notna(c) else f"Column {i+1}" for i, c in enumerate(raw.iloc[header_idx])]
    data = data.dropna(how="all")
    # Drop subtotal/total lines the report adds at the bottom
    is_total = data.astype(str).apply(lambda r: r.str.strip().str.upper().isin(["TOTAL", "GRAND TOTAL"]).any(), axis=1)
    return data[~is_total].reset_index(drop=True), None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--month", required=True, help="e.g. 08/2026")
    ap.add_argument("--url", help="The copied download link (instead of pasting it into REPORT_URL)")
    args = ap.parse_args()

    url = args.url or REPORT_URL
    if not url:
        sys.exit("❌ No report link yet. See ONE-TIME SETUP at the top of Rejection_Register.py "
                 "(copy the Excel download link from your browser once).")
    base = report_params_from_url(url)
    from_date, to_date, key = month_range(args.month)

    out_dir = os.path.join(DOWNLOAD_DIR, f"Rejection_Register_{key}")
    os.makedirs(out_dir, exist_ok=True)
    print(f"Rejection Register, {from_date} to {to_date} (report {base['report_id']})")

    combined, problems = [], []
    for company in sr.COMPANIES:
        session = sr.login(company["client_company"])
        if not session:
            problems.append(f"{company['name']}: login failed - all its units skipped")
            continue
        for unit in company["units"]:
            params = dict(base, plant=company["plant"], showroom=unit["showroom"],
                          from_date=from_date, to_date=to_date)
            content = session.get(sr.DATA_URL, params=params, headers=sr.HEADERS).content
            ext = ".xlsx" if content[:2] == b"PK" else ".xls"
            df, problem = read_report(content)
            if problem:
                problems.append(f"{unit['name']}: {problem}")
                print(f"  {unit['name']}: ❌ {problem}")
                continue
            with open(os.path.join(out_dir, unit["name"] + ext), "wb") as f:
                f.write(content)
            if ext == ".xls" and len(df) >= XLS_ROW_LIMIT - 100:
                problems.append(f"{unit['name']}: {len(df)} rows - the ERP's .xls limit, so rows may be cut off")
            print(f"  {unit['name']}: {len(df)} rows")
            if len(df):
                df.insert(0, "Unit", unit["name"])
                df.insert(0, "Company", company["name"])
                combined.append(df)
            time.sleep(0.5)

    if combined:
        all_units = pd.concat(combined, ignore_index=True)
        path = os.path.join(out_dir, f"Rejection_Register_{key}_ALL_UNITS.xlsx")
        all_units.to_excel(path, index=False)
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True)
        wb.save(path)
        print(f"\n✅ {len(all_units)} rows from {len(combined)} unit(s) -> {path}")
    else:
        print("\nNo rejection rows found for any unit this month.")

    if problems:
        print("\n⚠️  Please check:")
        for p in problems:
            print("   ", p)


if __name__ == "__main__":
    main()
