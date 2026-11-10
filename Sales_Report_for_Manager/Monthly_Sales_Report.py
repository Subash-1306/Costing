"""
Monthly Sales Report - same layout as "August_2026.xlsx".

ONE self-contained file: it logs in to the ERP itself, downloads the month's
Sales Register for every unit, and saves the report in the SAME FOLDER as
this file (or as Sales Report.exe). It needs nothing else from anyone's computer.

It asks for an ERP username and password every time it runs (the password
is hidden as you type and is never saved).

Every customer with sales that month is included, every item code (FG, SFG,
tooling, scrap...). "Actual Sales Value" = total Taxable Value (before GST).

Built to be edited:
  Sheet "Data"  - every downloaded invoice line.
  Report sheet  - each customer's value is a formula that adds up that
                  customer's lines from Data (SUMIFS on unit + customer name),
                  and every total is a formula too. Rename a customer, merge
                  or delete rows, or fix a line in Data, and all the numbers
                  update. Data also shows whether the report total still
                  equals the total of all lines.

Run: double-click this file (or Sales Report.exe), or
     python Monthly_Sales_Report.py --month 08/2026

First time on a new computer (only when running the .py, not the exe):
     pip install requests xlrd openpyxl
"""

import argparse
import calendar
import getpass
import os
import sys
import time
from collections import defaultdict
from datetime import datetime, timedelta

try:
    import openpyxl
    import requests
    import xlrd
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
except ImportError as e:
    print(f"A required package is missing ({e.name}).")
    print("Open Command Prompt and run:   pip install requests xlrd openpyxl")
    input("\nPress Enter to close...")
    sys.exit(1)

# ---------------------------------------------------------------------------
# ERP CONNECTION (the same calls the ERP website makes)
# ---------------------------------------------------------------------------
BASE_URL = "https://unitech.tarka.in"
APP_URL = f"{BASE_URL}/app.php"
HEADERS = {"X-Requested-With": "XMLHttpRequest", "Referer": f"{BASE_URL}/MainPage.html"}

COMPANIES = [
    {"name": "Unitech", "client_company": "9", "plant": "9",
     "units": [{"name": "U1", "showroom": "21"}, {"name": "U2", "showroom": "25"},
               {"name": "U3", "showroom": "29"}, {"name": "U4", "showroom": "26"}]},
    {"name": "Vtech", "client_company": "11", "plant": "11",
     "units": [{"name": "V1", "showroom": "27"}, {"name": "V2", "showroom": "28"}]},
]


def erp_login(client_company, username, password):
    """A logged-in session for one company, or (None, reason)."""
    session = requests.Session()
    try:
        session.get(BASE_URL, timeout=60)
        resp = session.post(APP_URL, params={"a": "validateUsernameAndPassword.demo"},
                            data={"usr": username, "password": password, "client_company": client_company},
                            headers=HEADERS, timeout=60)
        data = resp.json()
    except Exception as e:
        return None, f"couldn't reach the ERP ({e})"
    if data.get("status") == "success":
        return session, None
    return None, str(data.get("message") or data)


def download_sales_register(session, plant, showroom, from_date, to_date):
    resp = session.get(APP_URL, params={
        "a": "getDownloadReportFile", "report_id": 193, "schema_name": "Sales~Register",
        "system_module_table_name": "sales_register", "plant": plant, "showroom": showroom,
        "from_date": from_date, "to_date": to_date, "customer": 0, "item": 0, "item_group": 0,
        "item_sub_group": 0, "data_type": "Detailed", "report_type": "Excel", "from_time": "", "to_time": "",
    }, headers=HEADERS, timeout=600)
    return resp.content


def weekly_chunks(from_date, to_date):
    """The ERP's .xls export can't hold more than 65,536 rows, so a busy
    unit's month is downloaded one week at a time."""
    start, end = datetime.strptime(from_date, "%d/%m/%Y"), datetime.strptime(to_date, "%d/%m/%Y")
    chunks = []
    while start <= end:
        chunk_end = min(start + timedelta(days=6), end)
        chunks.append((start.strftime("%d/%m/%Y"), chunk_end.strftime("%d/%m/%Y")))
        start = chunk_end + timedelta(days=1)
    return chunks


# ---------------------------------------------------------------------------
# REPORT LAYOUT - sections in report order. Each lists EVERY customer of that
# unit with sales in the month (largest first). A unit with no sales is left
# out. Sections with the same "group" get a combined total after the last one.
# ---------------------------------------------------------------------------
SECTIONS = [
    {"unit": "U1", "title": "UNITECH PLASTO COMPONENTS (UNIT-1)", "group": "UNITECH PLASTO COMPONENTS  TOTAL SALE"},
    {"unit": "U2", "title": "UNITECH PLASTO COMPONENTS (UNIT-2)", "group": "UNITECH PLASTO COMPONENTS  TOTAL SALE"},
    {"unit": "U3", "title": "UNITECH PLASTO COMPONENTS (UNIT-3)", "group": "UNITECH PLASTO COMPONENTS  TOTAL SALE"},
    {"unit": "U4", "title": "UNITECH PLASTO COMPONENTS (UNIT-4)", "group": "UNITECH PLASTO COMPONENTS  TOTAL SALE"},
    {"unit": "V1", "title": "VEETECH AUTOMOTIVE -TN", "group": "VEETECH AUTOMOTIVE TOTAL SALES"},
    {"unit": "V2", "title": "VEETECH AUTOMOTIVE -AP", "group": "VEETECH AUTOMOTIVE TOTAL SALES"},
]

# Columns of the Data sheet (the report's formulas point at Unit, Customer and Taxable Value)
DATA_COLUMNS = ["Unit", "Invoice Date", "Invoice No", "Invoice Type", "Customer", "FG Code",
                "Customer Part No", "Item Name", "Quantity", "Taxable Value"]
DATA_UNIT, DATA_CUSTOMER, DATA_VALUE = "A", "E", "J"

if getattr(sys, "frozen", False):
    # Running as Sales Report.exe: __file__ points into a temporary folder that
    # is deleted afterwards, so save next to the exe itself
    SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = SCRIPT_DIR   # the report is saved next to this file


def norm(name):
    """Compare names ignoring capitals and extra spaces."""
    return " ".join(str(name or "").upper().split())


def parse_month(text):
    t = str(text).strip().strip("[]").strip().replace("-", "/").replace(" ", "")
    try:
        if "_" in t:
            year, mon = t.split("_")
        else:
            mon, year = t.split("/")
        if len(year) == 4 and 1 <= int(mon) <= 12:
            return int(mon), int(year)
    except ValueError:
        pass
    return None


def read_lines(file_bytes):
    """Every invoice line in one downloaded Sales Register file (all item
    codes), as dicts. None if the download wasn't a real report."""
    try:
        book = xlrd.open_workbook(file_contents=file_bytes, ignore_workbook_corruption=True)
    except Exception:
        return None
    sheet = book.sheet_by_index(0)
    rows = [[v.strip() if isinstance(v, str) else v for v in sheet.row_values(r)] for r in range(sheet.nrows)]
    header_idx = next((i for i, r in enumerate(rows) if "Invoice No" in r), None)
    if header_idx is None:
        return None
    col = {name: j for j, name in enumerate(rows[header_idx]) if name not in ("", None)}

    def number(v):
        try:
            return float(str(v).replace(",", ""))
        except ValueError:
            return 0.0

    out = []
    for r in rows[header_idx + 1:]:
        get = lambda name: r[col[name]] if name in col and col[name] < len(r) else ""
        invoice = str(get("Invoice No")).strip()
        customer = " ".join(str(get("Customer")).split())   # one spelling, even if the ERP has stray spaces
        if not invoice or not customer or "TOTAL" in invoice.upper():
            continue   # blank lines and the report's subtotal/total rows
        out.append({"invoice_date": get("Invoice Date"), "invoice_no": invoice, "invoice_type": get("Invoice Type"),
                    "customer": customer, "fg_code": get("FG Code"), "part_no": get("Item Code"),
                    "item_name": get("Item Name"), "qty": number(get("Quantity")), "value": number(get("Taxable Value"))})
    return out


def download_month(month, year, username, password):
    """{unit: [(customer, value, code, invoice_no), ...]} for every ERP unit, and a list of problems."""
    last = calendar.monthrange(year, month)[1]
    from_date, to_date = f"01/{month:02d}/{year}", f"{last:02d}/{month:02d}/{year}"
    by_unit, problems = defaultdict(list), []
    for company in COMPANIES:
        session, why = erp_login(company["client_company"], username, password)
        if not session:
            problems.append(f"{company['name']}: login failed ({why}) - its units are missing from the report")
            print(f"  {company['name']}: ❌ login failed - {why}")
            continue
        for unit in company["units"]:
            for c_from, c_to in weekly_chunks(from_date, to_date):
                try:
                    content = download_sales_register(session, company["plant"], unit["showroom"], c_from, c_to)
                except Exception as e:
                    content = b""
                    print(f"    {unit['name']} {c_from}: {e}")
                lines = read_lines(content)
                if lines is None:
                    problems.append(f"{unit['name']} {c_from} - {c_to}: download failed")
                else:
                    by_unit[unit["name"]].extend(lines)
                time.sleep(0.5)
            print(f"  {company['name']} / {unit['name']}: {len(by_unit[unit['name']])} invoice lines")
    return by_unit, problems


def write_report(month, year, by_unit, problems, path):
    # Colours and fonts taken from August_2026.xlsx
    ORANGE, DARK_GREY, LIGHT_GREY, PALE_GREY, MID_GREY = "ED7D31", "7C7C7C", "C9C9C9", "DBDBDB", "A5A5A5"
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    money = "#,##0.00"
    def font(name="Times New Roman", size=11, bold=False, white=False):
        return Font(name=name, size=size, bold=bold, color="FFFFFF" if white else "000000")
    def fill(hex_colour):
        return PatternFill("solid", start_color=hex_colour, end_color=hex_colour)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = datetime(year, month, 1).strftime("%B %Y")
    report = f"'{ws.title}'"   # quoted - the sheet name has a space
    ws.column_dimensions["A"].width = 7.8
    ws.column_dimensions["B"].width = 66.7
    ws.column_dimensions["C"].width = 17.5
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1   # whole report on one A4 page, like the August file
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.3

    def band(row, text, cols, fnt, colour=None, fmt=None, align="center"):
        first, last = cols
        if first != last:
            ws.merge_cells(f"{first}{row}:{last}{row}")
        c = ws[f"{first}{row}"]
        c.value = text
        c.font, c.alignment = fnt, Alignment(horizontal=align, vertical="center")
        if fmt:
            c.number_format = fmt
        for col in range(ord(first) - 64, ord(last) - 63):
            cell = ws.cell(row, col)
            cell.border = box
            if colour:
                cell.fill = fill(colour)

    band(1, datetime(year, month, 1).strftime("%b-%y") + " Month Sales Details ", "AC", font(size=20, bold=True, white=True), ORANGE)
    band(2, datetime(year, month, 1), "AC", font(size=18, bold=True, white=True), ORANGE, fmt="mmm-yy")
    ws.row_dimensions[1].height, ws.row_dimensions[2].height = 24.6, 22.8

    # Customers per unit, largest first (the order only - the values are formulas)
    totals = defaultdict(lambda: defaultdict(float))
    for unit, lines in by_unit.items():
        for l in lines:
            totals[unit][l["customer"]] += l["value"]

    shown = [sec for sec in SECTIONS if totals.get(sec["unit"])]
    row, group_totals, grand_cells = 3, defaultdict(list), []
    for i, sec in enumerate(shown):
        veetech = "VEETECH" in sec["group"].upper()
        band(row, sec["title"], "AC", font(bold=True, white=True), ORANGE if veetech else DARK_GREY)
        row += 1
        for col, head in zip("ABC", ("SL.NO.", "Customer Name", "Actual Sales Value")):
            band(row, head, col * 2, font(bold=True), ORANGE if veetech else LIGHT_GREY)
        row += 1
        first = row
        customers = sorted(totals[sec["unit"]].items(), key=lambda kv: -kv[1])
        for n, (cust, _) in enumerate(customers, start=1):
            a_, b_ = ws.cell(row, 1, n), ws.cell(row, 2, cust)
            a_.font, a_.alignment = font(size=9), Alignment(horizontal="center")
            b_.font, b_.alignment = font(size=9), Alignment(horizontal="left")
            # This customer's lines from Data: same unit, same name (change the name and it follows)
            c = ws.cell(row, 3, f'=SUMIFS(Data!${DATA_VALUE}:${DATA_VALUE},Data!${DATA_UNIT}:${DATA_UNIT},"{sec["unit"]}",'
                                f'Data!${DATA_CUSTOMER}:${DATA_CUSTOMER},$B{row})')
            c.font, c.number_format = Font(name="Calibri", size=11), money
            for col in range(1, 4):
                ws.cell(row, col).border = box
            row += 1
        band(row, "Customer Total", "AB", Font(name="Calibri", size=11, bold=True))
        tot = ws.cell(row, 3, f"=SUM(C{first}:C{row - 1})")
        tot.font, tot.number_format, tot.border = Font(name="Calibri", size=11, bold=True), money, box
        group_totals[sec["group"]].append(f"C{row}")
        row += 1
        if i == len(shown) - 1 or shown[i + 1]["group"] != sec["group"]:
            band(row, sec["group"], "AB", font(size=10, bold=True), PALE_GREY)
            g = ws.cell(row, 3, "=" + "+".join(group_totals[sec["group"]]))
            g.font, g.number_format, g.border, g.fill = Font(name="Calibri", size=12, bold=True), money, box, fill(MID_GREY)
            grand_cells.append(f"C{row}")
            row += 1
        row += 1   # blank line between sections

    # ---- Data sheet: every invoice line the report adds up ----
    data = wb.create_sheet("Data")
    data.append(DATA_COLUMNS)
    for unit in [s_["unit"] for s_ in SECTIONS] + sorted(set(by_unit) - {s_["unit"] for s_ in SECTIONS}):
        for l in by_unit.get(unit, []):
            data.append([unit, l["invoice_date"], l["invoice_no"], l["invoice_type"], l["customer"], l["fg_code"],
                         l["part_no"], l["item_name"], l["qty"], l["value"]])
    last = data.max_row
    for cell in data[1]:
        cell.font = Font(bold=True)
    for col, w in zip("ABCDEFGHIJ", (6, 12, 22, 14, 50, 12, 18, 40, 10, 16)):
        data.column_dimensions[col].width = w
    for r in data.iter_rows(min_row=2, min_col=9, max_col=10):
        r[0].number_format, r[1].number_format = "#,##0.##", money
    data.freeze_panes = "A2"
    data.auto_filter.ref = f"A1:J{max(last, 2)}"

    # Check: does the report still add up to every line in Data?
    data["L1"], data["M1"] = "Total of all lines above", f"=SUM({DATA_VALUE}2:{DATA_VALUE}1048576)"
    data["L2"], data["M2"] = "Total shown in the report", "=" + ("+".join(f"{report}!{c}" for c in grand_cells) or "0")
    data["L3"], data["M3"] = "Difference (should be 0)", "=M1-M2"
    for r_ in (1, 2, 3):
        data[f"L{r_}"].font = Font(bold=True)
        data[f"M{r_}"].number_format = money
        data[f"M{r_}"].font = Font(bold=True)
    data.column_dimensions["L"].width, data.column_dimensions["M"].width = 26, 18

    if problems:
        p = wb.create_sheet("Download problems")
        p.append(["These downloads failed, so the report may be missing sales. Run it again."])
        p["A1"].font = Font(bold=True, color="C00000")
        for x in problems:
            p.append([x])
        p.column_dimensions["A"].width = 90

    wb.calculation.fullCalcOnLoad = True   # Excel works out every formula when the file opens
    wb.save(path)
    return [s_["unit"] for s_ in shown]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--month", help="e.g. 08/2026")
    ap.add_argument("--user", help="ERP username (the password is always asked, never stored)")
    a = ap.parse_args()
    parsed = parse_month(a.month) if a.month else None
    interactive = a.month is None
    while not parsed:
        parsed = parse_month(input("Month (mm/yyyy), e.g. 08/2026: "))
        if not parsed:
            print("   Please type the month like 08/2026")
    month, year = parsed

    username = a.user or input("ERP username: ").strip()
    password = getpass.getpass("ERP password (hidden as you type): ")

    # Check the login once before downloading everything
    session, why = erp_login(COMPANIES[0]["client_company"], username, password)
    if not session:
        print(f"\n❌ Couldn't log in: {why}\nCheck the username and password and run it again.")
        if interactive:
            input("\nPress Enter to close...")
        return

    print(f"\nDownloading {calendar.month_name[month]} {year} sales from the ERP (all item codes)...")
    by_unit, problems = download_month(month, year, username, password)

    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.abspath(os.path.join(OUT_DIR, f"Sales_Report_{calendar.month_name[month]}_{year}.xlsx"))
    units = write_report(month, year, by_unit, problems, path)

    print(f"\n✅ Report saved: {path}")
    print(f"   Units: {', '.join(units) or 'no sales found'} | every value and total is a formula over the 'Data' sheet")
    if problems:
        print("⚠️  Some downloads failed - see the 'Download problems' sheet. Run it again.")
    if interactive:
        input("\nPress Enter to close...")


if __name__ == "__main__":
    main()
