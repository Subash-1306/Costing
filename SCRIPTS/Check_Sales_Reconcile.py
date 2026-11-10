"""
Sales Register reconciliation check - READ ONLY, changes nothing.

Explains why a customer's total in the dashboard differs from the ERP.
It downloads that customer's sales fresh from the ERP for a date range,
then splits the ERP's total into exactly where every rupee went:

  1. Non-FG lines   - excluded on purpose (SFG, tooling, packing, blank FG code...)
  2. Duplicate-key  - genuine lines dropped because they share Invoice No +
                      FG Code + Date + Quantity with another line
  3. Missing weeks  - lines the ERP has but the database doesn't (e.g. a
                      weekly download that failed and was never re-fetched)

Usage (from the SCRIPTS folder):
    python Check_Sales_Reconcile.py --customer "LS AUTO" --from 01/06/2026 --to 31/08/2026
    python Check_Sales_Reconcile.py --customer "LS AUTO" --from 01/07/2026 --to 31/07/2026 --unit U1

--customer matches any customer name CONTAINING that text (not case-sensitive),
so "LS AUTO" catches "LS AUTOMOTIVE INDIA PVT LTD" and any spelling variants.
Use the same date range as the ERP total you're comparing against.
"""

import argparse
import os
import sqlite3
import time
from collections import defaultdict

import pandas as pd
import xlrd

import Sales_Register as sr   # reuse its login, download, key and settings exactly


def to_num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def parse_all_lines(file_bytes):
    """Same reading as Sales_Register.parse_sales_register, but WITHOUT the
    FG-only filter, so non-FG lines can be counted too. Keeps only real
    invoice lines (an Invoice No and a Customer), which leaves out the
    report's subtotal/total rows."""
    try:
        book = xlrd.open_workbook(file_contents=file_bytes, ignore_workbook_corruption=True)
    except Exception:
        return None   # the ERP sent back an error page instead of a file
    sheet = book.sheet_by_index(0)
    raw = pd.DataFrame([sheet.row_values(r) for r in range(sheet.nrows)]).replace("", pd.NA)

    header_idx = None
    for i in range(len(raw)):
        if "Invoice No" in raw.iloc[i].astype(str).tolist():
            header_idx = i
            break
    if header_idx is None:
        return None   # download failed / not a real report

    data = raw.iloc[header_idx + 1:].reset_index(drop=True)
    data.columns = raw.iloc[header_idx].tolist()
    data = data.dropna(how="all")
    inv = data["Invoice No"].astype(str)
    data = data[data["Invoice No"].notna() & data["Customer"].notna() & ~inv.str.upper().str.contains("TOTAL")]
    return data


def is_fg(code):
    code = str(code)
    return "FG" in code and "SFG" not in code


def analyse(erp_lines, db_rows, customer_text):
    """
    erp_lines: DataFrame of ERP invoice lines (with a 'Company' column added)
    db_rows:   list of (unit, invoice_date, taxable_value) from sales_register
    Returns a dict with every bucket, so it can be printed or tested.
    """
    needle = customer_text.upper()
    lines = erp_lines[erp_lines["Customer"].astype(str).str.upper().str.contains(needle, regex=False)].copy()
    lines["_value"] = lines["Taxable Value"].map(to_num)

    erp_total = lines["_value"].sum()

    fg_mask = lines["FG Code"].map(is_fg) if "FG Code" in lines.columns else pd.Series(False, index=lines.index)
    non_fg = lines[~fg_mask]
    fg = lines[fg_mask]

    # Reproduce Sales_Register's exact duplicate key; the first line with a
    # given key is kept, every later one is what INSERT OR IGNORE drops
    seen, kept_idx, dropped_idx = set(), [], []
    for idx, row in fg.iterrows():
        key = sr.make_row_key(row["Company"], row)
        (dropped_idx if key in seen else kept_idx).append(idx)
        seen.add(key)
    dropped = fg.loc[dropped_idx]
    kept = fg.loc[kept_idx]

    # What the database SHOULD hold for this customer vs what it DOES hold,
    # per unit and week - pinpoints any missing weeks
    def week_of(date_text):
        try:
            d = pd.to_datetime(date_text, format="%d/%m/%Y")
            return (d - pd.Timedelta(days=d.weekday())).strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            return "?"

    expected = defaultdict(float)
    for _, row in kept.iterrows():
        expected[(str(row["Unit"]), week_of(row["Invoice Date"]))] += row["_value"]
    actual = defaultdict(float)
    for unit, date, value in db_rows:
        actual[(str(unit), week_of(date))] += to_num(value)
    week_gaps = []
    for k in sorted(set(expected) | set(actual)):
        diff = expected.get(k, 0.0) - actual.get(k, 0.0)
        if abs(diff) >= 1:
            week_gaps.append((k[0], k[1], expected.get(k, 0.0), actual.get(k, 0.0), diff))

    return {
        "erp_total": erp_total,
        "non_fg_total": non_fg["_value"].sum(),
        "non_fg_by_code": non_fg.groupby(non_fg["FG Code"].fillna("(blank)").astype(str))["_value"].sum()
                                .sort_values(ascending=False) if len(non_fg) else pd.Series(dtype=float),
        "non_fg_names": non_fg.groupby(non_fg["FG Code"].fillna("(blank)").astype(str))["Item Name"].first()
                              if len(non_fg) and "Item Name" in non_fg.columns else {},
        "dup_total": dropped["_value"].sum(),
        "dup_lines": dropped,
        "expected_db_total": kept["_value"].sum(),
        "actual_db_total": sum(to_num(v) for _, _, v in db_rows),
        "week_gaps": week_gaps,
        "line_count": len(lines),
    }


def money(x):
    return f"Rs {x:,.2f}"


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--customer", required=True, help='Text in the customer name, e.g. "LS AUTO"')
    ap.add_argument("--from", dest="from_date", required=True, help="dd/mm/yyyy")
    ap.add_argument("--to", dest="to_date", required=True, help="dd/mm/yyyy")
    ap.add_argument("--unit", help="Only this unit, e.g. U1 (default: all units)")
    args = ap.parse_args()
    only_unit = args.unit.strip().upper() if args.unit else None

    # ---- 1. Fresh download from the ERP, every company/unit, weekly chunks ----
    frames, failed = [], []
    for company in sr.COMPANIES:
        session = sr.login(company["client_company"])
        if not session:
            failed.append(f"{company['name']} (login failed)")
            continue
        for unit in company["units"]:
            if only_unit and unit["name"].upper() != only_unit:
                continue
            for c_from, c_to in sr.daterange_chunks(args.from_date, args.to_date, chunk_days=7):
                df = parse_all_lines(sr.download_sales_register(session, company["plant"], unit["showroom"], c_from, c_to))
                if df is None:
                    failed.append(f"{unit['name']} {c_from}-{c_to}")
                elif len(df):
                    df.insert(0, "Company", company["name"])
                    frames.append(df)
                time.sleep(0.3)
            print(f"  downloaded {company['name']} / {unit['name']}")
    erp = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(
        columns=["Company", "Customer", "Taxable Value", "FG Code", "Unit", "Invoice Date"])

    # ---- 2. What the database holds for the same customer and range ----
    conn = sqlite3.connect(sr.DB_FILE)
    lo = args.from_date[6:] + args.from_date[3:5] + args.from_date[:2]
    hi = args.to_date[6:] + args.to_date[3:5] + args.to_date[:2]
    db_rows = conn.execute("""
        SELECT unit, invoice_date, taxable_value FROM sales_register
        WHERE UPPER(customer) LIKE ?
          AND substr(invoice_date,7,4) || substr(invoice_date,4,2) || substr(invoice_date,1,2) BETWEEN ? AND ?
          AND (? IS NULL OR UPPER(unit) = ?)
    """, (f"%{args.customer.upper()}%", lo, hi, only_unit, only_unit)).fetchall()
    conn.close()

    r = analyse(erp, db_rows, args.customer)

    # ---- 3. Report ----
    print()
    print(f"Customer containing '{args.customer}', {args.from_date} to {args.to_date}, "
          f"{'unit ' + only_unit if only_unit else 'all units'}")
    print("=" * 70)
    print(f"ERP total (all invoice lines)                 {money(r['erp_total']):>22}")
    print(f"  minus non-FG lines (excluded on purpose)    {money(r['non_fg_total']):>22}")
    print(f"  minus lines dropped by the duplicate key    {money(r['dup_total']):>22}")
    print(f"  = what the database SHOULD hold             {money(r['expected_db_total']):>22}")
    print(f"What the database DOES hold                   {money(r['actual_db_total']):>22}")
    missing = r["expected_db_total"] - r["actual_db_total"]
    print(f"  difference (missing weeks / other)          {money(missing):>22}")
    print("=" * 70)

    if failed:
        print(f"\n⚠️  {len(failed)} download(s) failed during THIS check, so the ERP total above may be low too:")
        for f in failed[:10]:
            print("   ", f)

    if r["non_fg_total"]:
        print("\nNon-FG lines (top 15 by value) - the dashboard deliberately leaves these out:")
        for code, val in r["non_fg_by_code"].head(15).items():
            name = r["non_fg_names"].get(code, "") if hasattr(r["non_fg_names"], "get") else ""
            print(f"   {code:<16} {str(name)[:40]:<40} {money(val):>18}")

    if len(r["dup_lines"]):
        print(f"\nLines dropped by the duplicate key ({len(r['dup_lines'])}) - each shares Invoice No + FG Code + Date + Qty")
        print("with a line that WAS kept, so only the first one is stored:")
        for _, row in r["dup_lines"].head(25).iterrows():
            print(f"   {row['Unit']!s:<4} {row['Invoice Date']!s:<11} {row['Invoice No']!s:<22} {row['FG Code']!s:<10} "
                  f"qty {row['Quantity']!s:<8} {money(row['_value']):>16}")

    if r["week_gaps"]:
        print("\nWeeks where the database doesn't match the ERP (unit, week starting):")
        for unit, week, exp, act, diff in r["week_gaps"][:25]:
            print(f"   {unit:<4} week of {week}   ERP {money(exp):>16}   DB {money(act):>16}   missing {money(diff):>16}")

    print("\nNothing was changed - this only reads.")
