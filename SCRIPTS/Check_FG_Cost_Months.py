"""
For one month: which FG sales have ALL their production costs from that same
month - the FG itself and every SFG in its BOM tree produced that month, none
costed from an earlier month's production. READ ONLY - changes nothing.

(Purchase prices carried forward from an earlier GRN are NOT counted here -
only production costs.)

Each FG sold in the month falls in one group:
  1. All production costs from this month
  2. Needs an earlier month's production cost (the FG or an SFG wasn't
     produced this month, so its cost comes from the latest month it was)
  3. Cost incomplete - left out by Overview when
     "Include lines with missing cost data" is unticked

Keep this file in SCRIPTS. Double-click it and it asks for the details, or:
    python Check_FG_Cost_Months.py --month 07/2026 --unit U1 --customer "LS AUTO"
(--unit and --customer are optional; leave them out for all.)

Writes an Excel to DOWNLOADS:
  sheet 1 - every sales line in group 1, with cost and margin
  sheet 2 - every FG sold, its group, and which parts came from which month
"""

import argparse
import calendar
import os
import re
import sqlite3
import sys
from collections import defaultdict

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, "..", "DASHBOARD"))
import ProjectDashboard as dash   # the dashboard's real costing code

import openpyxl
import openpyxl.styles
import openpyxl.utils

PART_COST_RE = re.compile(r"Part cost \((\d{4}_\d{2}), (.*), ([^,]*)\)$")
G1 = "1. All production costs from this month"
G2 = "2. Needs an earlier month's production cost"
G3 = "3. Cost incomplete"
GROUPS = [G1, G2, G3]


def cr(x):
    return f"Rs {x / 1e7:,.4f} Cr (Rs {x:,.2f})"


def sales_lines(M, from_date, to_date, units, needle):
    """Every sales line Overview reads for the month (ERP Sales Register +
    uploaded Other Sales, same filters), with invoice details added."""
    lo, hi = dash.ddmmyyyy_to_sortable(from_date), dash.ddmmyyyy_to_sortable(to_date)
    out = []
    for db, table in ((os.path.join(dash.DATABASE_DIR, "Sales_Register.db"), "sales_register"),
                      (dash.OTHER_SALES_DB, "other_sales")):
        if not os.path.exists(db):
            continue
        conn = sqlite3.connect(db)
        if table == "sales_register":
            dash.ensure_sales_register_year_month(conn)
        cols = {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}
        inv = "invoice_no" if "invoice_no" in cols else "NULL"
        q = f"""SELECT unit, customer, invoice_date, {inv}, item_code, quantity, taxable_value FROM {table}
                WHERE year_month = ? AND item_code IS NOT NULL AND TRIM(item_code) != ''
                  AND substr(invoice_date,7,4) || substr(invoice_date,4,2) || substr(invoice_date,1,2) BETWEEN ? AND ?"""
        params = [M, lo, hi]
        if units:
            q += f" AND unit IN ({', '.join('?' for _ in units)})"
            params += units
        try:
            rows = conn.execute(q, params).fetchall()
        except sqlite3.OperationalError:
            rows = []
        conn.close()
        for unit, cust, date, invno, fg, qty, val in rows:
            if needle not in str(cust or "(no customer)").upper():
                continue
            try:   # Overview skips values it can't read as numbers - so do we
                qty = float(qty) if qty else 0.0
                val = float(val) if val else 0.0
            except (TypeError, ValueError):
                continue
            out.append({"unit": unit, "customer": cust or "(no customer)", "date": date, "invoice_no": invno,
                        "fg": fg, "qty": qty, "value": val, "source": "ERP" if table == "sales_register" else "Uploaded"})
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--month", required=True, help="e.g. 07/2026")
    ap.add_argument("--unit", default="", help="e.g. U1 (blank = all units)")
    ap.add_argument("--customer", default="", help="text in the customer name (blank = all customers)")
    interactive = len(sys.argv) == 1
    if interactive:
        print("Press Enter to accept the value in [brackets].\n")
        a = ap.parse_args(["--month", input("Month (mm/yyyy) [07/2026]: ").strip() or "07/2026",
                           "--unit", input("Unit, e.g. U1 [all units]: ").strip(),
                           "--customer", input("Customer name contains [all customers]: ").strip()])
    else:
        a = ap.parse_args()

    mon, year = a.month.strip().split("/")
    M = f"{int(year)}_{int(mon):02d}"
    last = calendar.monthrange(int(year), int(mon))[1]
    from_date, to_date = f"01/{int(mon):02d}/{year}", f"{last:02d}/{int(mon):02d}/{year}"
    units = [a.unit.strip().upper()] if a.unit.strip() else None
    needle = a.customer.strip().upper()
    month_name = calendar.month_name[int(mon)] + " " + year

    lines = sales_lines(M, from_date, to_date, units, needle)

    # ---- Cost every FG sold, noting any part costed from an earlier month ----
    item_types = dash.get_item_types_for_month(M)
    weighted = dash.get_weighted_prices_for_month(M)
    months_desc = sorted([m for m in dash.get_all_relevant_months() if m <= M], reverse=True)
    cost_cache = {}
    conn = sqlite3.connect(dash.ANALYTICS_DB)
    try:
        has_tree = {r[0] for r in conn.execute("SELECT DISTINCT root_fg FROM bom_tree WHERE month = ?", (M,))}
    except sqlite3.OperationalError:
        has_tree = set()
    conn.close()

    fg_info = {}
    for fg in sorted({l["fg"] for l in lines}):
        info = {"group": G3, "unit_cost": None, "earlier": [], "why_incomplete": ""}
        if fg not in has_tree:
            info["why_incomplete"] = "No BOM tree for this month"
        else:
            d = dash.get_single_fg_tree_detail(fg, M, item_types, weighted, months_desc, cost_cache)
            if d["is_missing"]:
                info["why_incomplete"] = ("FG's own cost not found" if d["root_cost_unknown"]
                                          else f"{d['missing_count']} part(s) with no cost")
            else:
                info["unit_cost"] = d["total_cost"]
                for code, source in [(fg, d["root_source"])] + [(n["code"], n["cost_source"]) for n in d["sub_items"]]:
                    m = PART_COST_RE.match(source or "")
                    if m and m.group(1) != M:
                        info["earlier"].append(f"{code} ({m.group(2)}, {m.group(3)}) from {m.group(1)}")
                info["group"] = G2 if info["earlier"] else G1
        fg_info[fg] = info

    # ---- Summary ----
    tot = {g: {"lines": 0, "value": 0.0, "fgs": set()} for g in GROUPS}
    for l in lines:
        t = tot[fg_info[l["fg"]]["group"]]
        t["lines"] += 1; t["value"] += l["value"]; t["fgs"].add(l["fg"])
    all_value = sum(l["value"] for l in lines)

    # Same sales as Overview? (sanity check - should always match)
    ov_all = [r for r in dash.compute_margin_rows(from_date, to_date, units=units, include_missing=True)
              if needle in str(r["customer"]).upper()]
    ov_unticked = sum(r["revenue"] for r in dash.compute_margin_rows(from_date, to_date, units=units, include_missing=False)
                      if needle in str(r["customer"]).upper())

    scope = f"{month_name} | {units[0] if units else 'all units'} | {a.customer.strip() or 'all customers'}"
    print(f"\n{scope}")
    print("=" * 84)
    print(f"All FG sales: {len(lines)} lines, {len({l['fg'] for l in lines})} FGs, {cr(all_value)}")
    if abs(all_value - sum(r["revenue"] for r in ov_all)) >= 1:
        print("  ❌ This doesn't match Overview's total for the same filters - send me this output.")
    print()
    for g in GROUPS:
        t = tot[g]
        share = t["value"] / all_value * 100 if all_value else 0
        print(f"  {g}")
        print(f"      {t['lines']:>6} sales lines | {len(t['fgs']):>5} FGs | {cr(t['value'])} | {share:5.1f}% of value")
    complete = tot[G1]["value"] + tot[G2]["value"]
    print(f"\n  Groups 1 + 2 = {cr(complete)}  vs  Overview with the box unticked = {cr(ov_unticked)}  "
          + ("✅ same" if abs(complete - ov_unticked) < 1 else "❌ different - send me this output"))

    # ---- Excel ----
    from costing_config import downloads_dir
    out_dir = downloads_dir()
    os.makedirs(out_dir, exist_ok=True)
    tag = f"{M}_{units[0] if units else 'ALL'}" + (f"_{re.sub(r'[^A-Z0-9]+', '_', needle)}" if needle else "")
    path = os.path.abspath(os.path.join(out_dir, f"FG_Sales_{tag}.xlsx"))
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Only this month's production"[:31]
    ws.append(["Customer", "Unit", "Invoice Date", "Invoice No", "FG Code", "Qty", "Sales Value (Rs)",
               "Cost per piece (Rs)", "Cost (Rs)", "Margin (Rs)", "Margin % (of cost)", "Source"])
    g1_lines = sorted((l for l in lines if fg_info[l["fg"]]["group"] == G1),
                      key=lambda l: (l["customer"], dash.ddmmyyyy_to_sortable(l["date"]), str(l["invoice_no"])))
    for l in g1_lines:
        uc = fg_info[l["fg"]]["unit_cost"]
        cost = round(uc * l["qty"], 2)
        margin = round(l["value"] - cost, 2)
        ws.append([l["customer"], l["unit"], l["date"], l["invoice_no"], l["fg"], l["qty"], round(l["value"], 2),
                   round(uc, 4), cost, margin, round(margin / cost * 100, 2) if cost else None, l["source"]])
    ws.append([])
    ws.append(["TOTAL", "", "", "", "", sum(l["qty"] for l in g1_lines), round(sum(l["value"] for l in g1_lines), 2), "",
               round(sum(fg_info[l["fg"]]["unit_cost"] * l["qty"] for l in g1_lines), 2),
               round(sum(l["value"] - fg_info[l["fg"]]["unit_cost"] * l["qty"] for l in g1_lines), 2)])
    widths = (40, 7, 12, 22, 12, 10, 16, 16, 14, 14, 16, 10)

    ws2 = wb.create_sheet("All FGs")
    ws2.append(["FG Code", "Group", "Sales lines", "Qty sold", "Sales value (Rs)", "Cost per piece (Rs)",
                "Parts costed from an earlier month", "Why cost is incomplete"])
    per = defaultdict(lambda: [0, 0.0, 0.0])
    for l in lines:
        p = per[l["fg"]]; p[0] += 1; p[1] += l["qty"]; p[2] += l["value"]
    for fg in sorted(fg_info, key=lambda f: (fg_info[f]["group"], -per[f][2])):
        i, p = fg_info[fg], per[fg]
        ws2.append([fg, i["group"], p[0], p[1], round(p[2], 2),
                    round(i["unit_cost"], 4) if i["unit_cost"] is not None else None,
                    "; ".join(i["earlier"]), i["why_incomplete"]])

    for sheet, w in ((ws, widths), (ws2, (12, 44, 12, 12, 16, 16, 70, 28))):
        for cell in sheet[1]:
            cell.font = openpyxl.styles.Font(bold=True)
        for i, width in enumerate(w, start=1):
            sheet.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
        sheet.freeze_panes = "A2"
    for cell in ws[ws.max_row]:
        cell.font = openpyxl.styles.Font(bold=True)
    wb.save(path)

    print(f"\nExcel: {path}")
    print(f"  Sheet 1: the {len(g1_lines)} sales lines in group 1, with cost and margin")
    print(f"  Sheet 2: every FG sold, its group, and which parts came from which month")
    print("\nNothing was changed - this only reads.")
    if interactive:
        input("\nPress Enter to close...")


if __name__ == "__main__":
    main()
