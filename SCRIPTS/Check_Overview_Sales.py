"""
Why does Overview show less than the Sales Register download? READ ONLY.

Runs, on YOUR database:
  - the exact query behind Data Sync > Sales Register > Download (date range), and
  - Overview's own calculation (imported straight from ProjectDashboard.py),
for one customer, unit and period - then lists every line Overview leaves
out and the reason, and checks the reasons add up to the full difference.

Keep this file in the SCRIPTS folder (it uses ..\\DASHBOARD\\ProjectDashboard.py).
Double-click it (or press Run in your editor) and it asks for the details,
or from Command Prompt in the SCRIPTS folder:
    python Check_Overview_Sales.py --customer "LS AUTO" --unit U1 --from 01/07/2026 --to 31/07/2026
"""

import argparse
import os
import sqlite3
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, "..", "DASHBOARD"))
import ProjectDashboard as dash   # Overview's real code, not a copy of it


def cr(x):
    return f"Rs {x / 1e7:,.4f} Cr  (Rs {x:,.2f})"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--customer", required=True, help='Text in the customer name, e.g. "LS AUTO"')
    ap.add_argument("--unit", required=True, help="e.g. U1")
    ap.add_argument("--from", dest="from_date", required=True, help="dd/mm/yyyy")
    ap.add_argument("--to", dest="to_date", required=True, help="dd/mm/yyyy")

    # Started with no details (double-click, or an editor's Run button)?
    # Ask for them instead of stopping with an error.
    interactive = len(sys.argv) == 1
    if interactive:
        print("Press Enter to accept the value in [brackets].\n")
        def ask(prompt, default):
            answer = input(f"{prompt} [{default}]: ").strip()
            return answer or default
        argv = ["--customer", ask("Customer name contains", "LS AUTO"),
                "--unit", ask("Unit", "U1"),
                "--from", ask("From date (dd/mm/yyyy)", "01/07/2026"),
                "--to", ask("To date (dd/mm/yyyy)", "31/07/2026")]
        a = ap.parse_args(argv)
    else:
        a = ap.parse_args()
    needle, unit = a.customer.upper(), a.unit.strip().upper()
    lo, hi = dash.ddmmyyyy_to_sortable(a.from_date), dash.ddmmyyyy_to_sortable(a.to_date)

    # ---- 1. Exactly what the Data Sync download gives you ----
    db = os.path.join(dash.DATABASE_DIR, "Sales_Register.db")
    conn = sqlite3.connect(db)
    dash.ensure_sales_register_year_month(conn)   # Overview does this too before reading
    lines = conn.execute("""
        SELECT invoice_date, customer, item_code, quantity, taxable_value, unit, year_month
        FROM sales_register
        WHERE substr(invoice_date,7,4) || substr(invoice_date,4,2) || substr(invoice_date,1,2) BETWEEN ? AND ?
    """, (lo, hi)).fetchall()
    conn.close()
    lines = [l for l in lines if needle in str(l[1]).upper() and str(l[5]).upper() == unit]

    def num(v):
        """(amount, readable_by_overview). Overview reads a value only if it's a
        plain number; the amount also counts text like "1,23,456.00" (commas
        removed), because Excel may have turned it into a number in your total."""
        try:
            return (float(v) if v not in (None, "") else 0.0), True
        except (TypeError, ValueError):
            try:
                return float(str(v).replace(",", "").strip()), False
            except ValueError:
                return 0.0, False

    download_total = sum(num(l[4])[0] for l in lines)

    # ---- 2. Overview's own calculation, toggle ON and OFF ----
    def overview(include_missing):
        rows = dash.compute_margin_rows(a.from_date, a.to_date, units=[unit], include_missing=include_missing)
        return rows, sum(r["revenue"] for r in rows if needle in str(r["customer"]).upper()), \
               sum(r["revenue"] for r in rows)
    rows_on, ov_on, all_on = overview(True)
    rows_off, ov_off, all_off = overview(False)

    # ---- 3. Each download line Overview would drop, and why (Overview's own rules) ----
    reasons = {"Taxable Value isn't a readable number": [], "Blank FG code": [],
               "Invoice month doesn't match the date (odd date format)": [],
               "FG cost incomplete (dropped ONLY when the 'missing cost' box is unticked)": []}
    cost_cache = {}
    for l in lines:
        inv_date, cust, code, qty, val, u, ym = l
        expected_ym = f"{inv_date[6:10]}_{inv_date[3:5]}" if inv_date else None
        if not num(val)[1] or not num(qty)[1]:
            reasons["Taxable Value isn't a readable number"].append(l)
        elif not code or not str(code).strip():
            reasons["Blank FG code"].append(l)
        elif ym != expected_ym:
            reasons["Invoice month doesn't match the date (odd date format)"].append(l)
        else:
            if ym not in cost_cache:
                cost_cache[ym] = dash.get_all_fg_total_costs(ym)
            info = cost_cache[ym].get(code)
            if info is None or info["is_missing"]:
                reasons["FG cost incomplete (dropped ONLY when the 'missing cost' box is unticked)"].append(l)

    # ---- 4. Report ----
    print(f"\n{a.customer} | {unit} | {a.from_date} to {a.to_date}")
    print("=" * 78)
    print(f"Sales Register download (same dates, this customer, {unit}):  {cr(download_total)}  [{len(lines)} lines]")
    print(f"Overview, 'missing cost' box TICKED   - this customer:       {cr(ov_on)}")
    print(f"Overview, 'missing cost' box UNTICKED - this customer:       {cr(ov_off)}")
    print("-" * 78)
    print(f"Overview's TOTAL REVENUE card ({unit}, ALL customers) ticked: {cr(all_on)}")
    print(f"Overview's TOTAL REVENUE card ({unit}, ALL customers) unticked: {cr(all_off)}")
    print("=" * 78)

    names = sorted({str(l[1]) for l in lines})
    if len(names) > 1:
        print(f"\nNOTE: this customer appears under {len(names)} names - Overview's 'By customer' table shows each separately:")
        for n in names:
            print(f"   {n}: {cr(sum(num(l[4])[0] for l in lines if str(l[1]) == n))}")

    print("\nLines in the download that Overview leaves out:")
    any_reason = False
    for reason, ls in reasons.items():
        if not ls:
            continue
        any_reason = True
        total = sum(num(l[4])[0] for l in ls)
        print(f"\n  {reason}: {len(ls)} lines, {cr(total)}")
        for l in ls[:15]:
            print(f"     {l[0]}  {str(l[2]):<12} qty {str(l[3]):<10} value {str(l[4]):<16} month '{l[6]}'")
        if len(ls) > 15:
            print(f"     ... and {len(ls) - 15} more")
    if not any_reason:
        print("  none")

    # Do the reasons account for the whole gap?
    always_dropped = sum(num(l[4])[0] for r, ls in reasons.items() if not r.startswith("FG cost") for l in ls)
    cost_dropped = sum(num(l[4])[0] for l in reasons["FG cost incomplete (dropped ONLY when the 'missing cost' box is unticked)"])
    gap_on, gap_off = download_total - ov_on, download_total - ov_off
    print("\nCheck that the reasons explain the whole difference:")
    print(f"  box ticked:   gap {cr(gap_on)}   explained {cr(always_dropped)}"
          f"   {'✅ fully explained' if abs(gap_on - always_dropped) < 1 else '❌ NOT fully explained - send me this output'}")
    print(f"  box unticked: gap {cr(gap_off)}   explained {cr(always_dropped + cost_dropped)}"
          f"   {'✅ fully explained' if abs(gap_off - always_dropped - cost_dropped) < 1 else '❌ NOT fully explained - send me this output'}")
    print("\nNothing was changed - this only reads.")
    if interactive:
        input("\nPress Enter to close...")


if __name__ == "__main__":
    main()
