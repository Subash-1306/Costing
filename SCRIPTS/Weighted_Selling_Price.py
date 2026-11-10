"""
Weighted Selling Price Master.

Computes a weighted-average SELLING price for every FG that appears in
Sales Register each month - mirroring exactly how Weighted Price Master
computes a weighted-average PURCHASE price from GRN, just from the sales
side instead.

For a given FG and month:
    weighted_price = SUM(Taxable Value) / SUM(Quantity)
                      for that FG's sales that month
    (only counted if Quantity > 0)

No carry-forward: Part-wise Margin only ever shows FGs actually sold in
the selected month, so a row for a month with no sales would never be
displayed anyway - carrying a price forward into an unsold month adds
complexity with no observable effect.

Customer and Unit are also captured, but only as REFERENCE fields - the
customer and unit from the FG's single MOST RECENT invoice that month (an
FG can legitimately be sold to several customers, from several units,
within one month; this isn't an attempt to represent all of them, just a
quick "who/where was this last invoiced" reference alongside the
aggregate price).

Saves into Analytics.db, table weighted_selling_price_master.
Safe to re-run any time - it recomputes everything from scratch each run.

Run: python Weighted_Selling_Price.py
"""

import sqlite3
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_DIR = os.path.join(SCRIPT_DIR, "..", "DATABASE")
os.makedirs(DATABASE_DIR, exist_ok=True)

SALES_DB = os.path.join(DATABASE_DIR, "Sales_Register.db")
OTHER_SALES_DB = os.path.join(DATABASE_DIR, "Other_Sales.db")   # hand-uploaded sales (Data Sync > Other Sales)
ANALYTICS_DB = os.path.join(DATABASE_DIR, "Analytics.db")


def parse_sale_date_to_month(date_str):
    """Sales dates are stored as dd/mm/yyyy text - return 'YYYY_MM' or None."""
    if not date_str or date_str.count("/") != 2:
        return None
    try:
        day, month, year = date_str.split("/")
        return f"{int(year)}_{int(month):02d}"
    except ValueError:
        return None


def date_sort_key(date_str):
    """dd/mm/yyyy -> yyyymmdd, so the latest date compares correctly as text."""
    try:
        day, month, year = date_str.split("/")
        return f"{year}{month}{day}"
    except (ValueError, AttributeError):
        return ""


def get_monthly_sales(conn_sales, table="sales_register", data=None):
    """
    Returns { item_code: { month: {"qty": x, "value": y, "latest_date": str,
                                     "customer": str, "unit": str} } }
    aggregated from sales_register - customer/unit are updated whenever a
    LATER-dated row is seen, so they end up holding the most recent
    invoice's values for that FG that month.
    """
    cur = conn_sales.cursor()
    cur.execute(f"""
        SELECT item_code, customer, unit, invoice_date, quantity, taxable_value
        FROM {table}
        WHERE item_code IS NOT NULL
    """)

    if data is None:
        data = {}
    for item_code, customer, unit, invoice_date, quantity, taxable_value in cur.fetchall():
        month = parse_sale_date_to_month(invoice_date)
        if month is None:
            continue
        try:
            qty = float(quantity) if quantity else 0.0
            value = float(taxable_value) if taxable_value else 0.0
        except (TypeError, ValueError):
            continue

        if qty <= 0:
            continue

        bucket = data.setdefault(item_code, {}).setdefault(
            month, {"qty": 0.0, "value": 0.0, "latest_date": "", "customer": None, "unit": None}
        )
        bucket["qty"] += qty
        bucket["value"] += value

        this_sort_key = date_sort_key(invoice_date)
        if this_sort_key >= bucket["latest_date"]:
            bucket["latest_date"] = this_sort_key
            bucket["customer"] = customer or "(no customer)"
            bucket["unit"] = unit or "(no unit)"

    return data


if __name__ == "__main__":
    if not os.path.exists(SALES_DB) and not os.path.exists(OTHER_SALES_DB):
        print(f"❌ No sales found - run Sales_Register.py (or upload Other Sales) first.")
        exit()

    # ERP sales and hand-uploaded Other Sales go into the SAME totals, so an
    # FG sold both ways gets one weighted price across all of its sales
    sales = {}
    for db_path, table in ((SALES_DB, "sales_register"), (OTHER_SALES_DB, "other_sales")):
        if not os.path.exists(db_path):
            continue
        conn_sales = sqlite3.connect(db_path)
        try:
            sales = get_monthly_sales(conn_sales, table, sales)
            print(f"Read {table}.")
        except sqlite3.OperationalError:
            pass
        conn_sales.close()

    print(f"Found sales history for {len(sales)} FG codes.")

    conn = sqlite3.connect(ANALYTICS_DB)
    cur = conn.cursor()
    cur.execute("DROP TABLE IF EXISTS weighted_selling_price_master")
    cur.execute("""
        CREATE TABLE weighted_selling_price_master (
            item_code TEXT,
            month TEXT,
            customer TEXT,
            unit TEXT,
            qty_sold REAL,
            value_sold REAL,
            final_price REAL,
            PRIMARY KEY (item_code, month)
        )
    """)
    cur.execute("CREATE INDEX idx_wspm_month ON weighted_selling_price_master(month)")

    rows_written = 0
    for item_code, month_data in sales.items():
        for month, d in month_data.items():
            qty, value = d["qty"], d["value"]
            final_price = value / qty if qty > 0 else 0
            cur.execute("""
                INSERT INTO weighted_selling_price_master
                    (item_code, month, customer, unit, qty_sold, value_sold, final_price)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (item_code, month, d["customer"], d["unit"], qty, value, final_price))
            rows_written += 1

    conn.commit()
    conn.close()

    print(f"✅ Done. {rows_written} (FG x month) rows written to weighted_selling_price_master.")
