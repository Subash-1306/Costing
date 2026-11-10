"""
Weighted Price Master.

Computes a monthly weighted-average purchase price for every item that
appears in GRN Register, then carries the price forward into any month
where nothing was purchased (so every item has a usable price every month,
even months with no fresh GRN).

For a given item and month:
    weighted_price = SUM(Taxable Value) / SUM(Accepted Qty)
                      for that item's GRN receipts that month
    (only counted if Accepted Qty > 0)

If no receipts that month:
    final_price = carried forward from the most recent earlier month that
                   DID have a price
    price_source_month = which month that price actually came from

If an item has NEVER been purchased: final_price = 0, price_source_month = None

Saves into Analytics.db, table weighted_price_master.
Safe to re-run any time - it recomputes everything from scratch each run
(GRN data doesn't change retroactively in practice, so this is cheap and simple).

Run: python Weighted_Price_Master.py
"""

import sqlite3
import os
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
from costing_config import database_dir   # set in costing\config\settings.json
DATABASE_DIR = database_dir()
os.makedirs(DATABASE_DIR, exist_ok=True)

GRN_DB = os.path.join(DATABASE_DIR, "GRN_Register.db")
ANALYTICS_DB = os.path.join(DATABASE_DIR, "Analytics.db")

BASELINE_YEAR_MONTH = (2026, 6)   # matches the overall data baseline (1 Jun 2026)


def months_between(start_year_month, end_year_month):
    y, m = start_year_month
    end_y, end_m = end_year_month
    result = []
    while (y, m) <= (end_y, end_m):
        result.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return result


def month_key(year, month):
    return f"{year}_{month:02d}"


def parse_grn_date_to_month(date_str):
    """GRN dates are stored as dd/mm/yyyy text - return (year, month) or None."""
    if not date_str or date_str.count("/") != 2:
        return None
    try:
        day, month, year = date_str.split("/")
        return (int(year), int(month))
    except ValueError:
        return None


def get_monthly_purchases(conn_grn):
    """
    Returns { item_code: { (year, month): {"qty": x, "value": y} } }
    aggregated from grn_register, using accepted_qty and taxable_value.
    """
    cur = conn_grn.cursor()
    cur.execute("""
        SELECT item_code, grn_date, accepted_qty, taxable_value
        FROM grn_register
        WHERE item_code IS NOT NULL
          AND LOWER(TRIM(COALESCE(grn_type, ''))) <> 'sample'   -- Sample GRNs aren't real purchases
    """)

    data = {}
    for item_code, grn_date, accepted_qty, taxable_value in cur.fetchall():
        ym = parse_grn_date_to_month(grn_date)
        if ym is None:
            continue
        try:
            qty = float(accepted_qty) if accepted_qty else 0.0
            value = float(taxable_value) if taxable_value else 0.0
        except (TypeError, ValueError):
            continue

        if qty <= 0:
            continue

        data.setdefault(item_code, {}).setdefault(ym, {"qty": 0.0, "value": 0.0})
        data[item_code][ym]["qty"] += qty
        data[item_code][ym]["value"] += value

    return data


def get_connection(db_path, create_sql):
    conn = sqlite3.connect(db_path)
    conn.execute(create_sql)
    conn.commit()
    return conn


if __name__ == "__main__":
    if not os.path.exists(GRN_DB):
        print(f"❌ {GRN_DB} not found - run GRN_Register.py first.")
        exit()

    conn_grn = sqlite3.connect(GRN_DB)
    purchases = get_monthly_purchases(conn_grn)
    conn_grn.close()

    print(f"Found purchase history for {len(purchases)} items.")

    now = datetime.now()
    current_year_month = (now.year, now.month)
    all_months = months_between(BASELINE_YEAR_MONTH, current_year_month)

    conn = get_connection(ANALYTICS_DB, """
        CREATE TABLE IF NOT EXISTS weighted_price_master (
            item_code TEXT,
            month TEXT,
            qty_purchased REAL,
            value_purchased REAL,
            weighted_price REAL,
            final_price REAL,
            price_source_month TEXT,
            PRIMARY KEY (item_code, month)
        )
    """)

    cur = conn.cursor()
    rows_written = 0

    # Rebuilt from scratch every run (in one transaction, so the dashboard
    # keeps seeing the old prices until the new ones are ready). This also
    # clears out old rows that no longer apply - e.g. "no price yet" rows
    # written by earlier versions, or items whose GRN has since been removed.
    cur.execute("DELETE FROM weighted_price_master")

    for item_code, month_data in purchases.items():
        last_known_price = None
        last_known_month = None

        for (year, month) in all_months:
            mkey = month_key(year, month)

            if (year, month) in month_data:
                qty = month_data[(year, month)]["qty"]
                value = month_data[(year, month)]["value"]
                weighted_price = value / qty if qty > 0 else None

                if weighted_price is not None:
                    final_price = weighted_price
                    price_source_month = mkey
                    last_known_price = weighted_price
                    last_known_month = mkey
                else:
                    final_price = last_known_price if last_known_price is not None else 0
                    price_source_month = last_known_month
            else:
                qty = None
                value = None
                weighted_price = None
                final_price = last_known_price if last_known_price is not None else 0
                price_source_month = last_known_month

            # No row for months BEFORE the item was ever purchased: there's no
            # GRN this month and nothing earlier to carry forward. A row is
            # written only if it was bought this month, or carried forward
            # from an earlier purchase.
            if price_source_month is None:
                continue

            cur.execute("""
                INSERT INTO weighted_price_master
                    (item_code, month, qty_purchased, value_purchased, weighted_price, final_price, price_source_month)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(item_code, month) DO UPDATE SET
                    qty_purchased = excluded.qty_purchased,
                    value_purchased = excluded.value_purchased,
                    weighted_price = excluded.weighted_price,
                    final_price = excluded.final_price,
                    price_source_month = excluded.price_source_month
            """, (item_code, mkey, qty, value, weighted_price, final_price, price_source_month))
            rows_written += 1

    conn.commit()
    conn.close()

    print(f"✅ Done. {rows_written} (item x month) rows written to weighted_price_master.")
