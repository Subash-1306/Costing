"""
One-time correction script - NOT part of the regular Data Sync pipeline.
Run this by hand, separately, whenever you need it.

Brings historical Production Register entries for ONE SPECIFIC PROCESS in
line with the CURRENT norm rate for each material, as recorded in a
reference month (August 2026 by default) - since norm rates can drift or
get corrected over time, and older entries may still be sitting on an
outdated rate.

For every (item_code, input_item) combination within the target process:
  1. Find the CURRENT reference norm_qty = the norm_qty from the LATEST
     dated entry for that combination within the reference month (same
     "latest wins" rule already used for Norm Qty in Month Posting and the
     Part Cost Calc material popup - if the rate changed more than once
     within the reference month itself, the most recent one wins).
  2. For EVERY row matching that (item_code, input_item, process) combo,
     across EVERY month (not just the reference month), overwrite
     norm_qty with this reference value, and RECOMPUTE
     input_qty = reference_norm_qty x that row's own prod_qty.
     prod_qty itself is never touched - only norm_qty and input_qty.

A combination with NO entry at all in the reference month is left alone
entirely (nothing to correct it against) and reported separately, since
those rows silently NOT changing is worth knowing about.

SAFETY - read this before running with --apply:
  - Dry run by default. This prints exactly what WOULD change and touches
    nothing. You must pass --apply to actually write anything.
  - Before ANY write, this automatically makes a timestamped backup copy
    of Production_Entry.db right next to the original.
  - You must specify the EXACT process name via --process. This script
    will not guess it for you - matching the wrong string could silently
    change the wrong rows. Use --list-processes first if you're not sure
    of the exact spelling/capitalization used in your data.

Usage:
    # See what process names actually exist, optionally filtered:
    python Fix_Process_Norms.py --list-processes
    python Fix_Process_Norms.py --list-processes --search pad

    # Dry run - shows exactly what would change, changes nothing:
    python Fix_Process_Norms.py --process "PAD PRINTING"

    # Dry run against a different reference month:
    python Fix_Process_Norms.py --process "PAD PRINTING" --reference-month 2026_08

    # Actually apply the fix (backs up the DB first, automatically):
    python Fix_Process_Norms.py --process "PAD PRINTING" --apply
"""

import argparse
import os
import shutil
import sqlite3
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
from costing_config import database_dir   # set in costing\config\settings.json
DATABASE_DIR = database_dir()
DB_FILE = os.path.join(DATABASE_DIR, "Production_Entry.db")

DEFAULT_REFERENCE_MONTH = "2026_08"


def date_sort_key(date_str):
    """dd/mm/yyyy -> yyyymmdd, so the latest date compares correctly as text."""
    try:
        day, month, year = date_str.split("/")
        return f"{year}{month}{day}"
    except (ValueError, AttributeError):
        return ""


def list_processes(conn, search=None):
    cur = conn.cursor()
    if search:
        cur.execute(
            "SELECT DISTINCT process, COUNT(*) FROM production_register "
            "WHERE process LIKE ? GROUP BY process ORDER BY process",
            (f"%{search}%",),
        )
    else:
        cur.execute("SELECT DISTINCT process, COUNT(*) FROM production_register GROUP BY process ORDER BY process")
    rows = cur.fetchall()
    if not rows:
        print("No matching processes found." if search else "No processes found at all - is the database populated?")
        return
    print(f"{'Process':<40} Row count")
    print("-" * 55)
    for process, count in rows:
        print(f"{process or '(blank)':<40} {count}")


def get_reference_norms(conn, process, reference_month):
    """
    {(item_code, input_item): reference_norm_qty} - the norm_qty from the
    LATEST-dated entry for that combination within the reference month, for
    the target process only.
    """
    cur = conn.cursor()
    cur.execute("""
        SELECT item_code, input_item, norm_qty, date
        FROM production_register
        WHERE process = ? AND year_month = ?
          AND item_code IS NOT NULL AND input_item IS NOT NULL
          AND norm_qty IS NOT NULL AND TRIM(norm_qty) != ''
    """, (process, reference_month))

    latest = {}   # (item_code, input_item) -> (sort_key, norm_qty)
    for item_code, input_item, norm_qty, date in cur.fetchall():
        key = (item_code, input_item)
        sk = date_sort_key(date)
        if key not in latest or sk >= latest[key][0]:
            latest[key] = (sk, norm_qty)

    return {k: v[1] for k, v in latest.items()}


def preview_changes(conn, process, reference_norms):
    """
    Returns (changes, unmatched_combos) where changes is a list of dicts
    describing every row that would be touched, and unmatched_combos is
    every (item_code, input_item) combo present for this process that has
    NO reference norm (so nothing to correct it against).
    """
    cur = conn.cursor()
    cur.execute("""
        SELECT id, item_code, input_item, prod_qty, norm_qty, input_qty, date, year_month
        FROM production_register
        WHERE process = ?
          AND item_code IS NOT NULL AND input_item IS NOT NULL
    """, (process,))

    changes = []
    unmatched = set()

    for row_id, item_code, input_item, prod_qty, old_norm, old_input_qty, date, year_month in cur.fetchall():
        key = (item_code, input_item)
        if key not in reference_norms:
            unmatched.add(key)
            continue

        new_norm = reference_norms[key]
        try:
            new_input_qty = round(float(new_norm) * float(prod_qty), 6)
        except (TypeError, ValueError):
            # prod_qty or the reference norm itself isn't a usable number -
            # skip this row rather than guess, and surface it separately.
            unmatched.add(key)
            continue

        if str(old_norm) == str(new_norm):
            continue   # already correct, nothing to do

        changes.append({
            "id": row_id, "item_code": item_code, "input_item": input_item,
            "year_month": year_month, "date": date,
            "prod_qty": prod_qty,
            "old_norm_qty": old_norm, "new_norm_qty": new_norm,
            "old_input_qty": old_input_qty, "new_input_qty": new_input_qty,
        })

    return changes, unmatched


def apply_changes(conn, changes):
    cur = conn.cursor()
    for c in changes:
        cur.execute(
            "UPDATE production_register SET norm_qty = ?, input_qty = ? WHERE id = ?",
            (str(c["new_norm_qty"]), str(c["new_input_qty"]), c["id"]),
        )
    conn.commit()


def backup_database():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = DB_FILE.replace(".db", f"_backup_{timestamp}.db")
    shutil.copy2(DB_FILE, backup_path)
    return backup_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--process", help="Exact process name to correct, e.g. \"PAD PRINTING\" (case-sensitive, must match exactly)")
    parser.add_argument("--reference-month", default=DEFAULT_REFERENCE_MONTH, help=f"Month to treat as the current/correct BOM (default: {DEFAULT_REFERENCE_MONTH})")
    parser.add_argument("--apply", action="store_true", help="Actually write the changes. Without this, it's a dry run only.")
    parser.add_argument("--list-processes", action="store_true", help="List distinct process names present in the data, then exit.")
    parser.add_argument("--search", help="With --list-processes, only show processes containing this text (case-insensitive).")
    args = parser.parse_args()

    if not os.path.exists(DB_FILE):
        print(f"❌ {DB_FILE} not found.")
        exit(1)

    conn = sqlite3.connect(DB_FILE)

    if args.list_processes:
        list_processes(conn, args.search)
        conn.close()
        exit(0)

    if not args.process:
        print("❌ --process is required (or use --list-processes to see what's available).")
        conn.close()
        exit(1)

    print(f"Reference month (the 'current BOM'): {args.reference_month}")
    print(f"Target process: {args.process!r}")
    print()

    reference_norms = get_reference_norms(conn, args.process, args.reference_month)
    if not reference_norms:
        print(f"❌ No entries found for process {args.process!r} in {args.reference_month}.")
        print("   Nothing to use as the 'current' reference - check the process name and reference month.")
        conn.close()
        exit(1)

    print(f"Found {len(reference_norms)} (item_code, input_item) combination(s) with a reference norm rate in {args.reference_month}.")
    print()

    changes, unmatched = preview_changes(conn, args.process, reference_norms)

    if not changes:
        print("✅ Nothing to change - every row already matches its reference month's norm rate.")
    else:
        print(f"{'MODE: DRY RUN (nothing will be written)' if not args.apply else 'MODE: APPLYING CHANGES'}")
        print(f"{len(changes)} row(s) would be updated:")
        print()
        print(f"{'Month':<9} {'Date':<11} {'Item Code':<14} {'Input Item':<14} {'Old Norm':>10} {'New Norm':>10} {'Old Qty':>10} {'New Qty':>10}")
        print("-" * 100)
        for c in changes[:50]:
            print(f"{c['year_month']:<9} {c['date']:<11} {c['item_code']:<14} {c['input_item']:<14} "
                  f"{str(c['old_norm_qty']):>10} {str(c['new_norm_qty']):>10} "
                  f"{str(c['old_input_qty']):>10} {str(c['new_input_qty']):>10}")
        if len(changes) > 50:
            print(f"... and {len(changes) - 50} more row(s)")

    if unmatched:
        print()
        print(f"⚠️  {len(unmatched)} (item_code, input_item) combination(s) for this process have NO entry in "
              f"{args.reference_month}, so they were left untouched (nothing to correct them against):")
        for item_code, input_item in sorted(unmatched)[:20]:
            print(f"   {item_code} / {input_item}")
        if len(unmatched) > 20:
            print(f"   ... and {len(unmatched) - 20} more")

    print()

    if not changes:
        conn.close()
        exit(0)

    if not args.apply:
        print("This was a DRY RUN - nothing was changed. Re-run with --apply to actually apply these changes.")
        conn.close()
        exit(0)

    print("Backing up Production_Entry.db before making any changes...")
    backup_path = backup_database()
    print(f"✅ Backup saved to: {backup_path}")

    apply_changes(conn, changes)
    print(f"✅ Done. {len(changes)} row(s) updated.")

    conn.close()
