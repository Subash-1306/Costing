r"""
One-time repair: adds item codes that entries have in the ERP but the
dashboard never saved.

Why: for a while in September/October a wrong database rule kept only one
row per entry + material. An entry posting 2+ codes from ONE shared
material (e.g. a family mould: LH and RH from the same resin, like
BP2627-46332) lost every code after the first. Those entries were never
downloaded again.

What it does, per stored entry: asks the ERP for the entry, and if the ERP
has codes the dashboard is missing, it ADDS rows for just those codes. Rows
already stored are not touched, so norm corrections made with
Fix_Process_Norms.py stay. Nothing is deleted.

    python Repair_Missing_Outputs.py                     every stored entry (several hours)
    python Repair_Missing_Outputs.py --from 01/09/2026   only entries dated on/after this day
    python Repair_Missing_Outputs.py --report            only list what's missing, change nothing

It can be stopped at any time (Ctrl+C or closing the window) and started
again: entries already checked are skipped. "--restart" checks everything
again from the beginning. Also on the Data Sync page as "Repair missing
output codes".

Afterwards, run BOM Register for the months it reports, so the BOM trees
include the added codes. Part cost calc, margins and Month posting read the
production rows directly and update by themselves.
"""

import sys
import time
from datetime import datetime

import Production_Entry as pe

TABLES = {
    "production_entry": (686, pe.flatten_production_entry),
    "sub_assembly_entry": (922, pe.flatten_subassembly_entry),
}
CLIENT_COMPANY = {"unitech": "9", "veetech": "11", "vtech": "11"}


def arg_value(name):
    if name in sys.argv:
        i = sys.argv.index(name)
        if i + 1 < len(sys.argv):
            return sys.argv[i + 1]
    return None


def to_yyyymmdd(ddmmyyyy):
    try:
        d, m, y = ddmmyyyy.split("/")
        return f"{y}{int(m):02d}{int(d):02d}"
    except (AttributeError, ValueError):
        return None


def main():
    report_only = "--report" in sys.argv
    from_date = arg_value("--from")
    from_key = to_yyyymmdd(from_date) if from_date else None
    if from_date and not from_key:
        print("❌ --from needs a date like 01/09/2026")
        return 1

    conn = pe.get_connection()
    conn.execute("""CREATE TABLE IF NOT EXISTS repair_outputs_done (
                        table_name TEXT, entry_id TEXT, checked_at TEXT, added INTEGER,
                        PRIMARY KEY (table_name, entry_id))""")
    if "--restart" in sys.argv:
        conn.execute("DELETE FROM repair_outputs_done")
    conn.commit()
    if not report_only:
        pe.load_process_master(conn)
        pe.load_mould_master(conn)
        pe.load_item_master(conn)

    # stored entries: entry id -> (entry_no, company, date) from the rows; which ERP list it came from
    entries = conn.execute("""
        SELECT p.table_name, p.entry_id
        FROM processed_entries p
        LEFT JOIN repair_outputs_done r ON r.table_name = p.table_name AND r.entry_id = p.entry_id
        WHERE r.entry_id IS NULL
        ORDER BY p.table_name, CAST(p.entry_id AS INTEGER)
    """).fetchall()
    # entry date per stored entry id (for --from); an id can repeat across companies/entry types
    stored_dates = {}
    for eid, date in conn.execute("SELECT DISTINCT source_entry_id, date FROM production_register"):
        stored_dates.setdefault(eid, set()).add(date)

    todo = []
    for table_name, eid in entries:
        dates = stored_dates.get(eid)
        if not dates:
            continue                      # nothing stored for it (an empty entry) - nothing to compare
        if from_key and max((to_yyyymmdd(x) or "0") for x in dates) < from_key:
            continue
        todo.append((table_name, eid))
    print(f"{len(todo)} entr(ies) to check{' (dated on/after ' + from_date + ')' if from_date else ''}"
          f"{' - report only, nothing will change' if report_only else ''}.")
    if todo:
        print(f"At about {pe.DELAY_BETWEEN_CALLS + 0.35:.1f} s each that is roughly "
              f"{len(todo) * (pe.DELAY_BETWEEN_CALLS + 0.35) / 3600:.1f} hour(s). Safe to stop and start again.\n")

    sessions = {}
    checked = added_entries = added_rows = failed = 0
    months = set()
    started = time.time()
    for table_name, eid in todo:
        entry_type, _, company_name = table_name.rpartition("_")
        schema_id, flatten = TABLES[entry_type]
        cc = CLIENT_COMPANY.get(company_name.lower(), "9")
        if cc not in sessions:
            sessions[cc] = pe.login(cc)
            if not sessions[cc]:
                print(f"❌ Could not log in for {company_name} - stopping (run again later).")
                break
        detail = pe.fetch_detail(sessions[cc], entry_type, schema_id, eid)
        time.sleep(pe.DELAY_BETWEEN_CALLS)
        if not (detail and detail.get("status") == "success"):
            failed += 1                    # not marked done: tried again next run
            continue

        rows = flatten(detail)
        code_col, entry_col, company_col = (pe.COLUMNS.index(c) for c in ("item_code", "entry_no", "company"))
        if not rows:
            checked += 1
            continue
        entry_no, company = rows[0][entry_col], rows[0][company_col]
        have = {c for (c,) in conn.execute(
            "SELECT DISTINCT item_code FROM production_register WHERE source_entry_id = ? AND entry_no IS ? AND company IS ?",
            (str(eid), entry_no, company))}
        if not have:
            checked += 1                   # this id belongs to a different entry here - not ours to touch
            continue
        missing_rows = [r for r in rows if r[code_col] not in have]
        n_added = 0
        if missing_rows:
            codes = sorted({r[code_col] for r in missing_rows})
            months.add(missing_rows[0][pe.COLUMNS.index("year_month")])
            print(f"  {entry_no} ({rows[0][pe.COLUMNS.index('date')]}): missing {', '.join(codes)}"
                  f"{'' if report_only else ' - added'}")
            if not report_only:
                placeholders = ", ".join("?" for _ in pe.COLUMNS)
                conn.executemany(
                    f"INSERT INTO production_register (source_entry_id, {', '.join(pe.COLUMNS)}) VALUES (?, {placeholders})",
                    [[str(eid)] + r for r in missing_rows])
                n_added = len(missing_rows)
                added_rows += n_added
            added_entries += 1
        if not report_only:
            conn.execute("INSERT OR REPLACE INTO repair_outputs_done VALUES (?, ?, ?, ?)",
                         (table_name, eid, datetime.now().strftime("%Y-%m-%d %H:%M"), n_added))
        checked += 1
        if checked % 200 == 0:
            conn.commit()
            rate = checked / (time.time() - started)
            print(f"  ... {checked}/{len(todo)} checked, {added_entries} with missing codes "
                  f"(about {(len(todo) - checked) / rate / 60:.0f} min left)")
    conn.commit()
    conn.close()

    print(f"\n✅ Checked {checked} entr(ies): {added_entries} had codes missing"
          f"{'' if report_only else f', {added_rows} row(s) added'}. {failed} couldn't be read (tried again next run).")
    if months and not report_only:
        print(f"Now run BOM Register for: {', '.join(sorted(m for m in months if m))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
