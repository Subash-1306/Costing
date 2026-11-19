r"""
Compares one or more Production / Sub Assembly entries as they are in the
ERP right now with what the dashboard has stored - to find entries where
the ERP shows 2+ item codes but the dashboard kept fewer.

    python Check_Production_Entry.py PE/U1/2627/012345
    python Check_Production_Entry.py PE/U1/2627/012345 PE/U2/2627/000987
    python Check_Production_Entry.py PE/U1/2627/012345 --fix

Use the entry number exactly as the ERP entry screen shows it.

Without --fix it only reads. It prints, per entry, the item codes in the ERP
and in the dashboard, and saves the ERP's full answer to
DOWNLOADS\entry_<number>.json (send that file if something looks wrong).

--fix replaces that entry's stored rows with the ERP's current version
(any norm correction made to that one entry with Fix_Process_Norms.py is
undone - run it again for that process if you had).
"""

import json
import os
import re
import sys

import Production_Entry as pe
from costing_config import downloads_dir

ENTRY_TABLES = [("production_entry", 686, pe.flatten_production_entry),
                ("sub_assembly_entry", 922, pe.flatten_subassembly_entry)]


def client_company_for(company):
    return "11" if (company or "").strip().lower().startswith("v") else "9"


def stored_outputs(conn, eid, entry_no, company):
    return conn.execute("""
        SELECT item_code, MAX(prod_qty), COUNT(*) FROM production_register
        WHERE source_entry_id = ? AND entry_no IS ? AND company IS ?
        GROUP BY item_code ORDER BY MIN(id)
    """, (eid, entry_no, company)).fetchall()


def fetch_erp_version(session, eid, entry_no):
    for table, schema_id, flatten in ENTRY_TABLES:
        detail = pe.fetch_detail(session, table, schema_id, eid)
        if detail and detail.get("status") == "success":
            d = pe._as_dict(pe._as_dict(detail.get("data")).get("details"))
            if str(d.get("entry_no")) == str(entry_no):
                return table, detail, flatten
    return None, None, None


def erp_outputs(detail, table):
    grp = pe._as_dict(pe._as_dict(detail.get("data")).get("grp_details"))
    outs = pe._as_list(grp.get(f"{table}_output_item_group"))
    ins = pe._as_list(grp.get(f"{table}_input_item_group"))
    return outs, ins, sorted(k for k in grp)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    fix = "--fix" in sys.argv
    if not args:
        print(__doc__)
        return 1

    conn = pe.get_connection()
    if fix:
        pe.load_process_master(conn)
        pe.load_mould_master(conn)
        pe.load_item_master(conn)
    sessions = {}
    out_dir = downloads_dir()
    problems = 0

    for entry_no in args:
        print(f"\n===== {entry_no} =====")
        stored = conn.execute("""
            SELECT DISTINCT source_entry_id, company, process FROM production_register WHERE entry_no = ?
        """, (entry_no,)).fetchall()
        if not stored:
            print("  The dashboard has NO rows for this entry number. Check the number is typed exactly as in the ERP;")
            print("  if it is, the entry hasn't been downloaded yet (or failed) - run Production Entry.")
            problems += 1
            continue

        for eid, company, process in stored:
            print(f"  Company {company} · process {process} · ERP internal id {eid}")
            have = stored_outputs(conn, eid, entry_no, company)
            print("  Stored in the dashboard:")
            for code, qty, n in have:
                print(f"    {code:<20} prod qty {qty:<10} ({n} material row(s))")

            cc = client_company_for(company)
            if cc not in sessions:
                sessions[cc] = pe.login(cc)
            session = sessions[cc]
            if not session:
                print("  ❌ Could not log in to the ERP - can't compare.")
                problems += 1
                continue

            table, detail, flatten = fetch_erp_version(session, eid, entry_no)
            if not detail:
                print("  ❌ The ERP didn't return this entry (deleted, or not readable with this login).")
                problems += 1
                continue

            safe = re.sub(r"[^A-Za-z0-9_-]+", "_", entry_no)
            json_path = os.path.join(out_dir, f"entry_{safe}.json")
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(detail, f, indent=2, ensure_ascii=False)

            outs, ins, groups = erp_outputs(detail, table)
            print(f"  In the ERP now ({table}): {len(outs)} output code(s), {len(ins)} material line(s)")
            for o in outs:
                print(f"    {str(o.get('txt_ref_schema_item_code')):<20} accepted {o.get('acc_qty')}  rejected {o.get('rejected_qty', '-')}")
            other_groups = [g for g in groups if g not in (f"{table}_output_item_group", f"{table}_input_item_group")]
            if other_groups:
                print(f"    (other item lists in this entry: {', '.join(other_groups)})")

            erp_codes = [str(o.get("txt_ref_schema_item_code")) for o in outs]
            have_codes = [str(c) for c, _, _ in have]
            missing = [c for c in erp_codes if c not in have_codes]
            extra = [c for c in have_codes if c not in erp_codes]
            if not missing and not extra:
                print("  ✅ Same item codes in the ERP and the dashboard.")
            else:
                problems += 1
                if missing:
                    print(f"  ❌ Missing in the dashboard: {', '.join(missing)}")
                if extra:
                    print(f"  ⚠️  In the dashboard but no longer in the ERP: {', '.join(extra)}")
                if fix:
                    rows = flatten(detail)
                    pe.REFETCH_ALL = True          # insert_rows then replaces this one entry
                    pe.insert_rows(conn, eid, rows)
                    pe.REFETCH_ALL = False
                    conn.commit()
                    print(f"  🔧 Replaced with the ERP's current version ({len(rows)} row(s)).")
                else:
                    print("  Run again with --fix to replace this entry with the ERP's current version.")
            print(f"  ERP's full answer saved to {json_path}")

    conn.close()
    print("\nDone." if not problems else f"\nDone - {problems} entr(y/ies) need attention, see above.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
