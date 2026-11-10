"""
Combined extractor: Production Entry + Sub Assembly Entry, for both companies
(Unitech, Vtech), all saved into the same production_register table in
Production_Entry.db (distinguished by the 'process' column).

- First run: fetches ALL entries of both types, for both companies
- Later runs: only fetches new entries (tracked in processed_entries)
- Safe to stop (Ctrl+C) and re-run later - it picks up where it left off.

Run locally: python extract_all.py
"""

import requests
import sqlite3
import time

BASE_URL = "https://unitech.tarka.in"
LOGIN_URL = f"{BASE_URL}/app.php"
DATA_URL = f"{BASE_URL}/app.php"

# ERP login comes from the dashboard (or is asked when run by hand) - nothing is saved in this file
from erp_credentials import get_credentials

COMPANIES = [
    {"name": "Unitech", "client_company": "9"},
    {"name": "Veetech", "client_company": "11"},
]

HEADERS = {
    "X-Requested-With": "XMLHttpRequest",
    "Referer": f"{BASE_URL}/MainPage.html",
}

import os
import sys
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_DIR = os.path.join(SCRIPT_DIR, "..", "DATABASE")
os.makedirs(DATABASE_DIR, exist_ok=True)
DB_FILE = os.path.join(DATABASE_DIR, "Production_Entry.db")
PAGE_SIZE = 100
DELAY_BETWEEN_CALLS = 0.15
PRINT_EVERY_N = 10

# Command-line options (the dashboard's Run button uses neither):
#   --full-scan    read every ERP page to find anything missed; never changes stored rows
#   --refetch-all  re-download EVERYTHING and replace stored rows (undoes norm corrections)
FULL_SCAN = "--full-scan" in sys.argv or "--refetch-all" in sys.argv
REFETCH_ALL = "--refetch-all" in sys.argv

# Loaded once at startup - maps a WIP store code to its real process name.
PROCESS_MASTER = {}


def load_process_master(conn):
    """
    Loads the process_master table into a dict keyed by (unit, store) for
    fast lookup during flatten. This table is populated via the dashboard's
    Process Master download-missing/upload-filled workflow (not from the
    CSV anymore, since the mapping now needs Unit too, which the old CSV
    format didn't have).

    Self-heals older (Store-only) databases by expanding each old row across
    every (Unit, Store) combination already seen in production_register.
    """
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='process_master'")
    if not cur.fetchone():
        conn.execute("""
            CREATE TABLE process_master (
                unit TEXT,
                wip_store TEXT,
                process TEXT,
                PRIMARY KEY (unit, wip_store)
            )
        """)
        conn.commit()
    else:
        cur.execute("PRAGMA table_info(process_master)")
        columns = {row[1] for row in cur.fetchall()}
        if "unit" not in columns:
            print("  Migrating process_master to be Unit-aware (one-time)...")
            cur.execute("SELECT wip_store, process FROM process_master")
            old_rows = cur.fetchall()
            cur.execute("SELECT DISTINCT unit, store FROM production_register WHERE unit IS NOT NULL AND store IS NOT NULL")
            known_combos = cur.fetchall()

            cur.execute("ALTER TABLE process_master RENAME TO process_master_old")
            conn.execute("""
                CREATE TABLE process_master (
                    unit TEXT, wip_store TEXT, process TEXT,
                    PRIMARY KEY (unit, wip_store)
                )
            """)
            old_lookup = dict(old_rows)
            for unit, store in known_combos:
                if store in old_lookup:
                    cur.execute(
                        "INSERT OR IGNORE INTO process_master (unit, wip_store, process) VALUES (?, ?, ?)",
                        (unit, store, old_lookup[store]),
                    )
            cur.execute("DROP TABLE process_master_old")
            conn.commit()

    cur.execute("SELECT unit, wip_store, process FROM process_master")
    global PROCESS_MASTER
    PROCESS_MASTER = {(unit, store): process for unit, store, process in cur.fetchall()}
    print(f"Loaded {len(PROCESS_MASTER)} entries from process_master.")


def get_process(unit, store, default_label):
    """Look up the real process from the master by (unit, store); fall back
    to the old default label if this combo isn't in the master yet."""
    return PROCESS_MASTER.get((unit, store), default_label)


# Loaded once at startup - maps a Mould Code to its Runner Wt (per-part).
MOULD_MASTER = {}


def load_mould_master(conn):
    cur = conn.cursor()
    cur.execute("""
        SELECT name FROM sqlite_master WHERE type='table' AND name='mould_master'
    """)
    if not cur.fetchone():
        print("mould_master table not found yet - run download_mould_master.py first. "
              "Runner/Part will be blank until then.")
        return

    cur.execute("SELECT code, runner_wt FROM mould_master")
    global MOULD_MASTER
    MOULD_MASTER = {code: runner_wt for code, runner_wt in cur.fetchall()}
    print(f"Loaded {len(MOULD_MASTER)} entries from mould_master.")


def get_runner_per_part(mould_field):
    """
    mould_field looks like 'U1MD1001-BI3 & BC3 TUMBLERPIECE 3R5R1A-1420'.
    Split at the FIRST hyphen to get the Mould Code, then look up its Runner Wt.
    """
    if not mould_field or "-" not in mould_field:
        return None
    code = mould_field.split("-", 1)[0].strip()
    return MOULD_MASTER.get(code)


# Loaded once at startup - maps an ERP Code to {month_key: item_type}.
ITEM_MASTER = {}


def load_item_master(conn):
    cur = conn.cursor()
    cur.execute("""
        SELECT name FROM sqlite_master WHERE type='table' AND name='item_master'
    """)
    if not cur.fetchone():
        print("item_master table not found yet - run import_item_master.py first. "
              "Item Type will be blank until then.")
        return

    cur.execute("PRAGMA table_info(item_master)")
    month_columns = [row[1] for row in cur.fetchall() if row[1].startswith("item_type_")]

    if not month_columns:
        print("item_master has no month columns yet - run import_item_master.py first.")
        return

    cur.execute(f"SELECT erp_code, {', '.join(month_columns)} FROM item_master")
    global ITEM_MASTER
    ITEM_MASTER = {}
    for row in cur.fetchall():
        erp_code = row[0]
        month_values = {}
        for col_name, value in zip(month_columns, row[1:]):
            month_key = col_name.replace("item_type_", "")  # e.g. "2026_09"
            month_values[month_key] = value
        ITEM_MASTER[erp_code] = month_values

    print(f"Loaded {len(ITEM_MASTER)} entries from item_master ({len(month_columns)} month columns).")


def get_item_type(item_code, date_str):
    """
    date_str looks like '11/09/2026' (dd/mm/yyyy). Look up this item's type
    as it was classified during THAT month's snapshot.
    """
    if not item_code or not date_str or date_str.count("/") != 2:
        return None
    day, month, year = date_str.split("/")
    month_key = f"{year}_{month}"
    return ITEM_MASTER.get(item_code, {}).get(month_key)

COLUMNS = [
    "company", "unit", "date", "year_month", "supervisor", "shift", "machine", "entry_no",
    "store", "to_store", "remarks", "plan_no", "mould",
    "item_code", "item_name", "process", "uom",
    "acc_qty", "rejected_qty", "prod_qty",
    "input_item", "item_description", "item_type", "norm_qty", "input_qty",
    "runner_per_part", "total_runner",
    "notes", "consumption_required",
]


# ---------- Database setup ----------

def get_connection():
    # Wait up to 60 s if another program (the dashboard, DB Browser) briefly
    # has the database busy, instead of crashing after SQLite's default 5 s.
    # It can't wait out DB Browser holding UNSAVED changes - click "Write
    # Changes" or "Revert Changes" there, or close it, before syncing.
    conn = sqlite3.connect(DB_FILE, timeout=60)
    cur = conn.cursor()

    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS production_register (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_entry_id TEXT,
            {", ".join(f"{col} TEXT" for col in COLUMNS)}
        )
    """)

    cur.execute("PRAGMA table_info(production_register)")
    existing_cols = {row[1] for row in cur.fetchall()}
    for col in COLUMNS:
        if col not in existing_cols:
            print(f"  Adding missing column to production_register: {col}")
            cur.execute(f"ALTER TABLE production_register ADD COLUMN {col} TEXT")

    # Speeds up every month-filtered query dramatically - without this, queries
    # that filter by month have to fall back to substr(date,...) which can't
    # use an index and forces a full table scan every time.
    cur.execute("CREATE INDEX IF NOT EXISTS idx_year_month ON production_register(year_month)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_year_month_group ON production_register(year_month, unit, store, process, item_code)")

    # Finds all stored rows of one entry quickly, so insert_rows can replace
    # an entry instead of adding a second copy of it.
    cur.execute("CREATE INDEX IF NOT EXISTS idx_entry_lookup ON production_register(source_entry_id, entry_no)")

    # An earlier version used a UNIQUE (entry, material) index here. That was
    # wrong: an entry with 2 output items legitimately has the same material
    # once per output, so it silently dropped the second output's rows.
    cur.execute("SELECT 1 FROM sqlite_master WHERE type='index' AND name='idx_unique_entry_material'")
    if cur.fetchone():
        cur.execute("DROP INDEX idx_unique_entry_material")
        print("  ⚠️  Removed the faulty (entry, material) duplicate guard from an earlier version.")
        print("     If that version ran, entries with 2+ output items may be missing rows -")
        print("     run once with --refetch-all to re-download everything cleanly.")

    # Backfill year_month for any existing rows that predate this column
    # (harmless/instant no-op once everything's been backfilled once).
    cur.execute("SELECT COUNT(*) FROM production_register WHERE year_month IS NULL AND date IS NOT NULL")
    if cur.fetchone()[0] > 0:
        print("  Backfilling year_month for existing rows (one-time)...")
        cur.execute("""
            UPDATE production_register
            SET year_month = substr(date,7,4) || '_' || substr(date,4,2)
            WHERE year_month IS NULL AND date IS NOT NULL
        """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS processed_entries (
            table_name TEXT,
            entry_id TEXT,
            PRIMARY KEY (table_name, entry_id)
        )
    """)
    cur.execute("CREATE TABLE IF NOT EXISTS migrations (name TEXT PRIMARY KEY)")
    conn.commit()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS failed_entries (
            table_name TEXT,
            entry_id TEXT,
            PRIMARY KEY (table_name, entry_id)
        )
    """)

    # ONE-TIME: entries stored twice before this fix. Each download of an
    # entry is saved in one go, so its rows get consecutive ids; a second
    # copy saved by a later download sits further on. Keep the FIRST copy
    # (with any corrections already made to it) and delete the later copies.
    # Nothing is re-downloaded, so no stored value changes. If an entry lists
    # the same material twice within ONE save, that may be genuine - it's
    # left alone and listed for you to check.
    if not cur.execute("SELECT 1 FROM migrations WHERE name='dedupe_keep_first_copy'").fetchone():
        candidates = cur.execute("""
            SELECT DISTINCT source_entry_id, entry_no, company FROM production_register
            GROUP BY source_entry_id, entry_no, company, item_code, input_item
            HAVING COUNT(*) > 1
        """).fetchall()
        removed_rows, removed_entries, check_these = 0, 0, []
        for eid, entry_no, company in candidates:
            ids = [r[0] for r in cur.execute(
                "SELECT id FROM production_register WHERE source_entry_id = ? AND entry_no IS ? AND company IS ? ORDER BY id",
                (eid, entry_no, company))]
            batches = [[ids[0]]]
            for prev, cur_id in zip(ids, ids[1:]):
                if cur_id == prev + 1:
                    batches[-1].append(cur_id)
                else:
                    batches.append([cur_id])
            if len(batches) > 1:
                extra = [i for b in batches[1:] for i in b]
                cur.executemany("DELETE FROM production_register WHERE id = ?", [(i,) for i in extra])
                removed_rows += len(extra)
                removed_entries += 1
            else:
                check_these.append(entry_no)
        if removed_entries:
            print(f"  Removed {removed_rows} duplicate rows ({removed_entries} entries had been stored twice); kept the first copy of each.")
        if check_these:
            print(f"  {len(check_these)} entries list the same output + material more than once in a single save -")
            print(f"  left as they are, please spot-check: {', '.join(str(x) for x in check_these[:10])}{' ...' if len(check_these) > 10 else ''}")
        cur.execute("INSERT INTO migrations (name) VALUES ('dedupe_keep_first_copy')")
        conn.commit()

    # Optional: python Production_Entry.py --refetch-all
    # Re-downloads every entry from the ERP (slow - hours). Safe: each entry
    # replaces its own stored rows, so nothing is ever doubled.
    if REFETCH_ALL:
        print("  --refetch-all: every entry will be downloaded again and REPLACE what's stored")
        print("  (undoes any norm corrections - re-run Fix_Process_Norms.py afterwards). Takes hours.")
        cur.execute("DELETE FROM processed_entries")
        conn.commit()

    conn.commit()
    return conn


def load_processed_ids(conn, table_name_key):
    cur = conn.cursor()
    cur.execute(
        "SELECT entry_id FROM processed_entries WHERE table_name = ?",
        (table_name_key,),
    )
    return set(row[0] for row in cur.fetchall())


def mark_processed(conn, table_name_key, entry_id):
    conn.execute(
        "INSERT OR IGNORE INTO processed_entries (table_name, entry_id) VALUES (?, ?)",
        (table_name_key, str(entry_id)),
    )


def insert_rows(conn, source_entry_id, rows):
    """Stores one ERP entry. If this entry (same entry id, entry no and
    company) is ALREADY stored, it's left exactly as it is - so corrections
    made afterwards (e.g. norm qty fixed with Fix_Process_Norms.py) are never
    overwritten, and an entry can never be stored twice. Only with
    --refetch-all is a stored entry replaced by the ERP's current version.
    Every row is kept, including one per output item when an entry has
    several outputs."""
    if not rows:
        return
    company, entry_no = rows[0][COLUMNS.index("company")], rows[0][COLUMNS.index("entry_no")]
    key = (str(source_entry_id), entry_no, company)
    already = conn.execute(
        "SELECT 1 FROM production_register WHERE source_entry_id = ? AND entry_no IS ? AND company IS ? LIMIT 1", key
    ).fetchone()
    if already:
        if not REFETCH_ALL:
            return
        conn.execute("DELETE FROM production_register WHERE source_entry_id = ? AND entry_no IS ? AND company IS ?", key)
    placeholders = ", ".join("?" for _ in COLUMNS)
    col_names = ", ".join(COLUMNS)
    conn.executemany(
        f"INSERT INTO production_register (source_entry_id, {col_names}) VALUES (?, {placeholders})",
        [[str(source_entry_id)] + row for row in rows],
    )


# ---------- ERP calls (generic - works for any table_name/schema_id) ----------

def login(client_company):
    username, password = get_credentials()
    session = requests.Session()
    session.get(BASE_URL)

    resp = session.post(
        LOGIN_URL,
        params={"a": "validateUsernameAndPassword.demo"},
        data={"usr": username, "password": password, "client_company": client_company},
        headers=HEADERS,
    )

    data = resp.json()
    if data.get("status") == "success":
        print("✅ Login successful")
        return session
    else:
        print("❌ Login failed - check your ERP username and password:", data)
        return None


def fetch_list(session, source_table_name, schema_id, page, rows):
    resp = session.get(
        DATA_URL,
        params={
            "a": "viewUserSchemaModuleDetails.demo",
            "table_name": source_table_name,
            "field_name": "",
            "schema_id": schema_id,
            "_search": "false",
            "rows": rows,
            "page": page,
            "sidx": "id",
            "sord": "desc",
        },
        headers=HEADERS,
    )
    return resp.json()


def get_new_entry_ids(session, source_table_name, schema_id, processed_ids, full_scan=False):
    """
    The ERP lists entries NEWEST FIRST. Read page 1, 2, ... and stop at the
    first page that has no new (not yet processed) entries - normally just
    1-2 pages instead of every page. full_scan=True reads every page (used
    once on the first run of this version, or with --full-scan).

    Returns the new ids OLDEST FIRST, each only once (if entries are added in
    the ERP while we page through, an id can show up on two pages).
    Processing oldest first means a crash only ever leaves the newest
    entries undone, which the next run finds on page 1.
    """
    new_ids, seen = [], set()
    page = 1
    while True:
        data = fetch_list(session, source_table_name, schema_id, page=page, rows=PAGE_SIZE)
        page_rows = data.get("rows") or []
        if not page_rows:
            break
        if page == 1:
            print(f"ERP has {data.get('records')} entries in total.")
        page_new = 0
        for row in page_rows:
            eid = str(row["id"])
            if eid in seen:
                continue
            seen.add(eid)
            if eid not in processed_ids:
                new_ids.append(eid)
                page_new += 1
        total_pages = int(data.get("total") or page)
        if page >= total_pages or (page_new == 0 and not full_scan):
            break
        page += 1
        if page % 20 == 0:
            print(f"  Scanned page {page}/{total_pages}...")
        time.sleep(DELAY_BETWEEN_CALLS)
    print(f"Read {page} page(s){' (full scan)' if full_scan else ''}.")
    return list(reversed(new_ids))


def fetch_detail(session, source_table_name, schema_id, internal_code):
    resp = session.post(
        DATA_URL,
        params={"a": "GetSchemaModuleDetailsForUpdateNew"},
        data={
            "internal_code": internal_code,
            "table_name": source_table_name,
            "schema_id": schema_id,
            "search_fields": "undefined",
            "full_arr": "",
            "ur_edit_days": 0,
        },
        headers=HEADERS,
    )
    try:
        return resp.json()
    except ValueError:
        return None


# ---------- Flatten functions (one per entry type, since field names differ) ----------

def date_to_year_month(date_str):
    """'11/09/2026' -> '2026_09' - lets month-filtered queries use a plain
    equality index instead of substr() (which can't use an index and forces
    a full table scan on every query)."""
    if not date_str or date_str.count("/") != 2:
        return None
    day, month, year = date_str.split("/")
    return f"{year}_{month}"


def flatten_production_entry(entry_json):
    rows = []
    d = entry_json["data"]["details"]
    grp = entry_json["data"]["grp_details"]

    header = [
        d.get("txt_ref_schema_plant"),
        d.get("txt_ref_schema_showroom"),
        d.get("date"),
        date_to_year_month(d.get("date")),
        d.get("txt_ref_schema_supervisor"),
        d.get("txt_ref_schema_shift"),
        d.get("txt_ref_schema_machine"),
        d.get("entry_no"),
        d.get("txt_ref_schema_stores"),
        d.get("txt_ref_schema_output_store"),
        d.get("remarks"),
        d.get("txt_ref_schema_work_order"),
        d.get("txt_ref_schema_mould"),
    ]

    mould_field = d.get("txt_ref_schema_mould")
    runner_per_part = get_runner_per_part(mould_field)

    output_items = grp.get("production_entry_output_item_group", []) or [{}]
    input_items = grp.get("production_entry_input_item_group", []) or [{}]

    for out in output_items:
        acc_qty = out.get("acc_qty") or 0
        rejected_qty = 0  # rejection entered in a different form for production entry
        try:
            prod_qty = float(acc_qty) + float(rejected_qty)
        except (TypeError, ValueError):
            prod_qty = acc_qty

        try:
            total_runner = float(runner_per_part) * float(prod_qty) if runner_per_part is not None else None
        except (TypeError, ValueError):
            total_runner = None

        out_row = [
            out.get("txt_ref_schema_item_code"),
            out.get("item_name"),
            get_process(d.get("txt_ref_schema_showroom"), d.get("txt_ref_schema_stores"), "Injection Moulding"),
            out.get("txt_ref_schema_uom"),
            acc_qty,
            rejected_qty,
            prod_qty,
        ]

        for inp in input_items:
            norm_qty = inp.get("norm_qty")
            try:
                allocated_input_qty = float(norm_qty) * float(prod_qty)
            except (TypeError, ValueError):
                allocated_input_qty = inp.get("input_qty")

            input_item_code = inp.get("txt_ref_schema_input_item")
            item_type = get_item_type(input_item_code, d.get("date"))

            in_row = [
                input_item_code,
                inp.get("item_descriprion"),
                item_type,
                norm_qty,
                allocated_input_qty,
                runner_per_part,
                total_runner,
                inp.get("notes"),
                inp.get("ip_consumption_required"),
            ]

            full_row = header + out_row + in_row
            full_row = [None if v is None else str(v) for v in full_row]
            rows.append(full_row)

    if len(output_items) > 1:
        print(f"  ⚠️ Entry {d.get('entry_no')} has {len(output_items)} output items — input qty was split. Please spot-check.")

    return rows


def flatten_subassembly_entry(entry_json):
    rows = []
    d = entry_json["data"]["details"]
    grp = entry_json["data"]["grp_details"]

    header = [
        d.get("txt_ref_schema_plant"),
        d.get("txt_ref_schema_showroom"),
        d.get("date"),
        date_to_year_month(d.get("date")),
        d.get("txt_ref_schema_supervisor"),
        d.get("txt_ref_schema_shift"),
        None,  # machine - not applicable
        d.get("entry_no"),
        d.get("txt_ref_schema_store"),
        d.get("txt_ref_schema_to_store"),
        d.get("remarks"),
        d.get("txt_ref_schema_plan_no"),
        None,  # mould - not applicable for Sub Assembly
    ]

    output_items = grp.get("sub_assembly_entry_output_item_group", []) or [{}]
    input_items = grp.get("sub_assembly_entry_input_item_group", []) or [{}]

    for out in output_items:
        acc_qty = out.get("acc_qty") or 0
        rejected_qty = out.get("rejected_qty") or 0
        try:
            prod_qty = float(acc_qty) + float(rejected_qty)
        except (TypeError, ValueError):
            prod_qty = acc_qty

        out_row = [
            out.get("txt_ref_schema_item_code"),
            out.get("item_name"),
            get_process(d.get("txt_ref_schema_showroom"), d.get("txt_ref_schema_store"), "Sub Assembly"),
            out.get("txt_ref_schema_uom"),
            acc_qty,
            rejected_qty,
            prod_qty,
        ]

        for inp in input_items:
            norm_qty = inp.get("norm_qty")
            try:
                allocated_input_qty = float(norm_qty) * float(prod_qty)
            except (TypeError, ValueError):
                allocated_input_qty = inp.get("input_qty")

            input_item_code = inp.get("txt_ref_schema_input_item")
            item_type = get_item_type(input_item_code, d.get("date"))

            in_row = [
                input_item_code,
                inp.get("item_descriprion"),
                item_type,
                norm_qty,
                allocated_input_qty,
                None,  # runner_per_part - not applicable for Sub Assembly
                None,  # total_runner - not applicable for Sub Assembly
                inp.get("notes"),
                inp.get("ip_consumption_required"),
            ]

            full_row = header + out_row + in_row
            full_row = [None if v is None else str(v) for v in full_row]
            rows.append(full_row)

    if len(output_items) > 1:
        print(f"  ⚠️ Entry {d.get('entry_no')} has {len(output_items)} output items — input qty was split. Please spot-check.")

    return rows


# ---------- Entry type registry ----------

ENTRY_TYPES = [
    {
        "label": "Production Entry",
        "tracking_key": "production_entry",
        "source_table_name": "production_entry",
        "schema_id": 686,
        "flatten_func": flatten_production_entry,
    },
    {
        "label": "Sub Assembly Entry",
        "tracking_key": "sub_assembly_entry",
        "source_table_name": "sub_assembly_entry",
        "schema_id": 922,
        "flatten_func": flatten_subassembly_entry,
    },
]


# ---------- Main ----------

def process_entry_type(conn, session, company_name, entry_type):
    label = entry_type["label"]
    source_table_name = entry_type["source_table_name"]
    schema_id = entry_type["schema_id"]
    flatten_func = entry_type["flatten_func"]
    table_name_key = f"{entry_type['tracking_key']}_{company_name}"

    print(f"\n--- {label} ({company_name}) ---")
    processed_ids = load_processed_ids(conn, table_name_key)

    # The first run of this version reads every page once, to catch anything
    # an earlier interrupted run left behind. After that, newest pages only.
    marker = f"full_scan_done_{table_name_key}"
    first_time = conn.execute("SELECT 1 FROM migrations WHERE name = ?", (marker,)).fetchone() is None
    full_scan = first_time or FULL_SCAN
    new_ids = get_new_entry_ids(session, source_table_name, schema_id, processed_ids, full_scan)

    # Entries that failed to download on an earlier run are retried every
    # run until they succeed (they may be older than where the scan stops).
    retry_ids = [r[0] for r in conn.execute(
        "SELECT entry_id FROM failed_entries WHERE table_name = ?", (table_name_key,))
        if r[0] not in processed_ids and r[0] not in set(new_ids)]
    to_fetch = retry_ids + new_ids
    print(f"Already processed: {len(processed_ids)} | New: {len(new_ids)} | Retrying earlier failures: {len(retry_ids)}")

    count = 0
    for eid in to_fetch:
        detail = fetch_detail(session, source_table_name, schema_id, eid)
        if detail and detail.get("status") == "success":
            rows = flatten_func(detail)
            insert_rows(conn, eid, rows)
            mark_processed(conn, table_name_key, eid)
            conn.execute("DELETE FROM failed_entries WHERE table_name = ? AND entry_id = ?", (table_name_key, eid))
            count += 1
            if count % PRINT_EVERY_N == 0:
                conn.commit()
                print(f"  Progress: {count}/{len(to_fetch)} saved...")
        else:
            print(f"  ❌ Failed entry {eid} - will retry next run")
            conn.execute("INSERT OR IGNORE INTO failed_entries (table_name, entry_id) VALUES (?, ?)", (table_name_key, eid))

        time.sleep(DELAY_BETWEEN_CALLS)

    if first_time:
        conn.execute("INSERT OR IGNORE INTO migrations (name) VALUES (?)", (marker,))
    conn.commit()
    print(f"✅ Done. Saved {count} entries." if to_fetch else "✅ Up to date.")


if __name__ == "__main__":
    conn = get_connection()
    load_process_master(conn)
    load_mould_master(conn)
    load_item_master(conn)

    for company in COMPANIES:
        name = company["name"]
        client_company = company["client_company"]

        print(f"\n===== {name} =====")
        session = login(client_company)
        if not session:
            print(f"❌ Could not log in for {name}, skipping.")
            continue

        for entry_type in ENTRY_TYPES:
            process_entry_type(conn, session, name, entry_type)

    conn.close()
    print(f"\n✅ All companies and entry types processed. Saved to {DB_FILE}.")
