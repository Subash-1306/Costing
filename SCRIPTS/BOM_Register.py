"""
BOM Register -> BOM Tree.

Downloads the BOM Register (as CSV - the Excel export blows the ERP's PHP
memory limit on large units), then stores ONLY the parent->child links that
are reachable from the FG codes actually SOLD that month, rather than
dumping the whole ~53k-row register.

Walk rules:
  - A node expands only if it's Production-type (per Item Master) AND has
    its own BOM rows.
  - Anything Purchase-type is a leaf, no matter whether it's FG or SFG -
    Item Master's classification overrides the BOM.
  - Duplicate output+input pairs have their Norm Qty summed (these are
    genuine separate consumption lines, mostly paint/RMP on spray painting).

The snapshot is stamped with the month you run it for. Re-running a month
DELETES that month's rows entirely and rebuilds them, so there's never any
stale leftover.

Requires: Sales_Register.db must be synced for the month first (that's where
the list of sold FG codes comes from).

Run: python BOM_Register.py            (defaults to the current month)
     python BOM_Register.py 2026 6     (or pass year and month)
"""

import requests
import pandas as pd
import sqlite3
import io
import os
import sys
import time
from datetime import datetime

BASE_URL = "https://unitech.tarka.in"
LOGIN_URL = f"{BASE_URL}/app.php"
DATA_URL = f"{BASE_URL}/app.php"

# ERP login comes from the dashboard (or is asked when run by hand) - nothing is saved in this file
from erp_credentials import get_credentials

HEADERS = {
    "X-Requested-With": "XMLHttpRequest",
    "Referer": f"{BASE_URL}/MainPage.html",
}

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
from costing_config import database_dir   # set in costing\config\settings.json
DATABASE_DIR = database_dir()
os.makedirs(DATABASE_DIR, exist_ok=True)

ANALYTICS_DB = os.path.join(DATABASE_DIR, "Analytics.db")
PRODUCTION_DB = os.path.join(DATABASE_DIR, "Production_Entry.db")
SALES_DB = os.path.join(DATABASE_DIR, "Sales_Register.db")
OTHER_SALES_DB = os.path.join(DATABASE_DIR, "Other_Sales.db")   # hand-uploaded sales (Data Sync > Other Sales)

COMPANIES = [
    {
        "name": "Unitech",
        "client_company": "9",
        "units": [
            {"name": "U1", "showroom": "21"},
            {"name": "U2", "showroom": "25"},
            {"name": "U3", "showroom": "29"},
            {"name": "U4", "showroom": "26"},
        ],
    },
    {
        "name": "Vtech",
        "client_company": "11",
        "units": [
            {"name": "V1", "showroom": "27"},
            {"name": "V2", "showroom": "28"},
        ],
    },
]


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
        print("  Login successful")
        return session
    print("  Login FAILED:", data)
    return None


def download_bom_register(session, showroom, max_retries=3):
    """
    Downloads one unit's BOM register. Large units can take the ERP a long
    time to build, and the server sometimes drops the connection mid-way -
    so we allow a generous timeout and retry a few times before giving up.
    """
    params = {
        "a": "getDownloadReportFile",
        "report_id": 206,
        "schema_name": "BOM~Register",
        "system_module_table_name": "bom_register",
        "showroom": showroom,
        "output_item": 0,
        "input_item": 0,
        "data_type": "Detailed",
        "report_type": "CSV",   # Excel export exhausts the ERP's PHP memory limit
    }

    for attempt in range(1, max_retries + 1):
        try:
            resp = session.get(DATA_URL, params=params, headers=HEADERS, timeout=(30, 900))
            return resp.content
        except requests.exceptions.RequestException as e:
            if attempt < max_retries:
                wait = attempt * 15
                print(f"    Attempt {attempt} failed ({type(e).__name__}) - retrying in {wait}s...")
                time.sleep(wait)
            else:
                print(f"    Attempt {attempt} failed ({type(e).__name__}) - giving up on this unit.")
                return None

    return None


# The BOM Register's columns, in the ERP's order (used only when a file arrives without its header row)
BOM_COLUMNS = ["COMPANY", "UNIT", "Effective Date", "OUTPUT CODE", "OUTPUT ITEM", "OUT UOM", "STD.QTY", "PROCESS",
               "TYPE", "INPUT CODE", "INPUT ITEM", "IN UOM", "NORM QTY", "WIP STORE", "TO STORE", "REFERENCE",
               "WORK LOCATION", "NOTES", "CONSUMPTION REQUIRED", "MOULD ID", "TONNAGE", "RUNNER WT", "PART WT",
               "SHOT WT", "TONNAGE.1", "PAINT TYPE", "MOULD TYPE"]


def parse_bom_csv(file_bytes):
    """
    The ERP puts a few title rows ("BOM REGISTER", "AS ON ...") above the
    column headers, and how many changes from time to time - so the header
    row is found by its contents (OUTPUT CODE + INPUT CODE), not by its
    position. The ERP sometimes appends a PHP fatal-error traceback to the
    end of the file, so rows without both an OUTPUT CODE and INPUT CODE are
    dropped.
    """
    if not file_bytes or not file_bytes.strip():
        print("    The ERP sent an empty file (no BOM for this unit?).")
        return pd.DataFrame()
    try:
        text = file_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = file_bytes.decode("latin-1")

    offset = 0
    header_found = False
    first_data = None          # where the data starts, if the file has no header row at all
    for line in text.splitlines(keepends=True)[:50]:
        upper = line.upper()
        if "OUTPUT CODE" in upper and "INPUT CODE" in upper:
            header_found = True
            break
        if first_data is None and line.lstrip('"').upper().startswith(("UNITECH", "VTECH")) and line.count(",") >= 12:
            first_data = offset
        offset += len(line)

    try:
        if header_found:
            df = pd.read_csv(io.StringIO(text[offset:]), low_memory=False)
        elif first_data is not None:
            print("    (no header row in the file - using the standard BOM Register column order)")
            df = pd.read_csv(io.StringIO(text[first_data:]), header=None, low_memory=False)
            df.columns = (BOM_COLUMNS + [f"extra_{i}" for i in range(len(df.columns))])[:len(df.columns)]
        else:
            first = " | ".join(l.strip()[:60] for l in text.splitlines()[:3])
            print(f"    No BOM rows in the file, skipping. It starts: {first}")
            return pd.DataFrame()
    except Exception as e:
        print(f"    Could not parse CSV: {e}")
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]

    required = {"OUTPUT CODE", "INPUT CODE", "NORM QTY"}
    if not required.issubset(df.columns):
        print(f"    Unexpected columns, skipping. Got: {list(df.columns)[:8]}")
        return pd.DataFrame()

    df = df[df["OUTPUT CODE"].notna() & df["INPUT CODE"].notna()]

    # NORM QTY can come through as text (the ERP quotes everything in the CSV),
    # which would make summing concatenate strings instead of adding numbers.
    df["NORM QTY"] = pd.to_numeric(df["NORM QTY"], errors="coerce").fillna(0)

    return df


def get_sold_fg_codes(month):
    """Distinct FG codes sold in the given month - from the ERP's Sales
    Register AND hand-uploaded Other Sales, so an FG sold only outside
    the ERP still gets a BOM tree (and therefore a cost)."""
    year, mon = month.split("_")
    codes = set()
    for db_path, table in ((SALES_DB, "sales_register"), (OTHER_SALES_DB, "other_sales")):
        if not os.path.exists(db_path):
            continue
        conn = sqlite3.connect(db_path)
        try:
            rows = conn.execute(f"""
                SELECT DISTINCT item_code FROM {table}
                WHERE substr(invoice_date, 4, 2) = ? AND substr(invoice_date, 7, 4) = ?
                  AND item_code IS NOT NULL AND TRIM(item_code) != ''
            """, (mon, year)).fetchall()
            codes.update(r[0] for r in rows)
        except sqlite3.OperationalError:
            pass
        conn.close()
    return codes


def get_item_types(month):
    """{item_code: 'Production'|'Purchase'} for the month, from Item Master."""
    if not os.path.exists(PRODUCTION_DB):
        return {}

    item_type_col = f"item_type_{month}"
    conn = sqlite3.connect(PRODUCTION_DB)
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(item_master)")
    columns = {row[1] for row in cur.fetchall()}
    if item_type_col not in columns:
        conn.close()
        return {}

    cur.execute(f"SELECT erp_code, {item_type_col} FROM item_master")
    result = {code: t for code, t in cur.fetchall() if t}
    conn.close()
    return result


def is_fg_or_sfg(code):
    return "FG" in str(code).upper()


def build_bom_lookup(df):
    """
    {output_code: [ {input_code, norm_qty, process, wip_store, unit, company, ...}, ... ]}

    Duplicate output+input+UNIT rows get their Norm Qty SUMMED (genuinely
    separate BOM lines for the same material within the same unit's
    process really do add up). Duplicate output+input rows that differ
    ONLY by unit do NOT get summed - the same material feeding the same
    output is often recorded once per unit it's made in, each copy
    carrying the SAME norm rate, not an additional requirement, so
    grouping without UNIT would silently double (or more) the true norm
    qty and discard every unit but one. Unit must stay part of the
    grouping key, not get collapsed away.
    """
    grouped = (
        df.groupby(["OUTPUT CODE", "INPUT CODE", "UNIT"], dropna=False)
          .agg({
              "NORM QTY": "sum",
              "COMPANY": "first",
              "PROCESS": "first",
              "WIP STORE": "first",
              "OUTPUT ITEM": "first",
              "INPUT ITEM": "first",
              "OUT UOM": "first",
              "IN UOM": "first",
          })
          .reset_index()
    )

    lookup = {}
    for _, r in grouped.iterrows():
        lookup.setdefault(r["OUTPUT CODE"], []).append({
            "input_code": r["INPUT CODE"],
            "input_item": r["INPUT ITEM"],
            "norm_qty": r["NORM QTY"],
            "company": r["COMPANY"],
            "unit": r["UNIT"],
            "process": r["PROCESS"],
            "wip_store": r["WIP STORE"],
            "output_item": r["OUTPUT ITEM"],
            "out_uom": r["OUT UOM"],
            "in_uom": r["IN UOM"],
        })
    return lookup


_norm_conflicts_warned = set()   # print each unit norm conflict once per run


def walk_tree(root_fg, bom_lookup, item_types):
    """
    Walks down from a sold FG, keeping only the links reachable from it.
    Expands a node only when it's FG/SFG, Production-type, and has its own
    BOM rows. Cycle-protected via the current path.

    A code's OWN children are only ever expanded (added to edges) ONCE per
    walk, tracked separately from the cycle-detection path. Without this,
    a code reached via more than one incoming edge - e.g. the same SFG
    recorded once per unit it's made in, which is legitimate and correctly
    produces multiple edges INTO that code - would have its entire
    downstream sub-tree appended again for every incoming edge, silently
    duplicating everything underneath it. The edges leading INTO a code can
    legitimately repeat (once per unit); what must not repeat is walking
    that same code's own BOM rows more than once.

    Returns a list of edge dicts.
    """
    edges = []
    already_expanded = set()

    def pick_one_per_child(parent, children, preferred_unit):
        """The same parent->child link is often recorded once per unit it's
        made in (U1, U2...). Those are ALTERNATIVE places it gets made, not
        extra quantity needed - keeping all of them makes the dashboard add
        the child's cost once per unit (double-counting). Keep exactly one
        per child: the same unit as the parent's chosen unit if available
        (keeps the chain in one plant), otherwise the first unit
        alphabetically so the choice is stable run to run."""
        by_code = {}
        for c in children:
            by_code.setdefault(c["input_code"], []).append(c)
        chosen = []
        for code, variants in by_code.items():
            variants = sorted(variants, key=lambda v: str(v["unit"]))
            match = [v for v in variants if v["unit"] == preferred_unit]
            pick = match[0] if match else variants[0]
            chosen.append(pick)

            norms = {v["norm_qty"] for v in variants}
            if len(norms) > 1 and (parent, code) not in _norm_conflicts_warned:
                _norm_conflicts_warned.add((parent, code))
                detail = ", ".join(f"{v['unit']}={v['norm_qty']}" for v in variants)
                print(f"    ⚠️  {parent} -> {code}: norm differs by unit ({detail}); using {pick['unit']}={pick['norm_qty']}")
        return chosen

    def expand(parent, level, path, parent_unit=None):
        if parent in path:
            print(f"    Cycle detected at {parent} (path: {' -> '.join(path)}), stopping this branch")
            return
        if parent not in bom_lookup:
            return
        if parent in already_expanded:
            return
        already_expanded.add(parent)

        for child in pick_one_per_child(parent, bom_lookup[parent], parent_unit):
            child_code = child["input_code"]
            child_type = item_types.get(child_code)

            edges.append({
                "root_fg": root_fg,
                "parent_code": parent,
                "child_code": child_code,
                "child_item": child["input_item"],
                "norm_qty": child["norm_qty"],
                "level": level,
                "company": child["company"],
                "unit": child["unit"],
                "process": child["process"],
                "wip_store": child["wip_store"],
                "child_item_type": child_type,
            })

            # Only Production-type FG/SFG codes expand further. Item Master's
            # classification overrides the BOM - a Purchase-type SFG is a leaf
            # even if it has BOM rows of its own.
            if is_fg_or_sfg(child_code) and child_type == "Production":
                expand(child_code, level + 1, path + [parent], child["unit"])

    expand(root_fg, 1, [])
    return edges


def get_connection():
    conn = sqlite3.connect(ANALYTICS_DB)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS bom_tree (
            month TEXT,
            root_fg TEXT,
            parent_code TEXT,
            child_code TEXT,
            child_item TEXT,
            norm_qty REAL,
            level INTEGER,
            company TEXT,
            unit TEXT,
            process TEXT,
            wip_store TEXT,
            child_item_type TEXT
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_bom_tree_month_root ON bom_tree(month, root_fg)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_bom_tree_parent ON bom_tree(month, parent_code)")
    conn.commit()
    return conn


if __name__ == "__main__":
    if len(sys.argv) >= 3:
        year, mon = int(sys.argv[1]), int(sys.argv[2])
    else:
        now = datetime.now()
        year, mon = now.year, now.month
    month = f"{year}_{mon:02d}"

    print(f"Building BOM tree for {month}\n")

    sold_fgs = get_sold_fg_codes(month)
    if not sold_fgs:
        print(f"❌ No FG codes sold in {month} - run Sales Register first. Nothing saved.")
        sys.exit(1)
    print(f"{len(sold_fgs)} distinct FG code(s) sold in {month}\n")

    item_types = get_item_types(month)
    if not item_types:
        print(f"❌ Item Master has no column for {month} yet - run Item Master first, then this again. Nothing saved.")
        sys.exit(1)
    print(f"{len(item_types)} item type(s) loaded from Item Master\n")

    # --- Download the register from every unit and merge ---
    all_frames = []
    for company in COMPANIES:
        print(f"===== {company['name']} =====")
        session = login(company["client_company"])
        if not session:
            continue

        for unit in company["units"]:
            print(f"  Downloading BOM register for {unit['name']}...")
            file_bytes = download_bom_register(session, unit["showroom"])
            if not file_bytes:
                print("    Skipping this unit - no data returned.")
                continue
            df = parse_bom_csv(file_bytes)
            print(f"    {len(df)} valid row(s)")
            if not df.empty:
                all_frames.append(df)
            time.sleep(0.5)

    if not all_frames:
        print("\n❌ No BOM data downloaded from the ERP - nothing saved.")
        sys.exit(1)

    combined = pd.concat(all_frames, ignore_index=True)
    print(f"\nTotal BOM rows across all units: {len(combined)}")

    bom_lookup = build_bom_lookup(combined)
    print(f"{len(bom_lookup)} distinct output code(s) in the register")

    # --- Walk only the sold FGs ---
    print("\nWalking trees for sold FG codes...")
    all_edges = []
    fgs_with_bom = 0
    for fg in sorted(sold_fgs):
        edges = walk_tree(fg, bom_lookup, item_types)
        if edges:
            fgs_with_bom += 1
            all_edges.extend(edges)

    print(f"{fgs_with_bom} of {len(sold_fgs)} sold FG(s) had a BOM")
    print(f"{len(all_edges)} tree edge(s) to store")

    # --- Replace this month's rows entirely ---
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM bom_tree WHERE month = ?", (month,))
    deleted = cur.rowcount
    if deleted:
        print(f"\nCleared {deleted} existing row(s) for {month}")

    cur.executemany("""
        INSERT INTO bom_tree (month, root_fg, parent_code, child_code, child_item,
                              norm_qty, level, company, unit, process, wip_store, child_item_type)
        VALUES (:month, :root_fg, :parent_code, :child_code, :child_item,
                :norm_qty, :level, :company, :unit, :process, :wip_store, :child_item_type)
    """, [{**e, "month": month} for e in all_edges])

    conn.commit()
    conn.close()

    print(f"\nDone. {len(all_edges)} row(s) saved for {month}.")
