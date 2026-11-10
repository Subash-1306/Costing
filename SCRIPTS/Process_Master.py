"""
Import (or update) the Process Master into the database.
Reads PROCESS_MASTER.csv from the same folder as this script.

Run this any time you edit the CSV, to sync changes into the database:
    python import_process_master.py

Later, when this becomes a local site, this same logic can be triggered by
an "upload master" button instead of reading a CSV file from disk.
"""

import sqlite3
import csv

import os
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_DIR = os.path.join(SCRIPT_DIR, "..", "DATABASE")
os.makedirs(DATABASE_DIR, exist_ok=True)
DB_FILE = os.path.join(DATABASE_DIR, "Production_Entry.db")
CSV_FILE = os.path.join(SCRIPT_DIR, "PROCESS_MASTER.csv")


def get_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS process_master (
            wip_store TEXT PRIMARY KEY,
            process TEXT
        )
    """)
    conn.commit()
    return conn


def import_csv(conn, csv_path):
    count = 0
    with open(csv_path, "r", encoding="utf-8-sig") as f:  # utf-8-sig handles the BOM at the file start
        reader = csv.DictReader(f)
        for row in reader:
            wip_store = row.get("WIP STORE", "").strip()
            process = row.get("Process", "").strip()
            if not wip_store:
                continue
            conn.execute(
                "INSERT INTO process_master (wip_store, process) VALUES (?, ?) "
                "ON CONFLICT(wip_store) DO UPDATE SET process = excluded.process",
                (wip_store, process),
            )
            count += 1
    conn.commit()
    return count


if __name__ == "__main__":
    conn = get_connection()
    count = import_csv(conn, CSV_FILE)
    print(f"✅ Imported/updated {count} rows into process_master.")

    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM process_master")
    print(f"Total entries in process_master: {cur.fetchone()[0]}")

    conn.close()
