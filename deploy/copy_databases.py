r"""Copies the database (.db) files from another folder into this install's
database folder - safely (deploy\Copy-Databases.bat runs this).

A plain Windows copy of a .db file can come out damaged ("database disk
image is malformed") when the file is copied while a dashboard or a sync is
using it: in WAL mode the newest changes sit in a separate -wal file that
the copy misses. This uses SQLite's own backup instead, which always gives a
complete, consistent copy, and checks every file before and after.

    Copy-Databases.bat                       asks for the folder
    python copy_databases.py "D:\Unitech Dashboard Database\DATABASE"
"""

import os
import socket
import sqlite3
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "SCRIPTS"))
import costing_config  # noqa: E402


def check(path):
    """'ok' or SQLite's description of the damage."""
    try:
        conn = sqlite3.connect(path, timeout=60)
        result = conn.execute("PRAGMA quick_check").fetchone()[0]
        conn.close()
        return result
    except sqlite3.DatabaseError as e:
        return str(e)


def dashboard_running_here():
    port = int(costing_config.load_settings().get("port", 5000))
    with socket.socket() as s:
        s.settimeout(1)
        return s.connect_ex(("127.0.0.1", port)) == 0


def main():
    source = sys.argv[1] if len(sys.argv) > 1 else input("Folder to copy the .db files FROM: ")
    source = source.strip().strip('"')
    target = costing_config.database_dir()
    if not os.path.isdir(source):
        print(f"❌ Folder not found: {source}")
        return 1
    if os.path.normcase(os.path.abspath(source)) == os.path.normcase(os.path.abspath(target)):
        print("❌ That is already this install's database folder - nothing to copy.")
        return 1
    if dashboard_running_here():
        print("❌ A dashboard is running on this PC. Close it first (or stop the 'Unitech Costing' task),")
        print("   so nothing writes to the files while they are replaced.")
        return 1

    # Sync_Log.db is this install's own run history - keep it, don't copy another PC's over it
    names = sorted(f for f in os.listdir(source)
                   if f.lower().endswith(".db") and "_backup_" not in f.lower() and f.lower() != "sync_log.db")
    if not names:
        print(f"❌ No .db files in {source}")
        return 1
    print(f"Copying {len(names)} database(s)\n  from {source}\n  to   {target}\n")

    failed = 0
    for name in names:
        src_path, dst_path = os.path.join(source, name), os.path.join(target, name)
        status = check(src_path)
        if status != "ok":
            print(f"  ❌ {name}: the SOURCE file is damaged ({status}) - not copied")
            failed += 1
            continue
        started = time.time()
        tmp_path = dst_path + ".copying"
        for leftover in (tmp_path, tmp_path + "-journal"):
            if os.path.exists(leftover):
                os.remove(leftover)
        src = sqlite3.connect(src_path, timeout=60)
        dst = sqlite3.connect(tmp_path)
        with dst:
            src.backup(dst)
        src.close()
        dst.close()
        if check(tmp_path) != "ok":
            os.remove(tmp_path)
            print(f"  ❌ {name}: the copy failed its check - the old file was left as it was")
            failed += 1
            continue
        for old in (dst_path + "-wal", dst_path + "-shm"):
            if os.path.exists(old):
                os.remove(old)                     # left over from the damaged file
        os.replace(tmp_path, dst_path)
        size = os.path.getsize(dst_path) / 1_048_576
        print(f"  ✅ {name}  ({size:,.0f} MB, {time.time() - started:.0f} s, checked OK)")

    print("\nDone." if not failed else f"\nDone, but {failed} file(s) were not copied - see above.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
