"""
Runs the ERP syncs on a timer, so nobody has to press Run.

One worker runs the jobs one after another (never two at once, so the
scripts don't fight over the same database). A job that is already waiting
or running is not queued twice. Jobs that need the ERP login are skipped,
with a note in the log, until a login is saved (deploy\\Set-ERP-Login.bat).

The schedule lives in SCHEDULE below; change the times there.
"""

import queue
import threading
import time
import traceback
from datetime import datetime, timedelta

# kind "every": every `minutes`, between `start` and `end` (24-hour clock)
# kind "daily": once a day at `at`
# args: "current_month" -> the script gets this month's year and month;
#       "current_and_previous_month" -> also last month, during the first 10 days
#       (late postings for the month that just ended still land then)
SCHEDULE = [
    {"key": "sales_register",         "kind": "every", "minutes": 15, "start": "07:00", "end": "23:00"},
    {"key": "grn_register",           "kind": "every", "minutes": 15, "start": "07:00", "end": "23:00"},
    {"key": "production_entry",       "kind": "every", "minutes": 15, "start": "07:00", "end": "23:00"},
    {"key": "issuance_report",        "kind": "every", "minutes": 60, "start": "07:30", "end": "23:00"},
    {"key": "transfer_register",      "kind": "every", "minutes": 60, "start": "07:30", "end": "23:00"},
    {"key": "return_register",        "kind": "every", "minutes": 60, "start": "07:30", "end": "23:00"},
    {"key": "weighted_price_master",  "kind": "every", "minutes": 60, "start": "07:45", "end": "23:00"},
    {"key": "weighted_selling_price", "kind": "every", "minutes": 60, "start": "07:45", "end": "23:00"},
    {"key": "item_master",            "kind": "daily", "at": "01:00"},
    {"key": "mould_master",           "kind": "daily", "at": "01:05"},
    {"key": "vendor_price_master",    "kind": "daily", "at": "01:10"},
    {"key": "opncls_stock",           "kind": "daily", "at": "01:20", "args": "current_and_previous_month"},
    {"key": "bom_register",           "kind": "daily", "at": "01:40", "args": "current_and_previous_month"},
    {"key": "backup",                 "kind": "daily", "at": "00:30"},
]


def _hm(text):
    h, m = text.split(":")
    return int(h), int(m)


def next_run(entry, now):
    """The next time this entry is due, strictly after `now`."""
    if entry["kind"] == "daily":
        h, m = _hm(entry["at"])
        due = now.replace(hour=h, minute=m, second=0, microsecond=0)
        return due if due > now else due + timedelta(days=1)

    sh, sm = _hm(entry["start"])
    eh, em = _hm(entry["end"])
    step = timedelta(minutes=entry["minutes"])
    day_start = now.replace(hour=sh, minute=sm, second=0, microsecond=0)
    day_end = now.replace(hour=eh, minute=em, second=0, microsecond=0)
    if now < day_start:
        return day_start
    if now >= day_end:
        return day_start + timedelta(days=1)
    steps = int((now - day_start) / step) + 1
    due = day_start + steps * step
    return due if due <= day_end else day_start + timedelta(days=1)


def month_args(mode, now):
    """List of argument lists for one entry, e.g. [["2026", "11"]]."""
    if not mode:
        return [[]]
    runs = [[str(now.year), str(now.month)]]
    if mode == "current_and_previous_month" and now.day <= 10:
        prev = now.replace(day=1) - timedelta(days=1)
        runs.insert(0, [str(prev.year), str(prev.month)])
    return runs


def describe(key):
    """Short text for the Data Sync card, e.g. 'Auto: every 15 min, 07:00-23:00'."""
    parts = []
    for e in SCHEDULE:
        if e["key"] != key:
            continue
        if e["kind"] == "daily":
            parts.append(f"daily at {e['at']}")
        else:
            parts.append(f"every {e['minutes']} min, {e['start']}-{e['end']}")
    return ("Auto: " + "; ".join(parts)) if parts else ""


class Scheduler:
    """run_job(key, args, trigger) must run one job to the end and return
    its exit code. can_run(key) returns (True, "") or (False, reason)."""

    def __init__(self, run_job, can_run, log=print, clock=datetime.now):
        self.run_job = run_job
        self.can_run = can_run
        self.log = log
        self.clock = clock
        self.queue = queue.Queue()
        self.pending = set()            # (key, args) waiting or running
        self.lock = threading.Lock()
        self.next_due = {}
        self.stopped = threading.Event()

    def start(self):
        now = self.clock()
        for i, e in enumerate(SCHEDULE):
            self.next_due[i] = next_run(e, now)
        threading.Thread(target=self._tick_loop, daemon=True, name="scheduler-tick").start()
        threading.Thread(target=self._work_loop, daemon=True, name="scheduler-work").start()
        self.log("Auto-sync started. Next runs: " + ", ".join(
            f"{SCHEDULE[i]['key']} {d:%H:%M}" for i, d in sorted(self.next_due.items(), key=lambda x: x[1])[:4]))

    def stop(self):
        self.stopped.set()

    def enqueue(self, key, args):
        item = (key, tuple(args))
        with self.lock:
            if item in self.pending:
                return False
            self.pending.add(item)
        self.queue.put(item)
        return True

    def tick(self):
        """Queues every entry that has come due. Called every 30 s."""
        now = self.clock()
        for i, e in enumerate(SCHEDULE):
            if now >= self.next_due[i]:
                for args in month_args(e.get("args"), now):
                    self.enqueue(e["key"], args)
                self.next_due[i] = next_run(e, now)

    def _tick_loop(self):
        while not self.stopped.is_set():
            try:
                self.tick()
            except Exception:
                self.log("Scheduler tick failed:\n" + traceback.format_exc())
            self.stopped.wait(30)

    def run_one(self, item):
        key, args = item
        try:
            ok, reason = self.can_run(key)
            if not ok:
                self.log(f"Auto-sync skipped {key}: {reason}")
                return None
            started = time.time()
            code = self.run_job(key, list(args), "auto")
            self.log(f"Auto-sync {key} {' '.join(args)} finished with code {code} in {time.time() - started:.0f} s")
            return code
        except Exception:
            self.log(f"Auto-sync {key} crashed:\n" + traceback.format_exc())
            return 1
        finally:
            with self.lock:
                self.pending.discard(item)

    def _work_loop(self):
        while not self.stopped.is_set():
            try:
                item = self.queue.get(timeout=5)
            except queue.Empty:
                continue
            self.run_one(item)
