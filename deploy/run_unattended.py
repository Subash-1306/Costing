"""Keeps the costing dashboard running on the server: starts it, restarts it
after a crash or a hang, and writes its output to logs\\dashboard.log.
Windows starts this at boot (task "Unitech Costing", made by Setup.bat)."""

import logging
import os
import socket
import subprocess
import sys
import threading
import time
import urllib.request
from logging.handlers import RotatingFileHandler
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]          # the costing folder
DASHBOARD = ROOT / "DASHBOARD"
sys.path.insert(0, str(ROOT / "SCRIPTS"))
import costing_config  # noqa: E402


def main():
    guard = socket.socket()
    try:
        guard.bind(("127.0.0.1", 17072))              # one copy only
    except OSError:
        return
    (ROOT / "logs").mkdir(exist_ok=True)
    log = logging.getLogger("costing")
    log.setLevel(logging.INFO)
    handler = RotatingFileHandler(ROOT / "logs" / "dashboard.log", maxBytes=5_000_000, backupCount=5, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(message)s"))
    log.addHandler(handler)

    while True:
        port = int(costing_config.load_settings().get("port", 5000))
        log.info("Starting the costing dashboard")
        env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
        child = subprocess.Popen([sys.executable, "-u", "ProjectDashboard.py"], cwd=DASHBOARD, env=env,
                                 stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                 text=True, encoding="utf-8", errors="replace")

        def watch(p):
            time.sleep(90)
            fails = 0
            while p.poll() is None:
                try:
                    with urllib.request.urlopen(f"http://127.0.0.1:{port}/health", timeout=15) as r:
                        fails = 0 if r.status == 200 else fails + 1
                except Exception:
                    fails += 1
                if fails >= 3 and p.poll() is None:
                    log.error("Not answering for 90 s; restarting")
                    p.terminate()
                    return
                time.sleep(30)

        threading.Thread(target=watch, args=(child,), daemon=True).start()
        for line in child.stdout:
            log.info(line.rstrip())
        log.info("The dashboard stopped (%s); restarting in 15 s", child.wait())
        time.sleep(15)


if __name__ == "__main__":
    main()
