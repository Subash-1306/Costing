"""
Asks for the ERP (Vyapak) username and password - nothing is saved anywhere.

Every script that logs in to the ERP gets its login from here:

  * Run from the dashboard's "Run" button -> uses the login you typed when
    you started the dashboard (kept in memory until you close the dashboard).
    Nothing is written to any file.
  * Run by hand (double-click or Command Prompt) -> it asks here, once per
    run. The password is hidden while you type.

The login is remembered only while the script is running, so a script that
logs in to several companies asks you just once.
"""

import getpass
import os
import sys

_cached = None


def get_credentials():
    """Returns (username, password) for this run."""
    global _cached
    if _cached:
        return _cached

    # 1. Started by the dashboard - it passes the login for this run only.
    username = os.environ.get("ERP_USERNAME", "").strip()
    password = os.environ.get("ERP_PASSWORD", "")
    if username and password:
        _cached = (username, password)
        return _cached

    # 2. Run by hand - ask now.
    if not (sys.stdin and sys.stdin.isatty()):
        print("❌ No ERP login given. Run this script from the dashboard's Run button, "
              "or open it in a Command Prompt so it can ask for your username and password.")
        sys.exit(1)

    print("ERP login (nothing is saved)")
    while True:
        username = input("  Username: ").strip()
        password = getpass.getpass("  Password (hidden while typing): ")
        if username and password:
            break
        print("  Username and password can't be blank - try again.")

    _cached = (username, password)
    return _cached
