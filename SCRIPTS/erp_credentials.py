r"""
Gives every script its ERP (Vyapak) login. In order:

  1. Started by the dashboard -> the login the dashboard hands over for
     that one run (never written anywhere by this file).
  2. A login saved on this PC with deploy\Set-ERP-Login.bat -> stored
     encrypted in costing\config\erp_login.dat (Windows protection; the
     file can't be read on any other PC). This is what the plant server uses.
  3. Run by hand with nothing saved -> it asks, once per run. The password
     is hidden while you type.

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

    # 2. Saved on this PC with deploy\Set-ERP-Login.bat.
    try:
        from costing_config import load_erp_login
        saved = load_erp_login()
    except Exception:
        saved = None
    if saved:
        _cached = saved
        return _cached

    # 3. Run by hand - ask now.
    if not (sys.stdin and sys.stdin.isatty()):
        print("❌ No ERP login given. Run this script from the dashboard's Run button, "
              "or open it in a Command Prompt so it can ask for your username and password.")
        sys.exit(1)

    print("ERP login (not saved - run deploy\\Set-ERP-Login.bat to save it on this PC)")
    while True:
        username = input("  Username: ").strip()
        password = getpass.getpass("  Password (hidden while typing): ")
        if username and password:
            break
        print("  Username and password can't be blank - try again.")

    _cached = (username, password)
    return _cached
