r"""Saves the ERP login on this PC (deploy\Set-ERP-Login.bat runs this).

The username and password are checked with the ERP, then stored encrypted in
costing\config\erp_login.dat with Windows' own protection (DPAPI): only this
PC can read the file. Nothing goes to GitHub."""

import getpass
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "SCRIPTS"))
import costing_config  # noqa: E402


def main():
    print("ERP (Tarka) login for the costing dashboard")
    print("Use the read-only account made for the server. It is saved encrypted on this PC only.\n")
    for _ in range(3):
        username = input("  Username: ").strip()
        password = getpass.getpass("  Password (hidden while typing): ")
        if not username or not password:
            print("  Username and password can't be blank.\n")
            continue
        print("  Checking with the ERP...")
        ok = costing_config.check_erp_login(username, password)
        if ok is False:
            print("  ❌ The ERP rejected that username/password. Try again.\n")
            continue
        if ok is None:
            print("  ⚠️  Couldn't reach the ERP to check it (internet down?). Saving it anyway.")
        costing_config.save_erp_login(username, password)
        print(f"\n  ✅ Saved in {costing_config.ERP_LOGIN_FILE}")
        print("  Run Restart.bat so the running dashboard picks it up.")
        return 0
    print("Nothing saved.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
