r"""Sets the password for opening the dashboard in a browser
(deploy\Set-Password.bat runs this). Only a hash of it is stored, in
costing\config\settings.json. Takes effect straight away; no restart needed."""

import getpass
import os
import sys

from werkzeug.security import generate_password_hash

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "SCRIPTS"))
import costing_config  # noqa: E402


def main():
    print("Dashboard password (asked when the dashboard is opened from a browser)\n")
    for _ in range(3):
        first = getpass.getpass("  New password (hidden, 8+ characters): ")
        if len(first) < 8:
            print("  Too short - use at least 8 characters.\n")
            continue
        if getpass.getpass("  Type it again: ") != first:
            print("  The two didn't match.\n")
            continue
        costing_config.save_settings({"password_hash": generate_password_hash(first)})
        print("\n  ✅ Password set. Browsers that were signed in stay signed in.")
        return 0
    print("Nothing changed.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
