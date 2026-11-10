# Unitech costing dashboard

Costing and margin analytics built from the Tarka ERP: Overview, Part-wise margin, BOM tree, Part cost calc, Month posting, Weighted price, Cycle time, Other cost and Data sync.

On the plant server it runs by itself: Windows starts it at boot, it pulls the ERP data on a timer, backs up the databases every night and restarts after a crash. It is for one person only and opens with a single password.

## Folders

| Folder | What's in it | On GitHub? |
|---|---|---|
| `DASHBOARD` | The web app (`ProjectDashboard.py`), its pages and the timer (`scheduler.py`) | Yes |
| `SCRIPTS` | One script per ERP register or master, plus `costing_config.py` (shared settings) | Yes |
| `deploy` | Setup, restart, ERP login and password tools for Windows | Yes |
| `config` | `settings.json` and the encrypted ERP login for this PC | **No** |
| `DATABASE` | The `.db` files (unless `settings.json` points somewhere else) | **No** |
| `backups`, `logs`, `venv`, `DOWNLOADS` | Made on each PC | **No** |

## Set up the server (Windows)

1. Install **Python 3.10 or newer (64-bit)** from python.org (tick "Add to PATH").
2. Download this repository from GitHub (green **Code** button, then **Download ZIP**), unzip it to a local disk and rename the folder to `costing`, for example `C:\Plant Apps\costing`.
3. Run Setup first (step 4), then bring the databases over with **`deploy\Copy-Databases.bat`** (see *Copying the databases* below).
4. Double-click `deploy\Setup.bat` and approve the administrator prompt. It:
   - installs the libraries into a private Python environment (`venv`)
   - asks where the database files are
   - asks for the **ERP login** (see below) and the **dashboard password**
   - registers the Windows task **"Unitech Costing"**: it starts at boot before anyone signs in and restarts after a failure
   - opens the port (5000) to the plant network only, and turns off sleep
5. Open `http://<server address>:5000` from your laptop and sign in.

## Where the ERP login goes

Double-click **`deploy\Set-ERP-Login.bat`** on the server (Setup runs it once for you). It asks for the username and password, checks them with the ERP, and saves them encrypted in `config\erp_login.dat` using Windows' own protection (DPAPI). Only that server can decrypt the file, and it is never on GitHub. Then run `deploy\Restart.bat`.

Use a **read-only account that can only download reports**. To change the login later, run `Set-ERP-Login.bat` again, then `Restart.bat`.

## The dashboard password

Run **`deploy\Set-Password.bat`** to set or change it. Only a hash is stored, in `config\settings.json`. Until a password is set, the dashboard opens only on the server itself.

## Automatic syncs

The Data Sync page shows each script's schedule and its last run (manual or automatic). The times are in `DASHBOARD\scheduler.py`:

| What | When |
|---|---|
| Sales Register, GRN Register, Production Entry | Every 15 minutes, 07:00 to 23:00 |
| Issuance, Transfer, Return, Weighted Price Master, Weighted Selling Price | Every hour, 07:30 to 23:00 |
| Item Master, Mould Master, Vendor Price Master | Every night, 01:00 to 01:10 |
| Opening/Closing Stock, BOM Register | Every night: this month, and last month during the first 10 days |
| Database backup | Every night at 00:30, newest 7 days kept in `backups\` |

Only one script runs at a time. A script you start by hand with **Run** still works as before. To switch the timer off, set `"auto_sync": false` in `config\settings.json` and run `Restart.bat`.

## `config\settings.json`

```json
{
  "database_dir": "C:\\Plant Apps\\costing\\DATABASE",
  "downloads_dir": "",
  "host": "0.0.0.0",
  "port": 5000,
  "auto_sync": true,
  "backup_dir": "",
  "backup_keep_days": 7
}
```

A blank folder means the default inside `costing`. `host` `127.0.0.1` keeps the dashboard to this PC only. After any change, run `Restart.bat`.

## Copying the databases

Never copy `.db` files with Windows Explorer while a dashboard or a sync might be using them. The databases run in WAL mode, so the newest changes sit in a separate `-wal` file, and a plain copy can come out damaged (`database disk image is malformed`).

Use **`deploy\Copy-Databases.bat`** instead:
1. Close every dashboard that uses the target folder (on the server: stop the "Unitech Costing" task, or run it before Setup's last step).
2. Run `Copy-Databases.bat` and type the folder to copy **from**, for example a USB drive or `\\laptop\D$\Unitech Dashboard Database\DATABASE`.
3. It checks each source file, copies it with SQLite's own backup, checks the copy, and only then replaces the file in this install's database folder.

## Testing on a laptop

Double-click `deploy\Run-In-This-Window.bat`. The dashboard runs in that window (no Windows task) and stops when you close it. With no saved ERP login, it asks for one in the window, as before.

## Logs

`logs\dashboard.log` holds everything the dashboard and the scripts print, including each automatic run. `logs\setup.log` holds the setup.
