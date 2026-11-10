r"""
Settings shared by the dashboard and every script.

Everything that differs between PCs lives in costing\\config\\ (never on
GitHub):

  config\\settings.json   where the databases are, the port, the dashboard
                          password (stored only as a hash)
  config\\erp_login.dat   the ERP username and password, encrypted with
                          Windows' own protection (DPAPI). The file is
                          useless on any other PC.

Set them with deploy\\Set-ERP-Login.bat and deploy\\Set-Password.bat.
"""

import base64
import json
import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
CONFIG_DIR = os.path.join(ROOT, "config")
SETTINGS_FILE = os.path.join(CONFIG_DIR, "settings.json")
ERP_LOGIN_FILE = os.path.join(CONFIG_DIR, "erp_login.dat")

DEFAULTS = {
    "database_dir": "",        # blank = costing\DATABASE
    "downloads_dir": "",       # blank = costing\DOWNLOADS
    "host": "0.0.0.0",         # 0.0.0.0 = reachable from other PCs on the plant network
    "port": 5000,
    "auto_sync": True,         # timed ERP syncs (see DASHBOARD\scheduler.py)
    "backup_dir": "",          # blank = costing\backups
    "backup_keep_days": 7,
    "password_hash": "",
    "secret_key": "",
}


# ---------- settings.json ----------

def load_settings():
    data = dict(DEFAULTS)
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            data.update(json.load(f))
    except FileNotFoundError:
        pass
    return data


def save_settings(changes):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    current = {}
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            current = json.load(f)
    except FileNotFoundError:
        pass
    current.update(changes)
    tmp = SETTINGS_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(current, f, indent=2)
    os.replace(tmp, SETTINGS_FILE)


def _folder(setting, env_name, default_name):
    path = os.environ.get(env_name) or load_settings().get(setting) or os.path.join(ROOT, default_name)
    os.makedirs(path, exist_ok=True)
    return path


def database_dir():
    """Folder holding the .db files. COSTING_DATABASE_DIR (set by the
    dashboard for the scripts it starts) wins, then settings.json."""
    return _folder("database_dir", "COSTING_DATABASE_DIR", "DATABASE")


def downloads_dir():
    return _folder("downloads_dir", "COSTING_DOWNLOADS_DIR", "DOWNLOADS")


def backup_dir():
    return _folder("backup_dir", "COSTING_BACKUP_DIR", "backups")


# ---------- ERP login (Windows DPAPI, machine scope) ----------
# Machine scope so the dashboard can read it when Windows starts it at boot
# (as SYSTEM, before anyone signs in). Only this PC can decrypt the file.

# Calls Windows' crypt32.dll directly (ctypes, part of Python), so nothing
# extra has to be installed.

_CRYPTPROTECT_UI_FORBIDDEN = 0x1
_CRYPTPROTECT_LOCAL_MACHINE = 0x4
_DPAPI_FLAGS = _CRYPTPROTECT_UI_FORBIDDEN | _CRYPTPROTECT_LOCAL_MACHINE


def _dpapi():
    if sys.platform != "win32":
        raise RuntimeError("Saving or reading the ERP login needs Windows.")
    import ctypes
    from ctypes import wintypes

    class DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    blob_p = ctypes.POINTER(DATA_BLOB)
    crypt32.CryptProtectData.argtypes = [blob_p, wintypes.LPCWSTR, blob_p, ctypes.c_void_p,
                                         ctypes.c_void_p, wintypes.DWORD, blob_p]
    crypt32.CryptProtectData.restype = wintypes.BOOL
    crypt32.CryptUnprotectData.argtypes = [blob_p, ctypes.c_void_p, blob_p, ctypes.c_void_p,
                                           ctypes.c_void_p, wintypes.DWORD, blob_p]
    crypt32.CryptUnprotectData.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p

    def run(func, data, flags, *middle):
        buf = ctypes.create_string_buffer(data, len(data))
        blob_in = DATA_BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
        blob_out = DATA_BLOB()
        if not func(ctypes.byref(blob_in), *middle, None, None, None, flags, ctypes.byref(blob_out)):
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            return ctypes.string_at(blob_out.pbData, blob_out.cbData)
        finally:
            kernel32.LocalFree(ctypes.cast(blob_out.pbData, ctypes.c_void_p))

    return crypt32, run


def _protect(raw):
    crypt32, run = _dpapi()
    return run(crypt32.CryptProtectData, raw, _DPAPI_FLAGS, "Unitech costing ERP login")


def _unprotect(blob):
    crypt32, run = _dpapi()
    return run(crypt32.CryptUnprotectData, blob, _CRYPTPROTECT_UI_FORBIDDEN, None)


def save_erp_login(username, password):
    """Saves the login, then reads it back to make sure it can be decrypted."""
    os.makedirs(CONFIG_DIR, exist_ok=True)
    raw = json.dumps({"username": username, "password": password}).encode("utf-8")
    blob = _protect(raw)
    if _unprotect(blob) != raw:
        raise RuntimeError("Windows encrypted the login but couldn't read it back - nothing saved.")
    with open(ERP_LOGIN_FILE, "wb") as f:
        f.write(base64.b64encode(blob))


def load_erp_login():
    """(username, password) saved on this PC, or None."""
    try:
        with open(ERP_LOGIN_FILE, "rb") as f:
            data = json.loads(_unprotect(base64.b64decode(f.read())).decode("utf-8"))
        if data.get("username") and data.get("password"):
            return data["username"], data["password"]
    except FileNotFoundError:
        return None
    except Exception as e:
        print(f"⚠️  Couldn't read the saved ERP login ({e}). Run deploy\\Set-ERP-Login.bat again.")
    return None


def check_erp_login(username, password):
    """True = ERP accepted it, False = ERP rejected it, None = couldn't reach the ERP."""
    try:
        import requests
        session = requests.Session()
        session.get("https://unitech.tarka.in", timeout=20)
        resp = session.post(
            "https://unitech.tarka.in/app.php",
            params={"a": "validateUsernameAndPassword.demo"},
            data={"usr": username, "password": password, "client_company": "9"},
            headers={"X-Requested-With": "XMLHttpRequest",
                     "Referer": "https://unitech.tarka.in/MainPage.html"},
            timeout=20,
        )
        return resp.json().get("status") == "success"
    except Exception:
        return None
