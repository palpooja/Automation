"""
core/store.py - everything the whole suite shares on disk.

* Where files go           -> paths()  (settings folder, logs, output root)
* Saved clients            -> load_clients() / save_clients()
  passwords are encrypted with Windows DPAPI (Data Protection API):
  only the same Windows user on the same PC can read them back.
  Code taken unchanged from ura_returns_bot.pyw v4.3 lines 161-238.
* Run history              -> add_history() / load_history()
* Suite settings           -> load_settings() / save_settings()

Old data is picked up automatically: if the suite's own clients.json does
not exist yet but the Returns Downloader's one does
(%APPDATA%\\URA_Returns_Downloader\\clients.json), it is copied over once.
"""
import base64
import ctypes
import datetime
import json
import os
import re
import shutil
import sys
import time
from pathlib import Path

APP_NAME = "URA Suite"

# ---------------------------------------------------------------- paths
def _local_appdata() -> Path:
    return Path(os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
                or Path.home())


DATA_DIR = _local_appdata() / "URA_Suite"
LOG_DIR = DATA_DIR / "logs"
CLIENTS_FILE = DATA_DIR / "clients.json"
HISTORY_FILE = DATA_DIR / "history.json"
SETTINGS_FILE = DATA_DIR / "settings.json"
OLD_RETURNS_CLIENTS = Path(os.environ.get("APPDATA", str(Path.home()))) \
    / "URA_Returns_Downloader" / "clients.json"
DEFAULT_OUTPUT_ROOT = Path.home() / "Downloads" / "URA_Suite"


def app_dir() -> Path:
    """Folder the program lives in (next to the .exe once built)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


def resource(name: str) -> Path:
    """A file packed into the .exe (icon etc.) or next to the source."""
    base = Path(getattr(sys, "_MEIPASS", app_dir()))
    return base / name


def ensure_dirs():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------- DPAPI
def _dpapi(data: bytes, protect: bool) -> bytes:
    if os.name != "nt":                     # not Windows: plain bytes (dev only)
        return data

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", ctypes.c_uint32),
                    ("pbData", ctypes.POINTER(ctypes.c_char))]
    crypt = ctypes.windll.crypt32
    kernel = ctypes.windll.kernel32
    buf = ctypes.create_string_buffer(data, len(data))
    inp = BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    out = BLOB()
    fn = crypt.CryptProtectData if protect else crypt.CryptUnprotectData
    if not fn(ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out)):
        raise OSError("DPAPI call failed")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        kernel.LocalFree(out.pbData)


def protect_secret(text: str) -> str:
    return base64.b64encode(_dpapi(text.encode("utf-8"), True)).decode("ascii")


def unprotect_secret(blob: str) -> str:
    """Decrypted text, or "" if it cannot be decrypted (see try_unprotect)."""
    value = try_unprotect(blob)
    return "" if value is None else value


def try_unprotect(blob: str):
    """Decrypted text, or None when this Windows account cannot decrypt it
    (file copied from another PC or another Windows user, profile reset)."""
    if not blob:
        return ""
    try:
        return _dpapi(base64.b64decode(blob, validate=True), False).decode("utf-8")
    except Exception:
        return None


def credential_store_kind() -> str:
    return ("Windows account encryption (DPAPI)" if os.name == "nt"
            else "encoded only - not encrypted on this system")


# ---------------------------------------------------------------- json
def _load_json(path: Path, default):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return default


def _save_json(path: Path, data, backup: bool = False):
    ensure_dirs()
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2)
    if backup and path.exists():
        shutil.copy2(path, path.with_name(path.name + ".bak"))
    os.replace(tmp, path)


# ---------------------------------------------------------------- clients
LOAD_WARNINGS = []      # messages for the window (damaged file, unreadable passwords)


def load_clients():
    """[{name, login, password}] with the password decrypted for use.

    Never loses data:
    * a password this Windows account cannot decrypt comes back as "" with
      password_unreadable=True, and its encrypted text is kept (``_blob``) so
      the next save writes it back unchanged;
    * a damaged clients.json is renamed clients.damaged-<date>.json (kept, not
      overwritten) and an empty list is returned, with a message in LOAD_WARNINGS."""
    LOAD_WARNINGS.clear()
    if not CLIENTS_FILE.exists() and OLD_RETURNS_CLIENTS.exists():
        try:                                # one-time import from the old program
            ensure_dirs()
            shutil.copy2(OLD_RETURNS_CLIENTS, CLIENTS_FILE)
        except OSError as e:
            LOAD_WARNINGS.append(f"Could not copy the old Returns Downloader clients ({e}).")
    if not CLIENTS_FILE.exists():
        return []
    try:
        raw = json.loads(CLIENTS_FILE.read_text(encoding="utf-8"))
        if not isinstance(raw, list):
            raise ValueError("not a list")
    except (OSError, ValueError) as e:
        kept = CLIENTS_FILE.with_name(
            f"clients.damaged-{datetime.datetime.now():%Y%m%d_%H%M%S}.json")
        try:
            os.replace(CLIENTS_FILE, kept)
            where = f"It was kept as {kept.name}"
        except OSError:
            where = "It could not be renamed"
        LOAD_WARNINGS.append(f"The saved client list could not be read ({type(e).__name__}). "
                             f"{where} in {DATA_DIR}; clients.json.bak there is the copy "
                             "before the last change.")
        return []
    out, unreadable = [], []
    for c in raw:
        blob = c.get("password", "")
        pw = try_unprotect(blob)
        entry = {"name": c.get("name", ""), "login": c.get("login", ""), "password": pw or ""}
        if pw is None:
            entry["password_unreadable"] = True
            entry["_blob"] = blob
            unreadable.append(entry["name"] or entry["login"])
        out.append(entry)
    if unreadable:
        LOAD_WARNINGS.append("These passwords cannot be decrypted by this Windows account and "
                             "must be typed again (Clients > Edit): " + ", ".join(unreadable)
                             + ". They are kept unchanged until you do.")
    return out


def save_clients(clients):
    rows = []
    for c in clients:
        if c.get("password_unreadable") and not c.get("password") and c.get("_blob"):
            stored = c["_blob"]               # still unreadable: keep the original text
        else:
            stored = protect_secret(c["password"])
        rows.append({"name": c["name"], "login": c["login"], "password": stored})
    _save_json(CLIENTS_FILE, rows, backup=True)


TIN_RE = re.compile(r"\d+")


def client_problem(client: dict, others) -> str:
    """Why this client cannot be saved, or None.  `others` = the other clients."""
    missing = [lbl for k, lbl in (("name", "client name"), ("login", "Login ID"),
                                  ("password", "password")) if not str(client.get(k, "")).strip()]
    if missing:
        return "Fill in the " + ", ".join(missing) + "."
    login = str(client["login"]).strip()
    if not TIN_RE.fullmatch(login):
        return "The Login ID (TIN) must be digits only - no spaces or other characters."
    if any(str(o.get("login", "")).strip() == login for o in others):
        return f"Login ID {login} is already in the client list."
    return None


# ---------------------------------------------------------------- history
def load_history():
    return _load_json(HISTORY_FILE, [])


def mask(text, secrets) -> str:
    """Replace every known password in text with ***."""
    text = str(text)
    for sec in secrets or ():
        if sec and len(sec) >= 2:
            text = text.replace(sec, "***")
    return text


def add_history(entry: dict, secrets=()):
    hist = load_history()
    hist.append({k: (mask(v, secrets) if isinstance(v, str) else v) for k, v in entry.items()})
    _save_json(HISTORY_FILE, hist[-1000:])


# ---------------------------------------------------------------- log folder clean-up
def prune_logs(days: int) -> int:
    """Delete files (screenshots, page copies, logs) and emptied folders in the
    log folder that are older than `days`.  0 = keep everything.  Returns the
    number of files deleted.  Runs at start-up, so it costs no run time."""
    if not days or days <= 0 or not LOG_DIR.exists():
        return 0
    cutoff = time.time() - days * 86400
    removed = 0
    for path in sorted(LOG_DIR.rglob("*"), key=lambda p: len(p.parts), reverse=True):
        try:
            if path.is_file() and path.stat().st_mtime < cutoff:
                path.unlink()
                removed += 1
            elif path.is_dir() and not any(path.iterdir()):
                path.rmdir()
        except OSError:
            continue                          # in use (e.g. today's log) - try next time
    return removed


# ---------------------------------------------------------------- settings
DEFAULT_SETTINGS = {
    "output_root": str(DEFAULT_OUTPUT_ROOT),
    "modules": {},            # per-module saved options, keyed by module id
}


def load_settings() -> dict:
    s = dict(DEFAULT_SETTINGS)
    s.update(_load_json(SETTINGS_FILE, {}))
    s.setdefault("modules", {})
    return s


def save_settings(settings: dict):
    _save_json(SETTINGS_FILE, settings)
