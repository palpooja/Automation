"""
core/net.py - pause while the internet (or the portal's server) is unreachable.

Called before every retry and before each module: if ura.go.ug cannot be
reached, the run waits - logging every minute - instead of turning a dropped
connection into a pile of 'failed' items.  Stop still works while waiting.
"""
import socket
import time
from urllib.parse import urlparse


def _target():
    from engines import ledger_engine as eng
    u = urlparse(eng.LOGIN_URL)
    return u.hostname or "ura.go.ug", u.port or (443 if u.scheme == "https" else 80)


def reachable(timeout=6):
    host, port = _target()
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def wait_online(log, check, max_wait_s=30 * 60):
    """Return True when the portal can be reached (at once if it already can)."""
    if reachable():
        return True
    host, _ = _target()
    log(f"The connection to {host} is down - waiting for it to come back "
        f"(checking every 15 s, up to {max_wait_s // 60} min). Press Stop to end the run.", "warn")
    t0 = time.monotonic()
    next_note = 60
    while time.monotonic() - t0 < max_wait_s:
        for _ in range(15):
            check()
            time.sleep(1)
        if reachable():
            log(f"Connection is back after {int(time.monotonic() - t0)} s - continuing.", "ok")
            return True
        if time.monotonic() - t0 >= next_note:
            log(f"   still no connection ({int(time.monotonic() - t0) // 60} min)...", "warn")
            next_note += 60
    log("The connection did not come back - continuing; anything that fails is listed in the summary.",
        "err")
    return False
