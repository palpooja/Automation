"""Starts the real window inside the current (virtual) screen and prints measurements as JSON.
Used by test_screen_sizes.py through xvfb-run with different screen sizes and DPI.
argv: <data folder> <number of clients> [<screenshot path>]"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from core import store  # noqa: E402

data = Path(sys.argv[1])
for name, rel in (("DATA_DIR", ""), ("LOG_DIR", "logs"), ("CLIENTS_FILE", "clients.json"),
                  ("HISTORY_FILE", "history.json"), ("SETTINGS_FILE", "settings.json"),
                  ("OLD_RETURNS_CLIENTS", "none.json")):
    setattr(store, name, data / rel if rel else data)
n = int(sys.argv[2])
store.CLIENTS_FILE.write_text(json.dumps([{"name": f"Client {i:02d} Limited", "login": f"10000000{i:02d}",
                                           "password": store.protect_secret("x")} for i in range(1, n + 1)]))
import app as app_module  # noqa: E402

a = app_module.App()
for _ in range(10):
    a.update()
a.after(800)
a.update()


def box(w):
    return [w.winfo_rootx(), w.winfo_rooty(), w.winfo_width(), w.winfo_height(), bool(w.winfo_ismapped())]


out = {"screen": [a.winfo_screenwidth(), a.winfo_screenheight()], "dpi": a.winfo_fpixels("1i"),
       "window": [a.winfo_rootx(), a.winfo_rooty(), a.winfo_width(), a.winfo_height()],
       "state": a.state(), "minsize": list(a.minsize()),
       "run": box(a.run_sel_btn), "stop": box(a.stop_btn), "log": box(a.logbox),
       "queue_rows": len(a.queue_frame.winfo_children()) if a.queue_frame else 0,
       "scroll": getattr(a, "right_scroll", None) is not None,
       "scrollbar_shown": bool(a.right_scroll.bar.winfo_ismapped()) if getattr(a, "right_scroll", None) else None,
       "hscroll_shown": bool(a.right_scroll.hbar.winfo_ismapped()) if getattr(a, "right_scroll", None) else None,
       "content_width": a.right_scroll.inner.winfo_reqwidth() if getattr(a, "right_scroll", None) else None,
       "view_width": a.right_scroll.canvas.winfo_width() if getattr(a, "right_scroll", None) else None}
if len(sys.argv) > 3:
    import subprocess
    subprocess.run(["import", "-window", "root", sys.argv[3]], check=False)
print("PROBE" + json.dumps(out))
a.destroy()
