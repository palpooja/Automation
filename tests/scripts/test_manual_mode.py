import sys, os, zipfile, glob; sys.path.insert(0,'/home/claude/ura_suite'); os.chdir('/home/claude/ura_suite')
from engines import ledger_engine as eng
eng.LOGIN_URL="http://localhost:8765/login.html"
eng.launch_browser=lambda p: p.chromium.launch(executable_path="/tmp/chr/chromium", headless=True, args=["--no-sandbox","--disable-gpu","--disable-dev-shm-usage"])
import app
from core import store
a=app.App(); a.clients=[{"name":"Global 7","login":"1000626970","password":"pw"}]; a.ticked_logins={"1000626970"}; a._refresh_clients()
a.open_manual=a.open_manual  # use as is, but skip the confirmation box:
app.messagebox.askokcancel=lambda *x, **k: True
step={"n":0}
def person():
    if a.ask.winfo_ismapped():
        txt=" ".join(w.cget("text") for w in a.ask_row.winfo_children() if hasattr(w,"_text"))
        if "Save this page" in [w._text for w in a.ask_row.winfo_children() if hasattr(w,"_text")][0]:
            step["n"]+=1
            if step["n"]==1: a._answer("capture")
            elif step["n"]==2: a.after(10, lambda: a._answer("capture"))
            else: a._answer("q")
        else:
            a._answer("")          # the login prompt: 'Log In'
    if a.worker and not a.worker.is_alive() and step["n"]>=3: a.after(300,a.destroy); return
    a.after(700, person)
a.after(500, a.open_manual); a.after(1500, person); a.mainloop()
log=open(sorted(store.LOG_DIR.glob("manual_*.txt"))[-1]).read()
print("\n".join(l[11:170] for l in log.splitlines() if any(k in l for k in ("browser is yours","Saved screen","Page captures","Logged","closed")) ))
z=sorted(store.LOG_DIR.glob("page_captures_*.zip"))[-1]
print("ZIP:", z.name, [n for n in zipfile.ZipFile(z).namelist()])
print(zipfile.ZipFile(z).read([n for n in zipfile.ZipFile(z).namelist() if n.endswith("frames.txt")][0]).decode())
