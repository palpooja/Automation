import sys, os, subprocess, time; sys.path.insert(0,'/home/claude/ura_suite'); os.chdir('/home/claude/ura_suite')
from core import store
st=store.load_settings(); st.setdefault("suite",{}).update({"theme":"revenue","open_folder_when_done":False,"skip_done_today":False,"nav_timeout_s":15}); store.save_settings(st)
from engines import ledger_engine as eng
eng.LOGIN_URL="http://localhost:8765/login.html"
eng.PORTAL_HOME_URL="http://localhost:8765/Portal/loginController/loadHomePageContent/"
eng.launch_browser=lambda p: p.chromium.launch(executable_path="/tmp/chr/chromium", headless=True, args=["--no-sandbox","--disable-gpu","--disable-dev-shm-usage"])
import app
a=app.App(); a.geometry("1240x820+0+0")
a.clients=[{"name":"Global 7","login":"1000626970","password":"pw"},{"name":"ABC Traders Ltd","login":"1000012345","password":"x"}]
a.ticked_logins={"1000626970"}; a._refresh_clients()
for v in a.menu_vars.values(): v.set(True)
rm,lm,cm=app.MODULES
[v.set(c=="Monthly VAT Return") for c,v in rm.cat_vars.items()]; rm.patient.set(False)
import engines.returns_common as rc
a.period_from.set("2026-06"); a.period_to.set("2026-07")
a.update(); a.show(rm); a.update(); subprocess.run(["import","-display",":99","-window","root","/tmp/new_returns.png"])
a.show(cm); a.update(); subprocess.run(["import","-display",":99","-window","root","/tmp/new_certs.png"])
a.show(lm); a.update(); subprocess.run(["import","-display",":99","-window","root","/tmp/new_ledgers.png"])
def person():
    if a.ask.winfo_ismapped(): a.after(300, lambda: a._answer(""))
    if a.worker and not a.worker.is_alive():
        a.after(800, lambda: (a.update(), subprocess.run(["import","-display",":99","-window","root","/tmp/new_review.png"]), a.destroy())); return
    a.after(600, person)
orig=a._apply_engine_settings
def patched(su):
    orig(su); rc.TAB_WAIT_SECONDS = rc.PAGE_WAIT_SECONDS = 5
a._apply_engine_settings=patched
a.start(all_ticked=True); a.after(1500, person); a.mainloop()
log=open(sorted(store.LOG_DIR.glob("run_*.txt"))[-1]).read()
for l in log.splitlines():
    if any(k in l for k in ("=====","-----","Downloading only","saved ->","could not be","failed for","Portal home page is open","Generate for","PAYE certificates:","employee(s)","Saved: ","All done","Run summary","WARNING")): print(l[11:150])
