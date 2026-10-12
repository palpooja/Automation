import sys, os, time, glob; sys.path.insert(0,'/home/claude/ura_suite'); os.chdir('/home/claude/ura_suite')
MODE=sys.argv[1]   # "nofix" or "fix"
from engines import ledger_engine as eng, returns_common
eng.LOGIN_URL="http://localhost:8765/login.html"
eng.PORTAL_HOME_URL="http://localhost:8765/Portal/loginController/loadHomePageContent/"
eng.launch_browser=lambda p: p.chromium.launch(executable_path="/tmp/chr/chromium", headless=True, args=["--no-sandbox","--disable-gpu","--disable-dev-shm-usage"])
returns_common.TAB_WAIT_SECONDS = returns_common.PAGE_WAIT_SECONDS = 5
_orig=eng.open_tax_statement
def ots(page):
    _orig(page)
    for f in page.frames:
        if f.locator("#tsTitle").count(): raise RuntimeError("TEST-STOP: Ledgers reached the Tax Statement screen")
    raise RuntimeError("TEST: Tax Statement screen not reached")
eng.open_tax_statement=ots
import core.session
if MODE=="nofix": core.session.go_home=lambda ctx,page,c,before="": None
import app
from core import store
a=app.App()
a.clients=[{"name":"ABC Traders Ltd","login":"1000012345","password":"pw1"}]; a.ticked_logins={"1000012345"}; a._refresh_clients()
for v in a.menu_vars.values(): v.set(True)
rm,lm,wm=app.MODULES
[v.set(c=="Monthly VAT Return") for c,v in rm.cat_vars.items()]; rm.patient.set(False)
[v.set(t=="Value Added Tax") for t,v in lm.tax_vars.items()]
wm.from_var.set("2026-06"); wm.to_var.set("2026-07"); wm.gen_vars["Payee"].set(True); wm.gen_vars["Payer"].set(True)
a.settings["suite"]={"open_folder_when_done":False,"skip_done_today":False,"nav_timeout_s":15}
def person():
    if a.ask.winfo_ismapped(): a.after(300, lambda: a._answer(""))
    if a.worker and not a.worker.is_alive(): a.after(300, a.destroy); return
    a.after(600, person)
a.start(all_ticked=True); a.after(1500, person); a.mainloop()
log=open(sorted(store.LOG_DIR.glob("run_*.txt"))[-1]).read()
keys=("-----","Returning to the portal home","Portal home page is open","Could not get back","TEST","failed for","Payee June","Payee July","Payer June","Payer July","Generate for:","Saved:","All done")
for l in log.splitlines():
    if any(k in l for k in keys): print(l[11:150])
