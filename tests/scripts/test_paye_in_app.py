import sys, os, glob; sys.path.insert(0,'/home/claude/ura_suite'); os.chdir('/home/claude/ura_suite')
from engines import ledger_engine as eng
eng.LOGIN_URL="http://localhost:8765/login.html"
eng.PORTAL_HOME_URL="http://localhost:8765/Portal/loginController/loadHomePageContent/"
eng.launch_browser=lambda p: p.chromium.launch(executable_path="/tmp/chr/chromium", headless=True, args=["--no-sandbox","--disable-gpu","--disable-dev-shm-usage"])
import app, subprocess
from core import store
a=app.App(); a.geometry("1240x820+0+0")
a.clients=[{"name":"Global 7","login":"1000626970","password":"pw"}]; a.ticked_logins={"1000626970"}; a._refresh_clients()
for k,v in a.menu_vars.items(): v.set(k=="ledger")
lm=app.MODULES[1]; [v.set(False) for v in lm.tax_vars.values()]; [v.set(k=="paye_tcc") for k,v in lm.other_vars.items()]
lm.from_var.set("2025-05"); lm.to_var.set("2026-06")
a.settings["suite"]={"open_folder_when_done":False,"skip_done_today":False,"nav_timeout_s":15}
a.show(lm); a.update(); subprocess.run(["import","-display",":99","-window","root","/tmp/ledger_paye_page.png"])
def person():
    if a.ask.winfo_ismapped(): a.after(300, lambda: a._answer(""))
    if a.worker and not a.worker.is_alive(): a.after(300, a.destroy); return
    a.after(600, person)
a.start(all_ticked=True); a.after(1500, person); a.mainloop()
log=open(sorted(store.LOG_DIR.glob("run_*.txt"))[-1]).read()
for l in log.splitlines():
    if any(k in l for k in ("PAYE","Period:","employee","Month-wise","Saved:","failed","All done","Run summary")): print(l[11:160])
