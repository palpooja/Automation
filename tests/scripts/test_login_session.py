import sys, queue, threading, time; sys.path.insert(0,'/home/claude/ura_suite')
from engines import ledger_engine as eng
from core.runner import RunContext
from core import session
from pathlib import Path
eng.LOGIN_URL="http://localhost:8765/login.html"
def lb(p): return p.chromium.launch(executable_path="/tmp/chr/chromium", headless=True, args=["--no-sandbox","--disable-gpu","--disable-dev-shm-usage"])
eng.launch_browser=lb
ui, rq, stop = queue.Queue(), queue.Queue(), threading.Event()
ctx=RunContext(ui, rq, stop, Path("/tmp/out_s"), Path("/tmp/session_test.log"))
asks=[]; LOG=[]
def human():                       # stands in for the person at the yellow bar
    while True:
        k,p=ui.get()
        if k=="ask": asks.append(p[3]); time.sleep(0.5); rq.put("")
        if k=="log": LOG.append(p[0])
threading.Thread(target=human, daemon=True).start()
import sys as _s, traceback
try:
  with session.open_session(ctx, {"name":"ABC Traders Ltd","login":"1000012345","password":"secret"}) as page:
    R=("IN SESSION url=%s attempts=%s filled=%s" % (page.url, page.evaluate("window.__loginAttempts"), page.evaluate("window.__loginFilled")))
except BaseException as e:
  R="EXCEPTION: "+"".join(traceback.format_exception(e))[-1500:]
_s.__stdout__.write(R+"\n"); _s.__stdout__.write("asks shown for: %s\n" % asks); _s.__stdout__.write("\n".join(LOG[-25:])+"\n")
time.sleep(1)
