import sys, os, shutil, glob; sys.path.insert(0,'/home/claude/ura_suite')
from playwright.sync_api import sync_playwright
from engines import returns_pw, returns_common
returns_common.TAB_WAIT_SECONDS = returns_common.PAGE_WAIT_SECONDS = 5   # BAD return must fail fast
out="/tmp/out_returns"; shutil.rmtree(out, ignore_errors=True)
logs=[]
def log(m): logs.append(m); print(m)
def run(skip):
    with sync_playwright() as p:
        b=p.chromium.launch(executable_path="/tmp/chr/chromium", headless=True, args=["--no-sandbox","--disable-gpu","--disable-dev-shm-usage"])
        pg=b.new_context().new_page(); pg.goto("http://localhost:8765/home.html")
        res={}
        returns_pw.run_on_page(pg,"1000012345","01/07/2025","03/10/2026",["Monthly VAT Return","PAYE Return","Local Excise Duty Return"],
                               log, lambda: None, out, patient=False, result=res, skip_existing=skip)
        b.close(); return res
r=run(True)
print("\n==== FILES"); [print("  ",os.path.relpath(f,out), os.path.getsize(f)) for f in sorted(glob.glob(out+"/**/*.*",recursive=True))]
print("==== ITEMS"); [print("  ",i) for i in r["items"]]
print("==== SECOND RUN (skip existing)")
logs.clear(); r2=run(True)
print("   skipped:", sum(1 for i in r2["items"] if i["status"]=="skipped"), "| saved:", sum(1 for i in r2["items"] if i["status"]=="saved"))
