import sys, os, shutil; sys.path.insert(0,'/home/claude/ura_suite')
from pathlib import Path
from playwright.sync_api import sync_playwright
from engines import ledger_engine as eng
out=Path("/tmp/out_ledger_bits"); shutil.rmtree(out, ignore_errors=True)
with sync_playwright() as p:
    b=p.chromium.launch(executable_path="/tmp/chr/chromium", headless=True, args=["--no-sandbox","--disable-gpu","--disable-dev-shm-usage"])
    pg=b.new_context(accept_downloads=True, viewport={"width":1400,"height":900}).new_page()
    pg.goto("http://localhost:8765/Portal/ledger_bits.html"); f=pg.main_frame
    print("export that works   :", eng.export_table_to_excel(pg, f.locator("body"), out, "A_L2.xlsx", table_locator=f.locator("#tblData"), wait_s=10))
    print("export that is stuck:", eng.export_table_to_excel(pg, f.locator("#dead"), out, "B_L2.xlsx", table_locator=f.locator("#tblData"), wait_s=3))
    import pandas as pd; print("   built file rows:", pd.read_excel(out/"B_L2.xlsx").values.tolist())
    n=eng.tick_cert_rows(pg, f, [{"n":0},{"n":2}])
    print("ticked:", n, "| states:", f.evaluate("()=>[...document.querySelectorAll('[data-bot-cert]')].map(e=>e.checked)"), "| box0 clicks:", f.evaluate("()=>window.c0"))
    eng.safe_click(pg, f.locator("#far"), "far link")
    print("off-screen link clicks:", f.evaluate("()=>window.farClicks"))
    b.close()
