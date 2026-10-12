import sys; sys.path.insert(0,'/home/claude/ura_suite')
from playwright.sync_api import sync_playwright
from engines import wht_engine as we, wht_outputs as wo
we.RESULT_WAIT_MS=20000
with sync_playwright() as p:
    b=p.chromium.launch(executable_path="/tmp/chr/chromium", headless=True, args=["--no-sandbox","--disable-gpu","--disable-dev-shm-usage"])
    pg=b.new_context().new_page(); pg.goto("http://localhost:8765/Portal/wht_live.html")
    logs=[]
    bot=we.WhtBot(pg, logs.append, lambda: None)
    for code in ("D","W","X"):
        months=[wo.MonthResult(2026,m) for m in (1,2,3)]
        we.run_months(bot, months, want_pdf=False, want_html=False, generate_for=code)
        print(code, [(r.label, r.status, r.rows, (list(r.df.columns)[:4] if r.df is not None else None), r.message[:70]) for r in months])
    b.close()
print("\n".join(l for l in logs if "answered" in l or "STOPPED" in l or "Generate for" in l)[:900])
