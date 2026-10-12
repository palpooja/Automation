import sys, time; sys.path.insert(0,'/home/claude/ura_suite')
from playwright.sync_api import sync_playwright
from engines import ledger_engine as eng
with sync_playwright() as p:
    b=p.chromium.launch(executable_path="/tmp/chr/chromium", headless=True, args=["--no-sandbox","--disable-gpu","--disable-dev-shm-usage"])
    pg=b.new_context().new_page(); pg.goto("http://localhost:8765/Portal/level1_slow.html")
    t=time.time()
    try:
        eng.SLOW_FACTOR=1.0
        orig=eng.find_frame_containing_text.__wrapped__ if hasattr(eng.find_frame_containing_text,'__wrapped__') else None
        # the OLD step: exact-text search for 20 s
        f=None
        deadline=time.time()+20
        while time.time()<deadline and f is None:
            for fr in pg.frames:
                if fr.get_by_text("IT-WHT", exact=True).count(): f=fr
            pg.wait_for_timeout(500)
        print(f"OLD search after {time.time()-t:.0f} s:", "found" if f else "NOT found -> 'Could not find text IT-WHT' (your log)")
    except Exception as e: print("old:", e)
    pg.goto("http://localhost:8765/Portal/level1_slow.html"); t=time.time()
    l2=eng.open_level2(pg, "IT-WHT")
    print(f"NEW open_level2 after {time.time()-t:.0f} s: Level 2 frame = {l2 is not None}, opened = {pg.evaluate('window.opened')}")
    b.close()
