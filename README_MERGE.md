# URA Suite - merge status (steps 1-4 done; installer pending)

One window, vertical menu on the left, three modules:

| Menu item | Code | State |
|---|---|---|
| Returns | `modules/returns.py` -> `engines/returns_pw.py` (Playwright port of ura_returns_bot.pyw 4.3; helpers in `returns_common.py`; old Selenium file moved to `legacy/`) | ported, **needs first live run** |
| Ledgers & Certificates | `modules/ledger.py` -> `engines/ledger_engine.py` (= ura_bot.pyw; `run_job` split into `run_job_on_page`) | Playwright, code path unchanged |
| WHT Receivable | `modules/wht_receivable.py` -> `engines/wht_engine.py` + `engines/wht_outputs.py` (rebuilt from bytecode) | rebuilt, **needs first live run** |

Shared pieces (`core/`): `store.py` (one DPAPI-encrypted `clients.json`, history, settings, output root)
and `runner.py` (module contract, thread-safe log/question plumbing).

## Where things are now
- Clients, history, settings, logs:  `%LOCALAPPDATA%\URA_Suite\`
  (an existing `%APPDATA%\URA_Returns_Downloader\clients.json` is copied in on first start)
- Downloads: `<output folder>\Returns\<TIN>\...`, `<output folder>\Ledgers\<TIN>\...`
  (default `Downloads\URA_Suite`; change with the "Output folder" menu item)

## Edits made to the two original engines
`engines/ledger_engine.py` (from ura_bot.pyw):
1. duplicated constants block (old lines 1404-1417) removed
2. `OUTPUT_ROOT` / `WORK_DIR` globals added; `run_job` saves under `OUTPUT_ROOT` and no longer
   `chdir`s to the program folder (so it works when installed under Program Files)
3. `python-dotenv` import made optional
`engines/returns_engine.py`: none. Its own `main()` window is not called; `run_bot()` is used directly
with `BASE_FOLDER` / `TEMP_PRINT_FOLDER` pointed at the suite's output folder.

## Known limits of this step
- Each module logs in to the portal itself, so a client with 2 ticked modules types the CAPTCHA twice.
  One login per client needs all modules on the same browser session = the Playwright port of Returns (step 5).
- The Returns engine has no Stop button of its own; Stop takes effect between modules/clients.
- `engines/wht_outputs.py` is a reconstruction from bytecode (see its docstring for what is exact and
  what is re-implemented). Test with one real month.

## Next steps
2. Rebuild the WHT Receivable browser engine (`Bot` class + `run_months`) from the bytecode listing.
3. Port Returns to Playwright (print-to-PDF, .mhtml, Headquarter/Both) -> one browser engine.
4. Single login per client shared by all modules; drop Selenium.
5. Inno Setup installer (per-user), one MAKE_INSTALLER.bat.

## Added in step 1b
- **Settings page** (menu): browser (auto/chrome/msedge), proxy, page-element wait, security-code wait,
  resume, open-folder-when-done, run summary, minimum free disk space. Saved to settings.json ("suite").
  Applied to the Ledger engine via `BROWSER_CHANNEL`, `PROXY`, `NAV_TIMEOUT_MS` (new globals in
  `engines/ledger_engine.py`). The Returns engine (Selenium) does not read them until its Playwright port.
- **Resume**: with "Resume" ticked, a client + module pair recorded as *Completed* today in history.json is
  skipped and listed in amber in the log. Tested with a fake module.
- **Pre-flight**: before a run the output folder is write-tested and free space checked (`App._preflight`).
- **Run_Summary_<stamp>.xlsx** in the output folder: client x module x status x files x seconds x note.
- **Yellow bar** now shows which client the question is about; the security-code wait times out after the
  Settings value instead of waiting for ever (`RunContext.ask`).
- **Clients > Test login**: opens the portal with the Ledger engine's login(), waits for the code, reports
  whether the portal menus appeared, logs out. Downloads nothing. (`modules/test_login.py`; needs Playwright,
  so it is untested here.)

## Speed - what costs time today and what to do about it
Measured from the code, not from a run (no run is possible in the sandbox):

| # | Where the time goes | Evidence | Change | Expected gain |
|---|---|---|---|---|
| 1 | Browser start + login once **per module** | each engine opens its own browser: `ledger_engine.run_job`, `returns_engine.run_bot` | one browser session per client shared by all modules (needs the Returns Playwright port) | ~30-60 s and one CAPTCHA per extra module per client |
| 2 | Fixed pauses | 40 `nap(page, N)` calls in the Ledger engine (8x1 s, 6x4 s, 6x1.5 s, 5x1.2 s, 3x3 s...), 10 `time.sleep` in Returns (one 5 s, two 2 s) | replace each with a wait for the element it is waiting for (`wait_find` already exists, Ledger engine line ~70) | up to ~45 s per ledger run, ~10 s per return |
| 3 | Pages load chat widgets / analytics / 3rd-party images | `speed_up()` in Ledger engine (route blocking, portal's own files never blocked) | already on by default in the suite's Ledger page; keep it on | depends on the connection; largest on slow links |
| 4 | Re-downloading files that exist | Ledger engine skips existing Level 2/3 files (lines ~951-1015); Returns renames to `_2`, `_3` instead of skipping (`returns_engine` line ~381) | Returns: skip when the return-number file already exists, unless its status changed | whole re-runs become minutes |
| 5 | Program start-up | a `--collect-all playwright` build unpacks ~1 GB to Temp on every start (`ura_bot_fast.py` docstring, point 3) | `BUILD_SUITE.bat` does not bundle Chromium: uses the installed Chrome/Edge | start in seconds instead of ~1 min |
| 6 | Waiting for you between clients | the run blocks at the yellow bar for each client | **pipeline**: open client N+1's login while client N is still downloading (second browser context); CAPTCHA typing then overlaps with downloads | for 10+ clients, run time approaches "typing time + longest client" instead of the sum. Portal rate limits are unknown - verify with 2 clients first |
| 7 | Level 3 detail for nil months | optional tick in the Ledger page (`skip_nil`) | leave ticked when you only need Level 2 | one page load per nil month saved |

## Steps 2-4 (this version)
- `engines/wht_engine.py`: WhtBot + run_months rebuilt from ura_bot.pyc. Month loop unit-tested with a mock bot
  (data / empty / error months, warnings, page copies).
- `engines/returns_pw.py`: Returns on Playwright. PDF via DevTools `Page.printToPDF` after clicking
  'Print with Schedules' (window.print disabled); .mhtml via `Page.captureSnapshot`; Stop honoured;
  **returns already on disk are skipped** (tick box on the page). No more Chrome-driver, no kiosk printing,
  no temp print folder. Selenium removed from requirements.
- `core/session.py`: one browser + one login per client; every ticked module runs on that page; logout.
  Verified with the real app loop and a fake session (2 clients x 3 modules, one simulated error).
- `engines/ledger_engine.py`: `run_job_on_page(page, settings)` added; `run_job` is now a thin wrapper.

## First live run - what to watch (cannot be tested outside Windows + Chrome + the portal)
1. Returns: does `Page.printToPDF` give the schedule view after 'Print with Schedules'? If the button
   opens a NEW window the port follows it; if it re-renders in place the port prints the same tab after 1.5 s.
   If the PDF lacks schedules, send the log and one .pdf.
2. WHT Receivable: the calendar heading format and the `#toDt` auto-fill (`sync_to_date`) - the log says
   "To date confirmed" when right.
3. Logout at the end of each client (core/session.logout) - the log says "Logged out".

## Still to do
5. Inno Setup installer (per-user) + MAKE_INSTALLER.bat.
6. Speed item 2 (replace the 40 fixed `nap()` pauses in the Ledger engine) and item 6 (pipeline the next
   client's login) - both after the first live run, so timing changes are measured against a baseline.

## Fix 3 Oct 2026
Build failed with `No module named 'selenium'`: `returns_engine.py` imported Selenium at the top and was still imported for its lists. Its Selenium-free parts are now `engines/returns_common.py`; the old file is in `legacy/` and not imported. Tested in a clean virtual environment with only requirements.txt.
