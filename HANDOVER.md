# URA Suite — Technical Handover

Status date: 2 October 2026. Three former programs (Returns Downloader 4.3, Tax Statement Downloader, WHT Extractor 3.0) merged into one Python/tkinter application on Playwright.

## 1. Layout
```
ura_suite/
  app.py                  window: vertical menu, pages, clients, run loop, question bar, summary writer
  core/
    store.py              paths, DPAPI client store, history, settings   (%LOCALAPPDATA%\URA_Suite)
    runner.py             Module contract, RunContext (log / ask / check / record), Stopped, PrintToLog
    session.py            open_session(): one browser + login per client, logout on exit
  modules/                one file per menu page; glue only, no browser code
    returns.py            -> engines/returns_pw.py
    ledger.py             -> engines/ledger_engine.py (run_job_on_page)
    wht_receivable.py     -> engines/wht_engine.py + wht_outputs.py
    settings_page.py      Settings page (not a run module)
    test_login.py         Clients > Test login
  engines/                all portal automation
    ledger_engine.py      = former ura_bot.pyw (3,090 lines) with: duplicate constants removed,
                          OUTPUT_ROOT/WORK_DIR/BROWSER_CHANNEL/PROXY/NAV_TIMEOUT_MS globals,
                          run_job split into _job_setup / run_job_on_page / run_job
    returns_pw.py         Playwright port of run_bot() from the Returns Downloader
    returns_common.py     Selenium-free parts of ura_returns_bot.pyw, extracted unchanged
                          (ALL_CATEGORIES, BRANCH_CODES, safe_name, unique_path, write_status_sheet)
    wht_engine.py         WhtBot + run_months, rebuilt from ura_bot.pyc bytecode
    wht_outputs.py        WHT cleaners + Excel/PDF/CSV writers, rebuilt from ura_outputs.pyc
    paye_engine.py        PAYE certificates (from the live recording)
    dialogs.py            one answer policy for portal confirm boxes (audit H7)
    sheet_safety.py       portal text stays text in Excel/CSV (audit L2)
  legacy/
    ura_returns_bot_v4_3_selenium.py   the old Returns program, for reference only - NOT imported
  tests/auto/             pytest suite (161 tests); tests/scripts/ = old print-only scripts
  BUILD_SUITE.bat         PyInstaller one-file build (no bundled Chromium)
  run.bat                 run from source
  requirements.txt        exact versions (see the file); requirements-build.txt adds pyinstaller;
                          requirements-dev.txt adds pytest
```

## 2. How a run works
1. `App.start()` reads each ticked module's options (`collect()`), runs `_preflight` (folder writable, free space), applies Settings to the Ledger engine (`_apply_engine_settings`), drops client+module pairs completed today if Resume is on, and starts ONE worker thread.
2. The worker, per client: `with open_session(ctx, client) as page:` → login (Ledger engine's `login()`, yellow bar via `ctx.ask`) → for each module `m.run(ctx, client, options, page)` → logout → browser closed. A module error is caught and recorded; the next module still runs. `Stopped` propagates and ends the run.
3. Modules report progress only through `ctx` (log/status/progress/record/ask). The worker never touches a widget; `App._pump()` drains `ctx.ui_queue` with `root.after(100)`.
4. At the end: `write_run_summary(rows, ctx.items, ...)` → `Run_Summary_<stamp>.xlsx` + append to `URA_Suite_Summary.xlsx`; history.json updated per job.

## 3. Contracts
**Module** (`core/runner.py`): `id`, `title`, `build_panel(parent) -> Frame`, `collect() -> dict` (raise `ValueError(msg)` for bad input), `run(ctx, client, options, page) -> dict` (small ints/strs for the Overview row; call `ctx.record()` for each file/failure).

**RunContext**: `log(text, kind='' | 'ok' | 'warn' | 'err')`, `status(text)`, `progress(i, n, text)`, `check()` (raises `Stopped`), `sleep(s)`, `ask(message, kind='login'|'dates'|'continue', defaults, poll)`, `set_scope(client, login, module)`, `record(item, status, detail='', file='')`, `items`, `output_root`, `log_path`, `login_timeout_s`.

**Engines and the Ledger engine's globals the suite sets**: `OUTPUT_ROOT`, `WORK_DIR`, `BROWSER_CHANNEL` ('' | 'chrome' | 'msedge'), `PROXY`, `NAV_TIMEOUT_MS`, `SPEED_UPS`, `DEBUG_DUMPS`, `PROMPT_HANDLER(message, kind, defaults, poll)`, `STOP_EVENT`.

## 4. Settings keys (`settings.json`)
`output_root`; `modules.<id>.ticked`; `suite.{browser, proxy, nav_timeout_s, login_timeout_s, skip_done_today, open_folder_when_done, summary_xlsx, min_free_mb}`.

## 5. Build and install
`BUILD_SUITE.bat` → `dist\URA_Suite.exe` (needs Python 3.11+, internet for pip). Hidden imports listed in the bat; add new engine files there. Installer (Inno Setup, per-user) not yet written — base it on the old `wht_installer.iss`, install to `{localappdata}\Programs\URA_Suite`.

## 6. Verification status
See TEST_REPORT.md for the full mock-portal test results (3 Oct 2026). Summary below is from before that.

| Verified in sandbox | Not verified (needs Windows + Chrome + portal) |
|---|---|
| Window, pages, question bar, settings persistence | Any portal interaction |
| Run loop with fake session: one login/client, module order, error containment, resume, pre-flight, summaries | Returns PDF via `Page.printToPDF` after "Print with Schedules" |
| `wht_outputs` cleaners + writers on a fake grid | WHT calendar walking and To-date sync on the live page |
| `wht_engine.run_months` with a mock bot | DPAPI (falls back to base64 outside Windows), PyInstaller build |

## 7. Known limits / next work
1. Installer (step 5).
2. Speed: 40 fixed `nap()` pauses in `ledger_engine.py`; pipelining the next client's login; both after a baseline run.
4. `RESULT_MATCH = "TIN"` in `wht_engine.py` is an assumption (original value not recoverable); the keyword fallback covers it.
5. `returns_pw` prints the schedule view: if the live portal re-renders in place slower than 1.5 s, raise the wait at the "rendered in place" branch.

## 8. Conventions
- All user-facing text spells out abbreviations once (TIN, DPAPI, CAPTCHA).
- No widget access outside the main thread. No `print()` in new code — use `ctx.log`; engine `print()` is redirected by `PrintToLog`.
- Selectors live at the top of each engine (`SEL` in wht_engine, constants in returns_pw, named constants in ledger_engine). Change them there, nowhere else.
- Every skip on timeout must be recorded with `ctx.record(..., 'failed', reason)`.

## 9. Behaviour added during testing (3 Oct 2026)
- `returns_pw`: MARK_STALE_JS / FRESH_JS — old results are tagged `data-bot-stale` before Search; only fresh rows are read.
  A return with no Print button within the limit is a failure (never printed half-loaded). 30-s heartbeat in long waits.
  Status sheet written also when stopped. `URA_DEBUG=1` prints step timings.
- `wht_engine`: same stale-marking; `RESULT_WAIT_MS` (60 s, or the Settings element wait if larger); a month with no
  new results and no "no records" message is FAILED; one retry per month.
- `ledger_engine.login_rejected()`: wrong security code detected immediately.
- Modules report items in `finally`, so Stop never hides work already done; WHT saves captured months on Stop.
- BUILD_SUITE.bat must keep `--collect-all playwright` (Playwright's driver).
- `tests/`: mock portal + scripts. They need a Chromium path; edit `executable_path` in the scripts.

## 10. Designs (3 Oct 2026)
`core/theme.py` holds three designs as data (A ledger, B revenue - default, C runboard); `fonts/` holds the bundled
open-licence typefaces, registered privately at start-up (`theme.register_fonts`). Settings > Design chooses one; it
applies at the next start. BUILD_SUITE.bat must keep `--add-data "fonts;fonts"`. Window colours come only from the
theme dict - add new colours there, never as hex in app.py or modules.

## 11. Web-style controls (3 Oct 2026)
`core/widgets.py`: WebButton (rounded, hover shade, presses down 3 px with the shadow collapsing, fires on release
only if the pointer is still on it, Space/Enter work, focus ring, disabled = grey), TextField (rounded, accent border
while typing), PasswordField (eye icon: open eye = show, slashed eye = hide, hover tip, click by position).
WebButton accepts tk.Button-style configure(text/state/command) and ignores colour keys - colours come from the theme.
All buttons in app.py, the module pages and the Clients dialogs use it. Tests: /tmp/t_widgets.py logic is described in
TEST_REPORT.md section 'Widgets'.

## 12. Between modules (3 Oct 2026)
app.py calls core.session.go_home(ctx, page, client, before=title) before every module except the first. It closes
extra tabs, then ledger_engine.ensure_portal_home (navigate the portal frame to PORTAL_HOME_URL; fallback via the public
site's Portal link; re-login only if the session ended). Never reload the outer page: it is the e-Tax login address and
reloading logs out (lesson recorded in the original Ledgers engine). On failure: home_failed_<TIN>.png + frame list.
WHT options: generate_for is now a list ["Payee","Payer"]; one run and one file set per type.

## 13. Live-portal fixes (3 Oct 2026, first real run)
ledger_engine: js_click(timeout=5 s), js_click_later (setTimeout click, returns at once), safe_click treats a vanished
element as a click that went through, export_table_to_excel (script click + expect_download 30 s, else build .xlsx from
the on-screen table with pandas), payment links by one script click, tick_cert_rows sets state by script and reads back.
wht_engine: IS_GRID_JS (TIN heading, >= 5 cols, not the form's fields), page-reload detection ('reloaded' -> 6 s grace ->
no records), GEN_WORDS / codes D and W (live), SetupError stops a type at once. app: Resume needs status == Completed
and files > 0 and asks. core/manual.py: manual browser + page capture zip. Note: Frame.evaluate takes NO timeout
argument; Locator.evaluate does.

## 14. Generate TCC (3 Oct 2026, from the live recording - tests/live_recording_2026-10-03.json)
wht_engine: TCC_TICK_JS (ticks rows with an empty TCC cell / #tccNo*), WhtBot.generate_missing_tcc (dialog handler
ACCEPTS - without one Playwright dismisses the confirm = cancel), waits for #TCCSummary, presses its Back, caller
searches the month again. CLEAN_HTML_JS copies text-box values into cells (TCC numbers live in inputs). A generation
error never fails the month (warning only). Option generate_tcc on the WHT page (default on). Ledgers page no longer
offers wht_tcc; ledger engine certificate step: first_wait 20 s, no manual hand-over, status "generate pressed".

## 15. PAYE certificates (3 Oct 2026)
engines/paye_engine.py: financial_years(from, to), PayeBot (open_page via menu link, search(fy) with stale marking,
grid(tick) on the INNERMOST table with 'Employee Tin', generate_missing (dialog accepted), month_wise(row) reading
#divPostContent in any frame and pressing its Close), save_workbook (Employees + Month-wise details), run(...).
modules/ledger.py: 'paye_tcc' is no longer passed to ledger_engine; it runs paye_engine after the ledger job (go_home
first when the ledger job ran). Options paye_generate / paye_monthwise; live financial-year preview under the tick box.

## 16. Version 10
Module ids/titles: returns "Returns", ledger "Ledgers", certificates "Certificates" (modules/wht_receivable.py, now WHT +
PAYE). Every module has `shared_period` set by app._period_changed (settings["period"]); Certificates also
`on_period_change`. app.ReviewPage renders (rows, items, summary path) sent as ("review", ...) at the end of a run.
core/net.py wait_online (host from ledger_engine.LOGIN_URL). Speed factor: settings suite.speed_factor -> ledger
SLOW_FACTOR / NAV_TIMEOUT_MS, wht RESULT_WAIT_MS, paye WAIT_S, returns_common TAB/PAGE_WAIT_SECONDS.
ledger_engine.open_level2 uses LEVEL1_READY_JS + TAX_LINK_JS. returns_pw.stable_label / already_saved.

## 17. Audit fixes and speed work (3 Oct 2026)
Full list in CHANGELOG.md. Technical points for the next maintainer:
- core/store.py: `try_unprotect` returns None when DPAPI cannot decrypt; such clients carry `password_unreadable` +
  `_blob` and `save_clients` writes the blob back unchanged. Damaged clients.json -> renamed clients.damaged-*.json,
  message in `store.LOAD_WARNINGS`. `.bak` before every save. `client_problem()` validates Add/Edit. `prune_logs(days)`
  at start-up (Settings `log_keep_days`, default 30). `add_history(entry, secrets)` masks passwords.
- core/runner.py: `needs_review(status)` is the single rule (fail/error/incomplete/check). `RunContext.add_secret`,
  `scrub`, `secrets`; log file writes under a lock.
- core/session.py: refuses an empty password; sets `ledger_engine.CREDS` / `CURRENT_TIN` for this client only and
  clears them in `finally`.
- ledger_engine: `login()` falls back to CREDS only for the same TIN. `fill_login_form` re-raises without
  Playwright's message (it repeats typed text). `RECHECK_VALUES` with a date stamp; `mark_progress` writes atomically.
  `level2_is_current` / `level2_usable` / `keep_previous_level2` (option `refresh_l2`, global `REFRESH_L2`).
  `NOTES` + `drain_notes` turn "built from the screen" and "Level 3 page did not mention the month" into `check`
  items. `work_file` / `safe_screenshot` replace relative paths and `os.chdir`.
- returns_pw: `existing_downloads` (PDF <= 1,000 bytes, empty .mhtml and .part files do not count); `_write_whole`
  (write .part then rename); empty PDFs deleted; `wait_for_print_buttons` (one pass for both buttons every 0.25 s,
  `SCHED_GRACE_S` = 3.5 s, never shorter than the old 2.5-3.5 s); `_text_size` before/after the schedules click.
- wht_engine: `_poll` replaces three fixed calendar pauses (600, 700, 600-1800 ms) with the same final check at the
  same moment, stopping early on success. `RELOAD_GRACE_S` (6 s x speed setting); `no_table_after_reload` note.
- core/leakcheck.py + Clients > "Check old files for passwords": scans the log folder, history.json and summary
  workbooks for any saved password; reports names only; can blank them (JSON re-validated before writing).
- Generate TCC options default off (`WhtReceivableModule.DEFAULT_GENERATE_*`). Unused password code removed from
  wht_outputs; unreachable PAYE branch removed from modules/ledger.py.

### Speed - measured on the mock portal (same results, see tests/auto/test_speed_equivalence.py)
| Test run | Before | After |
|---|---|---|
| WHT: months ok/empty/ok | 20.6 s | 14.7 s |
| WHT: live-style page, Payee + Payer + unknown | 54.2 s | 40.9 s |
| WHT: generate TCC | 32.9 s | 25.2 s |
| Returns: first run + resume run | 52.1 s | 47.7 s |
Every WHT month (status, rows, full table) and every Returns file and status was compared old vs new: identical; the
only differences are the intended wording changes (no-records note, "Generate pressed for N row(s)", schedules note).
Not done, and why:
- One browser for the whole run (new private session per client): measured 0.45 s saved per client here - too
  little for the new risk that one browser crash affects every later client.
- Remaining fixed pauses in the ledger Level 3 loop (ledger_engine `download_level3_for_month`, `click_back`) and
  PAYE month-wise: the mock portal cannot show they are unnecessary on the live portal. Needs one timed live run.
- Typing the next client's security code while the current client downloads (largest possible gain): needs a
  redesign - the engines share global state and the yellow bar is shared.

## 18. First live run of 10.1 (3 Oct 2026, 23:52) - findings and 10.2 changes
Where the time went (one client, 4 months, all three steps; total 6 min 50 s): login 35 s (25 s typing the
code); Returns 56 s (4 returns, about 6 s each, of which 4 s is waiting for a new tab that never opens);
Ledgers 160 s (about 10-12 s after each Level 3 month, of which 5 s is `dismiss_pre_login_popup` waiting for a
pop-up that never appeared; about 9 s back to Level 1 per tax type; 21 s to export an empty Local Excise Duty
Level 2); WHT 126 s (about 30 s per month: calendar 15-25 s, search 9-10 s); PAYE 20 s.
- Ledger Excel files (reported by the user after the run: "corrupted or extension error" for every portal
  export): engines/excel_files.py - `kind()` by magic bytes, `html_to_xlsx()` (lxml grid with colspan/rowspan,
  `cell_value()` rule: amounts -> numbers, identifiers/leading zeros/10+ digits -> text, read-back check),
  `save_portal_excel()` used by ledger_engine.save_download for .xlsx names, `repair_folder()` at app start and
  in `_job_setup`, `excel_ok()` + ledger_engine.month_file_done(): 'saved' in _progress.json no longer counts
  when the workbook is missing or not real. Not convertible -> renamed *.broken so it is fetched again.
- Logout question: engines/dialogs.py `expecting("logout")`; core/session.py `logout()` verifies the Logout
  button has gone (up to LOGOUT_WAIT_S) before reporting it. tests/mock_portal/home.html now asks the question.
- WHT calendar (10.2, REPLACED in 10.6 - see section 22): `step()` and day 1 used script clicks, which the live
  calendar ignores. tests/mock_portal/Portal/wht_slowclick.html is a home-made calendar that reacts to plain
  clicks - it could not show that fault; wht_dynarch.html uses the real calendar code.
- Open questions (need the person running the program): the 4 returns were almost certainly the 4 "to check"
  items (schedules not confirmed) - open one PDF: if it has the schedules, the check needs tuning (the new
  numbers-only log line shows what changed on the page); if not, the PDFs are missing schedules, which
  version 10 labelled "with schedules" without checking.
- Not changed without more live evidence: `dismiss_pre_login_popup` 5 s wait after each Level 3 month; the 4 s
  wait for a schedules tab; exporting an empty Level 2.

## 19. Combined Ledgers workbook (10.3)
engines/ledger_workbook.py `build()` - called by ledger_engine.build_combined_workbook() in the `finally` of
run_job_on_page (before write_summary), so it also runs after Stop or an error. Reads only files on disk
(`excel_ok`), statuses from this run's RESULTS then _progress.json. Header rows of a Level 3 file = leading rows
with no number. A failure is recorded as 'check', never fails the step. The mock portal cannot run the full
Ledgers step (no #btnSubmit page), so tests/auto/test_ledger_workbook.py drives run_job_on_page with the portal
steps replaced.

## 20. Returns schedules (10.4)
User, 4 Oct 2026: the return PDFs are the portal's HTML printed to PDF (by design: DevTools Page.printToPDF);
PAYE schedules are there; VAT schedules are not - **URA no longer gives VAT schedules**. Root cause of the old
timing: the init script (`install_print_catcher`) always recorded the portal's print() call in
window.__botPrint, but nothing waited for it; the PDF was made 4 s (new-tab wait) + 1.5 s after the click.
`press_print_with_schedules()` now: print request -> print; else quiet (no request in flight - `_Activity` -
and no change in visible text) for QUIET_S = 5.5 s -> print (VAT: not flagged; others: 'check'); else cap ->
'check'. tests/auto/test_schedules.py + tests/mock_portal/Portal/schedules_test.html?mode=fast|slow|newtab|never|busy;
the test server's /slow?ms= answers late. Next live run: the log shows per return how long it took; if VAT
returns take ~6 s each, the portal's VAT button does not call print (then plain Print would save ~5 s each -
a decision for the user, since it would stop catching schedules if URA brings them back).

## 21. Return PDF layout (10.5)
returns_pw `_page_layout()` runs before every `Page.printToPDF`: adds a print-only style hiding the portal's
print/back/close buttons, then measures `scrollWidth` under print media at the printable width (A4, 0.3-inch
margins: PORTRAIT_PX 736, LANDSCAPE_PX 1064). Fits -> (portrait, 1.0), unchanged; else landscape; else
scale = LANDSCAPE_PX / width (min 0.1). Viewport and media are restored afterwards.
tests/auto/test_return_pdf_quality.py + tests/mock_portal/Portal/wide_return.html.

## 22. Live runs of 5 Oct 2026 (10.6)
Evidence: logs run_20261005_*.txt, history.json and screenshots from the user's log folder; Swaraj's notes.
- WHT: the portal calendar is The DHTML Calendar (calendar.js `_add_evs`: mouseover/mousedown/mouseout only;
  `tableMouseUp` calls `cellClick`). wht_engine: `PRESS_JS` + `_press()`, `close_calendar()` (x button =
  data-bot 'close', else `CAL_HIDE_JS`), `_direct_set()` closes before writing, `pick_month()` stops after
  `DEAD_PRESSES`, logs months really moved, sets `expected_from`; `search()` closes the calendar, refuses a wrong
  From date, never force-clicks, `_from_on_same_page()` catches a From date changed by the click. CAL_SCAN skips
  week-number cells (class 'day wn'). Tests: tests/auto/test_wht_real_calendar.py (7).
- Ledgers: `find_level3_box()`, `wait_for_level1_patiently()` + `LEVEL1_PATIENCE_S`, recover_to_level1 uses
  `ensure_portal_home()`, `click_first_visible()` checks every match, the retry loop in run_job_on_page protects
  its recovery call, `_back_to_level1_after()` after a finished tax type. Tests: test_ledger_recovery.py (5).
- Returns: `wait_table_changed()` + `NEXT_PAGE_WAIT_S` / `NEXT_PAGE_WAIT_PATIENT_S`. Test: test_returns_paging.py.
- Equivalence: old 10.5 and new 10.6 engines compared on wht_live, wht_slowclick (D and W, 3 months each) and the
  Returns mock: identical statuses, rows, tables and files.

## 23. Pop-up rule, Level 3 timings, screen-copy check (10.7)
- Universal rule (see CLAUDE_CONTEXT.md): `dismiss_pre_login_popup` checks at once; `core.session.
  press_logout_link_if_shown` for Test Login. Guard test: tests/auto/test_popup_rule.py.
- `download_level3_for_month` keeps `TIMINGS` (tax type, month, open s, save s, back s; cleared in `_job_setup`)
  and prints a times line per month; `summarise_level3_checks()` runs in the `finally` of run_job_on_page.
  The yearly (annual) Level 3 routine is not timed.
- Option B evidence: `check_screen_copy()` tries the whole #levelThreeTwoOfAR box, then its last table, through
  `excel_files.html_to_xlsx` in a temp folder and `excel_files.compare_cells()` against the saved portal file;
  `_remember_screen_check()` writes `_screen_check.json`. Decision rule for switching Level 3 to the screen
  copy: every month identical on the live portal, for every tax type.
- Option C (each month's Level 3 without leaving Level 2) waits for the user's page capture of Level 2 and
  Level 3 ('Open portal (manual)' > 'Save this page for Claude').

## 24. Option C and the manual-mode freeze (10.8)
- `level3_month()` chooses the route; `download_level3_hidden()` (HIDDEN_SUBMIT_JS: fresh hidden iframe each
  month, form target set, link.onclick() called, target/action restored) + `_wait_hidden_level3()` (an answer
  without the Level 3 box = error page, ends the wait at once) + `_save_level3()` shared with the click route.
  `refresh_level2()` = recover_to_level1 + open_level2 before falling back. ROUTE_STATE per tax type, cleared
  in `_job_setup`. TIMINGS entries carry the route ('click' / 'background'); proof copies are not timed.
- `excel_files.compare_cells(..., ignore_text=TIME_LINE_TEXT)`.
- Manual mode: core/manual.py unroutes the filter and polls with `keep_serviced`; core/session.open_session sets
  and clears `eng.BROWSER_KEEPALIVE`; `ask_user` uses it when no poll is given.
- Not tested on the live portal yet: whether the portal accepts several Level 3 requests from one Level 2
  (security token `_csrf`, `lastEventTS`). The first live run's `_route_check.json` and log answer it.

## 25. One workbook per tax head (10.9)
- engines/ledger_taxhead.py: `months_folder()`, `move_old_month_files()`, `read_level3()` (heading row = the row with
  Sr No./Payment Basis; 'Tax Position as at' and 'Period:' lines above it), `update_tax_head()` (merge: monthly
  files win, other months kept from the existing workbook; atomic replace).
- ledger_engine.process_tax_type saves monthly files in `month_dir`; `update_tax_head_files()` runs in the
  `finally` of run_job_on_page (replaces build_combined_workbook, still defined, no longer called).
- `summarise_level3_checks` reads every `_screen_check.json` under the client folder.
- Not yet confirmed on a real file: a Level 3 with a totals line (if the portal adds one, it is copied as a row).

## 26. Voice, screen sizes, Excel import (10.10)
- core/voice.py: phrase(), script() (PowerShell System.Speech, same voice list and order as Voice_Samples.ps1),
  say() (Popen, CREATE_NO_WINDOW, never raises). Called in app.py after each client's open_session has closed
  and after Stop. Setting suite.voice: "1" | "4" | "off".
- core/display.py: make_dpi_aware() before Tk(); place() for the main window and dialogs; ScrollArea (vertical
  only when needed, horizontal only when content is wider - about 1024 px screens). App._reflow() stacks the run
  buttons under the client list on narrow windows. QUEUE_ROWS = 12. WebButton padding scales with DPI.
- core/client_import.py: make_template(), read_clients() -> Result(new, updates, problems, unchanged), apply(),
  summary(). Uses store.client_problem; passwords only in the client records.
- tests/auto/window_probe.py starts the real App in an xvfb screen and prints measurements (screenshots too).

## 27. Clean stop when Chrome is closed by hand (10.11)
- core/runner.py: `BrowserClosed(Stopped)`, `browser_gone(page=None, exc=None)` (TargetClosedError, also wrapped,
  or page.is_closed()). App.start/work(): checks before each next module and after a client's last module; a
  module error with the browser gone -> row "Stopped: browser closed", `closed_note()`, raise BrowserClosed;
  outer handler "--- Stopped: the browser window was closed. ---" + voice "Job stopped.".
- Voice counter: `for ci, (c, mods_c) in enumerate(by_client)`; more = ci < len(by_client) - 1.
- Test helper `_run_with_fakes` in tests/auto/test_window.py drives the real run loop; it drains the window's log
  timer after the worker ends.


## 28. Ledger workbook like WHT, no file per month (10.12)
- engines/ledger_taxhead.py: `update_tax_head(save_folder, tax, log, statuses, keep_month_files)` writes Summary,
  Level 3-2 and a tab per month (`_copy_sheets` copies values, bold, number formats, merges, widths;
  `_rebuilt_tab` for a month known only from the table); then `remove_folded()` re-opens the saved workbook and
  removes only the monthly files whose tab is in it. `months_in_workbook()` (cached by file time) feeds
  `ledger_engine.month_file_done`.
- ledger_engine: `KEEP_MONTH_FILES` (from settings `keep_month_files` in `_job_setup`); `update_tax_head_files`
  passes this run's month statuses from RESULTS; month records link to the workbook; `_route_check_path()` /
  `move_route_check_up()` keep `_route_check.json` in the client folder.
- Yearly tax heads (Income Tax, Rental) unchanged: their Level 3 items stay separate files.

## 29. PAYE certificates from the live capture (10.12)
- paye_engine.FRESH_JS: 'none'/'reloaded' only in a document holding #TaxCreditCertForm/#period. search() looks
  at every frame each poll; RELOAD_GRACE_S = 10 after the PAYE page came back; sets `no_list_reason`,
  `evidence_frame`; `save_evidence(stem)` -> debug_dir html + png. run() records the reason.
- month_wise(): `_find_monthwise_impl` looks in new context pages (pop-up) first, then the portal window;
  pop-ups are closed with page.close(). MONTHWISE_WAIT_S = 30.
- generate_missing(fy): batches via GRID_JS {tick, limit} (unticks first so the page's sCount stays right;
  skips allowTCC 'N'); LIMIT_JS reads the page's `count` (5); re-searches after each batch; stops on no progress.
- wht_engine.FRESH_JS 'reloaded' requires #frmDt/#generateFor in that document.
- Still unknown: what the live page shows when a year truly has no employees (message or empty page) - the
  saved page copy will show it.

## 30. Audit-file names and checks (10.13)
- returns_common: `return_facts`, `friendly_label`, `fy_label`, `unique_label`, `load/save_saved_as`
  (`_saved_as.json` = {name: return number}), `BOT_FILE` 'Saved as' column. returns_pw: `choose_name()` (friendly
  name, resume by return number, renames 10.12 names), `no_records_screenshot()`. Mock: history.html WTHDRT rows
  laid out like the live portal.
- ledger_taxhead: Summary columns Level 3 principal / Difference / Check; `_check()`; `CHECK_COLUMN`;
  `update_tax_head(..., checks=[])` -> ledger_engine records 'check' items.
- engines/wht_crosscheck.py + WhtReceivableModule._check_vs_ledger (Payee only).

## 31. Slow returns set aside (10.14)
- returns_pw.download_one: one budget (first_try_s) from View to the Print button on the first try; `last['why']`
  = 'slow' | 'skipped'; `set_aside()` -> result['deferred']; retry rounds also set aside slow ones.
- `person` (RunContext): attention() says core.voice.attention_phrase() and shows bar kind 'skipwait' without
  blocking; poll_answer() 's' = Skip, 'w' = Keep waiting; hide_question().
- ctx.after_client: [(title, fn(page))], reset per client; app.py runs them after the client's modules (go_home
  first), errors recorded, Stop/BrowserClosed honoured.

## 32. Ledger check, yearly ledgers, PAYE search, second try (10.15)
- ledger_taxhead: `canonical()` / `is_period()` accept a month (MAY-2026) or a year ('JUL-2025 to JUN-2026');
  `level2_payments()` reads the Payment sub-columns of `<tax>_L2.xlsx` (two heading rows); `level3_figures()`
  (PRN payments per distinct reference, credits by Payment Basis, not allocated); `_check()` compares the totals.
  Run notes 'Payment X; Tax credit Y' are the fallback when the Level 2 file cannot be read.
- ledger_engine yearly: `ANNUAL_LINKS_JS` marks one `openLevelThreeTwo` link per year (`data-bot-annual`), skips '0';
  `download_level3_year()` click -> `find_level3_box` -> `_save_level3` -> Back; `process_annual_tax_type()` saves to
  `_months\<tax>\<tax>_L3_<year>.xlsx`, a failure per year (`save_failure_copy`); `update_tax_head_files` includes
  ANNUAL_TAX_TYPES. `find_level3_box` stops after ERROR_PAGE_GRACE_S (8 s) on a loadLevelThreeTwo page without the box.
- wht_crosscheck: `ledger_lines()` from Income Tax.xlsx 'Level 3-2', unfolded _months files and legacy files;
  `crosscheck()` returns (certs, left, files, others).
- paye_engine: `SearchTimeout`; request listener (searchPayeEntry) -> press again after PRESS_AGAIN_S (8 s) with
  frmSubmitFlag cleared; `_search_twice()`; `run(reopen=, failed=, years=)`.
- Second try: RunContext.final_try (record() updates an existing item of the client; QUIET_STATUSES keep the first
  try's line; `resolve()`), app.py runs ctx.after_client with final_try set and adds a whole-step second try when a
  module raises; modules register their own second try only when not final_try.
- Test portals: ledger_server kind='yearly' (fail_years, other_requests); paye_server drop_presses / hang.

## 33. Level 3 by the direct way (10.16)
- ledger_engine: L3_ROUTE ('auto' / 'direct' / 'screen'), DIRECT_SWITCH_AFTER_S, DIRECT_HEARTBEAT_S, DIRECT_SAFETY_S,
  DIRECT_STATE per tax type {on, verified}; set in _job_setup from settings 'l3_route' / 'direct_switch_s'.
- DIRECT_START_JS: fetch POST from the Level 2 frame, body = FormData of form 'levelTwoOfAR', URL built like
  openLevelThreeTwo; DIRECT_POLL_JS parses the answer with DOMParser and returns #tblData.outerHTML (else the table
  with 'Sr No.'); DIRECT_ABORT_JS. download_level3_direct() saves through excel_files.save_portal_excel (original in
  portal_original), TIMINGS route 'direct'.
- _with_direct(page, frame, folder, tax, label, find_link, screen_way): used by level3_month (monthly; the old body
  is level3_month_screen) and level3_year (yearly). compare_level3(): identical / SAME_LINES / different - also used
  by the background way's proof.
- Live facts: the security code (_csrf) is the same for the whole session (two answers 6 min apart, same code); a
  request from the top frame has no form -> 403 (that was the user's first console try).

