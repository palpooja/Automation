# CLAUDE_CONTEXT.md — paste this (plus only the file you want changed) when asking Claude for a change

You are helping maintain **URA Suite**, a Windows desktop program (Python 3.11+, tkinter, Playwright) that downloads documents from the URA e-Tax portal (ura.go.ug) for many clients. A human types the CAPTCHA at each login; never automate that.

## Rules for any change
1. Return the **complete replacement file**, not a diff, unless I ask for a patch.
2. Keep the user-facing style: plain English, abbreviations spelled out once (TIN = Tax Identification Number, DPAPI = Data Protection API, CAPTCHA = the picture security code).
3. Never call a tkinter widget from the worker thread; use `ctx.log/status/progress/ask` (queue → main thread).
4. Every timeout that skips an item must be logged and recorded via `ctx.record(item, 'failed', reason)`.
5. Selectors go in the engine's selector constants, nowhere else.
6. Do not add Selenium, bundled Chromium, plain-text password storage, or anything that bypasses the CAPTCHA.
7. Run the automated tests after every change: `python -m pytest` (see tests/auto). A change that turns a test red
   is not finished. A speed change must leave every result identical (tests/auto/test_speed_equivalence.py).
8. Passwords: never put one in a message, a file name or an exception. `ctx.add_secret()` masks known ones in every
   log line, summary row and history entry; `fill_login_form` re-raises errors without Playwright's text.
9. Portal questions (confirm boxes) are answered only through engines/dialogs.py: OK for the generate-TCC question
   inside `with bot.expecting_generate_confirm():`, Cancel for anything else.
10. Never mark something done unless it was saved: `already_done` retries FAILURE_VALUES always and RECHECK_VALUES
    ('not in ledger', 'no payment link') on a later day; Returns ignores PDFs of 1,000 bytes or less.
11. Portal downloads are judged by their bytes, not their names (engines/excel_files.py): the portal's
    Export To Excel is an HTML table and must go through save_portal_excel().
12. Spreadsheets: write portal text through engines/sheet_safety.py (as_text_cell / neutralise_sheet / csv_text).
13. Outcomes a person must look at use a status containing fail / error / incomplete / check (core.runner.needs_review).

## Files (what to send Claude for which change)
| Want to change… | Send this file |
|---|---|
| The window, menu, run loop, run summary workbook | `app.py` |
| Clients storage, settings, history, folders | `core/store.py` |
| Module contract, log/ask/record plumbing | `core/runner.py` |
| Login / logout / one-session-per-client | `core/session.py` (+ `login()` region of `engines/ledger_engine.py`, lines ~250–370) |
| Returns page options | `modules/returns.py` |
| Returns portal behaviour (dates, categories, print, mhtml, paging) | `engines/returns_pw.py` |
| Ledger/certificate/profile page options | `modules/ledger.py` |
| Ledger portal behaviour | the relevant function of `engines/ledger_engine.py` (it is 4,400 lines — send only the function and its line numbers) |
| WHT Receivable page options and file naming | `modules/wht_receivable.py` |
| WHT portal behaviour (calendar, search, scrape, page copy) | `engines/wht_engine.py` |
| WHT Excel/PDF/CSV layout and data cleaning | `engines/wht_outputs.py` |
| Settings page | `modules/settings_page.py` |
| Build | `BUILD_SUITE.bat`, `requirements.txt` (exact versions), `requirements-build.txt` |
| Portal question policy | `engines/dialogs.py` |
| Excel/CSV text safety | `engines/sheet_safety.py` |
| Portal Excel downloads (HTML table -> real .xlsx, repair) | `engines/excel_files.py` |
| Ledgers workbook per tax type (Summary, Level 3-2, a tab per month) | `engines/ledger_taxhead.py` |
| Tests | `tests/auto/` (pytest, mock portal served automatically); `tests/scripts/` = old print-only scripts |

## Architecture in 12 lines
- `app.py` builds the window (left vertical menu with a tick box per module; right: module page, shared client list, Run selected / Run this page / Stop, yellow question bar, activity log) and runs ONE worker thread.
- Worker, per ticked client: `with core.session.open_session(ctx, client) as page:` → logs in once → runs every ticked module in menu order on that `page` → logs out. Module errors are contained; `Stopped` ends the run.
- `core/runner.py`: `Module` (id, title, build_panel, collect, run(ctx, client, options, page)); `RunContext` (log, status, progress, check → raises Stopped, sleep, ask(message, kind 'login'|'dates'|'continue', defaults, poll), set_scope, record, items, output_root, log_path); `PrintToLog` redirects engine print() into the log.
- `core/store.py`: `%LOCALAPPDATA%\URA_Suite\{clients.json (DPAPI), settings.json, history.json, logs\}`; `load_clients/save_clients/load_settings/save_settings/add_history`.
- `core/session.py`: `launch_browser` (Chrome → Edge → bundled, honours BROWSER_CHANNEL/PROXY), `login()` from the Ledger engine (fills TIN+password, yellow bar waits for the code, refills on rejection), logout on exit.
- `engines/ledger_engine.py`: former Tax Statement Downloader; entry `run_job_on_page(page, settings)`; settings keys `tin, password, period_from, period_to (YYYY-MM), tax_types[], other_jobs[] (profile|tin_cert|wht_tcc|paye_tcc), skip_nil, annual_l3, annual_all_years`; suite-set globals `OUTPUT_ROOT, WORK_DIR, BROWSER_CHANNEL, PROXY, NAV_TIMEOUT_MS, SPEED_UPS, DEBUG_DUMPS, PROMPT_HANDLER, STOP_EVENT, RESULTS[]`.
- `engines/returns_pw.py`: `run_on_page(page, login_id, date_from, date_to, wanted_categories|None, log, check, base_folder, patient, result, skip_existing)`; opens "View return history", per category: fill dates, select type, Headquarter→Both for BRANCH_CODES, Search, for each row with a View link open tab → wait for #printWthSchid/#printid → click "Print with Schedules" (window.print disabled) → DevTools `Page.printToPDF`; excise → `Page.captureSnapshot` .mhtml; Next-page loop; retries; `write_status_sheet` → Returns_Status.xlsx; `result['items']` for the summary.
- `engines/wht_engine.py`: `SEL` dict (login, menu item "Generate Tax Credit Certificate for Withholding", #frmDt/#toDt, date icon, #generateFor, Search, results table.FORM_TAB_BORDER, Next, logout, dismiss); `WhtBot(page, log, check)` with `open_wht_page, form_frame, pick_month (CAL_SCAN JavaScript tags calendar arrows/day 1), sync_to_date, set_generate_for, search, scrape, scrape_paged, page_pdf (CDP), page_html, logout`; `run_months(bot, months[MonthResult], want_pdf, want_screenshot, want_html, generate_for 'D'|'P', continue_on_error, on_result)`.
- `engines/wht_outputs.py`: `MonthResult`, `month_sequence`, `clean_frame` (flatten 2-row headers, drop checkbox col, drop Total row, IDs → text, keep empty real columns), `align_columns`, `check_period`, `save_excel` (Summary + Combined + sheet per month, navy header, totals, freeze, filter), `save_table_pdf` (landscape, page x of y), `merge_pdfs` (bookmark per month), `save_csv`, `save_month_excels`.
- Summary: `app.write_run_summary(rows, items, path, master, run_stamp)` → sheets Overview / Files (hyperlinks) / Failed / Other outcomes; master `URA_Suite_Summary.xlsx` appended per run.
- Output tree: `<output_root>\Returns\<TIN>\<type>\`, `\Ledgers\<TIN>\`, `\WHT_Receivable\<TIN>\`, `Run_Summary_<stamp>.xlsx`, `URA_Suite_Summary.xlsx`.
- Settings page keys: `suite.{browser auto|chrome|msedge, proxy, nav_timeout_s 45, login_timeout_s 420, skip_done_today, open_folder_when_done, summary_xlsx, min_free_mb 500}`.

## Robustness rules already in the code (keep them)
- Before any Search/Next, mark old results `data-bot-stale` and wait for fresh ones (returns_pw, wht_engine). Never read the page right after clicking.
- A timeout is a FAILED item with a reason, never "no records" and never a saved half-page.
- Modules forward items to `ctx.record` in `finally` (Stop must not hide work).
- Long waits print a heartbeat at least every 30 s.
- BUILD_SUITE.bat keeps `--collect-all playwright`.
- Mock portal + tests in `tests/` — run `python -m pytest` after any change (needs Chromium: set URA_TEST_CHROMIUM
  or run `python -m playwright install chromium`).

## Known untested parts (no portal access in Claude's sandbox)
Returns: whether the live portal's VAT 'Print with Schedules' calls print() (if not, each VAT return takes the
5.5 s quiet wait - URA no longer gives VAT schedules, user 4 Oct 2026);
WHT calendar fix of 10.6 (tested with the portal's real calendar code, not yet on the live page); DPAPI (Windows only); PyInstaller build. `RESULT_MATCH="TIN"` in wht_engine is an assumption. Fixed `nap()` pauses in ledger_engine not yet reduced.

## How to ask
"Here is CLAUDE_CONTEXT.md and `engines/returns_pw.py`. On my PC the PDF for VAT returns has no schedules; log lines: … Please change the schedule handling and return the full file."

## Appearance
Colours and fonts come only from `core/theme.py` (designs: ledger, revenue, runboard). Send that file plus `app.py`
for visual changes. Never hard-code hex colours in app.py or modules; module pages use ttk styles Head.TLabel / Muted.TLabel.

## Controls
Use `core.widgets.WebButton(parent, text, command, kind='primary'|'secondary'|'danger'|'prompt'|'prompt_secondary')`
for every button, `TextField` / `PasswordField` for inputs. Do not use tk.Button / ttk.Button in new code.
Never name an attribute `_w` on a tkinter widget subclass (tkinter uses it internally).

## Module order rule
Every module may assume it starts on the portal HOME page; app.py guarantees it via core.session.go_home. A new
module must not rely on where the previous one left the portal. WHT generate_for is a list (Payee/Payer).

## Live portal facts (from the 3 Oct 2026 run - trust these over older code comments)
- Frames after login: 0 = https://ura.go.ug/en/etax-login/ (outer; never reload it), 1 = .../Portal/loginController/
  loadHomePageContent (menus), plus about:blank and the chatbot frame.
- Tax Statement opens .../Portal/accountReportDetailedController/loadFirstPage; Level 2 export button:
  <button onclick="exportTableToExcel('tblData', 'Level-2')">Export To Excel</button> (mouse click unreliable).
- WHT certificate page: frame 'tcs-resource' .../Portal/genTaxCreditCertController/searchTaxCreditCertificate; Search
  reloads the page. Generate For options: '' / D 'Payee ( Deductee )' / W 'Withholding Agent'.
- UNIVERSAL RULE (user, 6 Oct 2026), login to logout, every module: a pop-up / notice / banner that may not
  appear is checked AT ONCE (count + is_visible) and closed only if showing - never `click(timeout=...)`,
  `wait_for(...)` or a fixed pause waiting for it to appear. Waits for things the program NEEDS (a page, a
  download, a button that must come) are not pop-ups: they keep waits that end as soon as the thing appears.
  Guard: tests/auto/test_popup_rule.py times every pop-up routine with nothing on screen (< 1 s).
- On this portal a mouse click can work yet be reported as timed out: use script clicks + verify the result.
  EXCEPTION - the WHT date selector is "The DHTML Calendar" (Dynarch): it ignores script clicks (el.click());
  press its cells with wht_engine.PRESS_JS (mouse down, up, click on the element). Never force-click Search:
  a forced click presses a screen position and on 5 Oct 2026 landed on the open calendar. Test page with the
  real calendar code: tests/mock_portal/Portal/wht_dynarch.html.
- Ledger Level 2 -> Level 3 (live capture 6 Oct 2026): link `<a onclick="openLevelThreeTwo('MON-YYYY')">`,
  function points POST form `levelTwoOfAR` at `/Portal/accountReportDetailedController/loadLevelThreeTwo?period=`
  and submits; hidden fields include _csrf and lastEventTS. 10.8 loads it into a hidden frame (option C) with a
  per-tax-type proof against the click route and automatic fallback. Test copy: tests/mock_portal/ledger_server.py.
- Every wait for the person must keep the browser serviced (ledger_engine.BROWSER_KEEPALIVE): with the
  'faster loading' filter on, requests are held while the program does nothing with the browser.
- Ledger Level 3-2: find the box with ledger_engine.find_level3_box (all frames); the frame chosen by its
  address can lack the box. Back to Level 1: recover_to_level1 waits LEVEL1_PATIENCE_S, then ensure_portal_home.
- To see a live page: ask the user for Open portal (manual) > Save this page for Claude (zip in the log folder).

- WHT page (recording): results grid = #GenTaxCreditCert table[3], tick box td[1], #selectAll, #generateTCCBtn ->
  #TCCSummary (Back button inside) -> search again -> TCC numbers in #tccNo<n>. No PDF is produced.
- Tax Statement link is #mailbox ul li[7] a; Level 2 payment link = #tblData tbody tr[n] td[7] a; Level 3 Back = #backbtn.
- Outer page "Processing" modal: #process-static-modal, Close = button.btn-secondary.
- Way back to the portal from the public site: breadcrumb Home > #middle40Drop (Login alert) > Portal > Skip.
- "Generate Copy of Tax Credit Certificate for WTH": form #copyWTHTCC (#retBrnch H/B, dates, #deducteeTIN, #tccRefNo).

- PAYE certificates (recording 4:23 PM): menu 'Generate Tax Credit Certificate for Pay As You Earn'; #retBrnch H;
  #period values '2025-2026' etc. (July-June); grid nested in #TaxCreditCertForm table[2]; tick box #listPaye;
  'Get Month-Wise Details' (column 8) -> a SEPARATE POP-UP WINDOW (window.open 'post.htm' named
  'Uganda_Revenue_Authority' + form POST to getMonthsDtls) with #divPostContent and Close (corrected 8 Oct 2026:
  not the outer page). Search = form POST searchPayeEntry that replaces the page; GenerateTCC = POST
  generatePAYETCC, at most 5 ticks per press (var count = 5). Engine: engines/paye_engine.py; test copy of the
  live screens: tests/mock_portal/paye_server.py. Ask for a page capture of #divPostContent if its columns ever look wrong
  (the first one is saved as logs/paye_monthwise_example.html on every run).

## Version 10.18 notes (12 Oct 2026 - live log 12:56, 10.16 build from abroad)
- Returns: run_on_page(reopen=) + nested open_history / recover_history; a failed search or missing date boxes
  -> page copy in the client's Returns folder, go_home, 'View return history' again (once per type).
- WHT: open_wht_page waits for a certificate page the browser has asked for (WhtBot._on_request notes requests to
  WHT_PAGE_URL_MARK; _wait_for_loading_form, PAGE_OPEN_S) instead of re-clicking the menu link.
- Ledgers: process_tax_type contains a month's exception (recorded 'failed', screenshot, refresh_level2, next month).
- Tests: 233.

## Version 10.17 notes (12 Oct 2026 - run analysis Tronics vs Telecom)
- PAYE Search: ONLY the form's own button, in the form's frame (paye_engine.SEL['search'] - never a bare
  "button:has-text('Search')": the outer ura.go.ug page has a site search that sends nothing). A press that sends
  nothing is described in the log, pressed again another way, and given up after PRESS_AGAIN_S + UNSENT_GIVE_UP_S;
  SearchTimeout.sent=False keeps that year out of the end-of-client second try. save_evidence numbers repeated stems.
- Ledger Back after Level 3: ledger_engine.back_to_level2 (page requests show whether the press took; BACK_TOOK_S,
  LEVEL2_AFTER_BACK_S). ledger_table(require_months=False) / export_level2_to_excel(has_months=) for empty ledgers.
- Direct way proof: once per client (DIRECT_STATE[CLIENT_PROOF_KEY]); DIRECT_STATE is cleared per client run and by
  the ledger_portal test fixture.
- WHT calendar: pick_month takes months-only when it needs no more presses than years-then-months.
- app.py Overview 'Files' = file rows recorded by the step (also after an error). core.session.logout goes to the
  portal home page once when no Logout button is on screen.
- Tests: 229. The window run-loop tests stub core.net.wait_online (no internet needed).

## Version 10 notes
- Steps: 1 Returns, 2 Ledgers, 3 Certificates (WHT + PAYE), 4 Review; one shared period (year-month) for all.
- Level 1 (screenshot): table 'Tax Type | Outstanding Tax' with Value Added Tax, VAT-WHT, VAT-Imported Services,
  Income Tax, Rental, IT-WHT, PAYE, Gaming Tax, Goods Cause Tax, Local Excise Duty, Stamp Duty; buttons Back /
  Export To Excel / Export To PDF. It can take ~100 s to fill after Back on a slow portal.
- Level 2 (screenshot): S/N, Period, Return, Assessment, Penalty, Interest, Payment [PRN, WHT Credit, Offset Received
  (Other Period)], Amount Allocated to Other Periods, Tax waiver, Net Position. Level 3-2 'Details Of Payment': Sr No.,
  Payment Basis, Effective Date, Payment Registration Number, Amount, Adjusted Against Period, Liability Basis,
  Principal, Penalty, Interest, Refund.
- No PDFs for WHT/PAYE (user decision 3 Oct 2026).
