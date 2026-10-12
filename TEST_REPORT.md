# URA Suite — Test Report (3 October 2026)

> **Since the audit (3 Oct 2026, evening):** the checks below are now automated in `tests/auto/` and run with
> `python -m pytest`: 161 tests since 10.11 (95 before 10.6), all passing on the fixed code; 28 of the first 45 failed on the code before the fixes,
> which shows they catch the defects. The scripts this report describes are kept in `tests/scripts/` for reference.

## How it was tested
A real Chromium browser (version 153) driven by the real engines, against a **mock URA portal** built to copy
the portal's structure: menu in the outer page with content in an iframe, a login frame named `tcs-resource`,
return-history search with two-page results and the Headquarter/Both box, return tabs whose Print buttons
appear late, one return that never loads, a WHT page with a working date selector whose To date is filled
only on blur, a "Processing" box, and results that arrive after a delay. The mock and the test scripts are in
`tests/`. Plus: clean-environment import with only `requirements.txt`, and a real PyInstaller build.

**Not covered:** the live URA portal (its exact page timings and wording), Windows DPAPI, the Ledgers engine's
portal pages (not mocked; only its failure containment was tested).

## Defects found and fixed
| # | Defect | Severity | How found | Fix |
|---|---|---|---|---|
| 1 | Build failed: Returns page imported the old Selenium program | Blocker | your build | Selenium-free helpers moved to `engines/returns_common.py` |
| 2 | Returns read the **previous** search's table (PAYE showed VAT rows) | High | mock, real browser | old results marked stale before Search; only new rows read |
| 3 | A return that never loaded was **printed and logged "saved"** | High | mock | no Print button within the limit = failed, retried, recorded |
| 4 | WHT: slow portal → month recorded **"No records" though it had data** | High | mock with 6 s delay | waits for new results; too slow = failed with reason, never "no records" |
| 5 | WHT fix #4, first version, threw away a new table drawn into an old container | High | mock | old element removed only if nothing new appeared inside it |
| 6 | `.exe` would not contain Playwright's driver → could not open any browser | Blocker | PyInstaller hook check + real build | `--collect-all playwright` restored in BUILD_SUITE.bat |
| 7 | Stop: files saved before Stop missing from the summary | Medium | Stop test | modules report their items in a `finally` |
| 8 | Stop during WHT: captured months not saved; missing months not shown | Medium | Stop test | captured months saved; others listed "Not run - stopped by you" |
| 9 | Mistyped security code cost 60 s before re-asking | Medium | login test | rejection detected at once (`login_rejected`) |
| 10 | Ledgers "Speed-ups" tick ignored (session opened before options read) | Low | login test log | applied before the session starts |
| 11 | "Wait until each return downloads" silent for up to 15 min; guide contradicted the default | Medium | full-app test looked frozen | 30-second "still waiting" heartbeat; guide corrected |
| 12 | Overview said "Completed" when items inside had failed | Low | summary review | "Completed - N item(s) failed" |
| 13 | WHT month had one attempt; other modules retry | Low | review | one retry per month |

## Results after fixes
| Scenario | Result |
|---|---|
| Returns engine: 3 types, 2 result pages, one broken return, no-records type, excise .mhtml | correct; PDFs contain schedules and full table width; status sheet filled |
| Returns re-run | already-saved returns skipped |
| WHT engine: calendar from Oct back to Jun, To date confirmed, Processing box, 3 months | correct, totals correct |
| WHT with results 6 s late | correct |
| WHT slower than the limit | each month failed with reason, retried once |
| Login: first code rejected, second accepted; logout | correct, re-ask within 1 s |
| Full app: 2 clients x 3 modules, one login each, Ledgers failing | failure contained, WHT ran on same login, 20 files, summary 4 sheets correct |
| Stop during Returns (patient mode) | stopped in 3-4 s, logged out, files and partial status sheet reported |
| Stop during WHT | captured month saved, remaining months listed as not run |
| Clean environment, no Selenium | `import app` OK, all 13 modules import |
| PyInstaller build | 103 MB, contains driver; executable starts and shows the window |

## Diagnostics for your PC
Set the environment variable `URA_DEBUG=1` before starting the program to get step-by-step timing lines from
the Returns engine in the log — send those lines if a return behaves oddly.

## Widgets (3 Oct 2026)
Real-event tests: face moves 0 -> 3 px when pressed; fires on release, not on press; press-and-slide-off does not fire;
Space fires; disabled does nothing; password hidden as bullets, eye click shows it, value unchanged, click again hides.
Run/Stop states through a run: during = Run disabled / Stop enabled, after = reversed. Found and fixed: attribute `_w`
clashed with tkinter's internal name; eye click needed a prior mouse move (now position-based).

## Module switching and WHT Payee/Payer (3 Oct 2026)
Mock portal rebuilt to the structure the Ledgers engine expects (portal frame at .../Portal/loginController/
loadHomePageContent, menus in #mailbox, screens loading in place). Returns -> Ledgers -> WHT on one login:
- WITHOUT the fix: Ledgers "Could not find a frame containing heading 'Reports' within 90s", WHT "Could not open the
  WHT search form" - the defect you reported, reproduced. (My earlier full-app test showed the same Ledgers error and
  I wrongly put it down to the mock; it was this defect.)
- WITH the fix (core/session.go_home before every module after the first): portal home reached before each module,
  Ledgers clicked Reports -> Tax Statement (stops later only because the mock has no ledger form), WHT ran Payee and
  Payer: separate files, different data (Agent rows vs Supplier rows).
WHT 'Generate For' values D = Payee (Deductee), P = Payer (Deductor) taken from the original program's bytecode; the
selection is now read back and a wrong or missing option stops that month with a clear reason.

## First live-portal run (Global 7, 3 Oct 2026) - findings and fixes
| Live symptom (from your log) | Cause | Fix | Verified |
|---|---|---|---|
| Returns not run | "Resume" counted an earlier run as complete | only a clean "Completed" with files counts, and the program now ASKS before skipping | code |
| Payer: every month failed, 90 s each | live portal values are D = Payee, **W = Withholding Agent** (not P) | W used; a missing option stops the type once | mock with live options |
| WHT Feb-Jun "7 rows x 4 columns" | the search form's layout table read as data | grid must have a TIN heading, 5+ columns and not hold the form's fields | mock with live-like reload |
| WHT Jan "too slow" then "could not set the period" | portal answers by reloading the page; no-records page has no message | page reload detected = answered; no grid within 6 s = no records (page copy kept) | mock |
| Ledgers "Export To Excel" 45 s time-outs, then cascade | mouse click "outside of the viewport" | export by script; if no file in 30 s, Excel built from the table on screen | real browser |
| Level 3: ~55 s lost per month ("Could not click payment link") | click worked but was reported as timed out; fallback waited 45 s for a link already gone | one script click; a vanished element = click went through; js_click limit 5 s | real browser |
| Certificates "Ticked 0 of 8" -> "select at least one Taxpayer" | a click that worked + a fallback click = unticked again | set each box to the wanted state once, by script, then read back | real browser |
Not yet verifiable without the live page: whether the export now always downloads (the fallback covers it), and the exact
live HTML of the Level 2 / certificate / WHT result pages - use Open portal (manual) > Save this page for Claude.

## Live recording (3 Oct 2026) -> Generate TCC
Recording showed: results = 3rd table in #GenTaxCreditCert, tick box in the first cell, #selectAll, button
#generateTCCBtn -> summary #TCCSummary with a Back button, no PDF; TCC numbers in #tccNo1, #tccNo2...; the
"Processing" box is #process-static-modal in the OUTER page. Mock built to match (confirm box, "select at least one"
alert, summary + Back, numbers only after a new search, numbers held in text boxes).
Result: rows without a number ticked and generated, confirm accepted, month searched again, numbers captured and kept
as text in Excel (0NAM27012732931). Rows that already had a number were not touched. Regression caught: with no
Generate button the month's data was lost - fixed, the data is now kept and the reason recorded.
Ledgers: its WHT certificate job is now done by WHT Receivable; its PAYE certificate step no longer waits 90 s for a
PDF nor hands over to you (records "generate pressed").

## PAYE certificates (live recording 3 Oct 2026 4:23 PM + screenshot)
New engines/paye_engine.py built from the recording: #retBrnch = H, #period value 'YYYY-YYYY' ('July 2025 to June 2026'
= 2025-2026), Search in #TaxCreditCertForm, nested employees grid (innermost table), tick box #listPaye, GenerateTCC,
'Get Month-Wise Details' opens #divPostContent in the OUTER page with Close at its foot.
Financial years: 2025-05..2026-06 -> 2024-2025, 2025-2026; 2025-07..2026-06 -> 2025-2026; 2023-01..2025-12 -> 4 years.
Mock test: a year with no employees, a year with all numbers, a year with 2 missing -> generated, searched again,
0 left; month-wise read for every employee (12 months + Total) and the box closed each time; 2 workbooks.
Defects found by the tests and fixed: the previous year's "No Records Found" message was read as the next year's
answer (old messages now marked too); the outer wrapper table was read instead of the nested grid; pandas keeps empty
cells as missing after text conversion (text joins made safe - same issue fixed in the WHT period check).
Full app: Ledgers with only PAYE ticked -> login, both years, generation, month-wise, 2 workbooks, run summary.

## Version 10 (3 Oct 2026, after the first full live run)
| Live finding | Fix | Test |
|---|---|---|
| IT-WHT "Could not find text within 20 s" x3; PAYE ok only on try 3 | wait for the Level 1 TABLE to fill (up to 120 s x speed, heartbeat), script click, reopen Tax Statement once | mock with table arriving at 25 s: old search fails exactly as in the log; new step opens Level 2 |
| Return names begin with the row number (shifts monthly) and are cut mid-word | names from return no. + period + filing date; duplicates judged by return no. (old names recognised) | unit check |
| ~60 s silent waits while the portal opened some returns | new tab watched directly, heartbeat every 30 s | Returns regression |
| 11 of 16 VAT months had payment 0 but opened Level 3 | "skip nil-payment months" on by default | - |
| Slow internet | Internet speed factor x1/x2/x4 on every engine; connection check before each retry and each client | - |
| No PDFs wanted for WHT/PAYE | PDF options removed | full app run: only .xlsx/.html |
| Client briefing layout | shared period, steps 1-4, Certificates = WHT + PAYE, Review page | full app run on the mock + screenshots |

## Version 10.6 (6 Oct 2026, after the live runs of 5 Oct 2026)
Each new test was run on 10.5 first and failed with the live symptom, then passed on 10.6.
| Live finding | Fix | Test (result on 10.5 -> 10.6) |
|---|---|---|
| WHT: "Stepped 60 month(s) to 'October, 2026'", ~110 s per month | mouse-like press (PRESS_JS), stop after 2 dead presses | real calendar: 109.9 s and wrong date -> 2.2 s, 01/07/2025 |
| WHT: wrong From date (always a Tuesday), Search never pressed, "portal too slow" | close the calendar, no forced Search click, From date checked before and after | forced click changed 01/07/2025 to 10/07/2025, no search sent -> search sent with 01/07/2025 |
| WHT: January (found while fixing) - week number 1 could be taken for day 1 | skip 'wk' cells | fails -> passes |
| Ledgers: a recovery error ended the whole step | retry loop protects the recovery | TimeoutError escaped -> 3 attempts, then the next tax type |
| Ledgers: Portal link clicked nothing (5 of 5) | every match checked; route via ensure_portal_home | hidden entry first: no click -> visible link clicked |
| Ledgers: Level 1 given 10 s | up to 90 s x speed | Level 1 at 4 s (x0.25 scale): Reports timeout -> "Back on Level 1." |
| Ledgers: 45 s wait for the Level 3-2 box in the wrong frame | find_level3_box searches every frame | box found in the inner frame |
| Returns: page 2 dropped after ~13 s (client B VAT) | up to 60 s (180 s patient) | 15 s page turn: "treating this as the last page" -> page 2 processed |
Full run: 108 passed. Same results old vs new on all existing practice pages.

## Version 10.7 (6 Oct 2026)
| Change | Test (result before -> after) |
|---|---|
| Pre-login notice absent / hidden | 5.0 s -> under 0.5 s; a showing notice is still closed |
| Every pop-up routine with nothing on screen (7 routines) | 6 already under 1 s; pre-login notice 5 s -> under 1 s |
| Test Login logout link, hidden one first | helper missing -> visible link pressed at once |
| One Level 3 month on the practice page | 7.4 s -> 2.4 s; times line and screen-copy line in the log |
| Screen copy equal / not equal to the portal's file | reported "identical" / "DIFFERENT"; the saved file is the portal's in both |
Full run: 122 passed. Old (10.5) and new give identical Level 3 files, WHT months and Returns.

## Version 10.8 (6 Oct 2026)
| Change | Test (result before -> after) |
|---|---|
| Option C, 3 months, session token | feature missing -> all saved, Level 2 never rebuilt |
| Background files vs click-route files | identical except the download-time line (which always differs) |
| First month proved, then background | route check "identical"; only month 1 used Back |
| Portal refusing a repeated form (one-use token) | falls back, all 3 months saved |
| Manual mode, filter on, second frame after login | not loaded during the wait -> loads |
| hand_over during a run | not loaded during the wait -> loads |
Full run: 129 passed (run in three parts). Click-route Level 3 files identical to 10.5.

## Version 10.9 (6 Oct 2026)
| Change | Test (result before -> after) |
|---|---|
| One table per tax head, oldest first, PRN text / amounts numbers | module missing -> passes |
| Next run updates the same file (re-fetched month replaced, others kept, no new file) | passes |
| 10.8 monthly files moved, not fetched again; real monthly routine on the Ledger test portal | passes |
| Update time, 12 months x 15 payments | under 2 s (three live-size tax heads: 0.25 s) |
| Option B check off by default | passes |
Full run: 135 passed (three parts).

## Version 10.10 (6 Oct 2026)
| Change | Test (result before -> after) |
|---|---|
| Window at 1366x768, 1280x720, 1024x768 | larger than the screen -> fits |
| 1920x1080 at 150 % | Run button not shown, text not scaled -> shown, 1.5x |
| 60 ticked clients, left run list | 60 rows -> 12 at most |
| Excel import: new / update / problem rows, headings in any order, template | module missing -> passes |
| Import through the real Clients window: password stored encrypted | passes |
| Voice phrases, sample 1/4 mapping, background, off, text quoting | module missing -> passes |
Full run: 154 passed (three parts). Soak: 40 clients back to back through open_session, memory +5 MB.

## Version 10.12 (8 Oct 2026)
| Change | Test (result before -> after) |
|---|---|
| Workbook layout | only 'Level 3-2' -> Summary, Level 3-2, a tab per month equal to the portal's file (cells, merges) |
| Monthly files after the step | 14 per tax type left in _months -> none (kept only with the option) |
| Month only inside the workbook | fetched again -> counts as downloaded; workbook deleted -> fetched again |
| Workbook open in Excel | (n/a) -> monthly files kept, folded in on the next run |
| 10.9 workbook + files | tabs built from the files, a lost file's month rebuilt from the table |
| Real routine on the test copy of the live Ledger form | second run fetches nothing |
| _route_check.json | inside _months\<tax> -> client folder, old copy merged |
| PAYE search, portal answers after 3 s, a widget frame reloads | 'No employees listed' -> 3 employees listed |
| PAYE year with no list | 'no records', no reason -> reason recorded, page copy + screenshot saved |
| PAYE month-wise details (pop-up window) | 30 s then failed per employee -> 12 months each, pop-ups closed |
| PAYE generate, 7 without TCC, limit 5 per press | 5 made, 2 'incomplete' -> 7 made in 2 batches |
| Withholding Tax Return default | not ticked -> ticked |
Full run: 175 passed (3 run-loop tests need ura.go.ug reachable or the network check stubbed).

## Version 10.11 (6 Oct 2026)
| Change | Test (result before -> after) |
|---|---|
| Chrome closed during client 1's Returns | Ledgers, Certificates and client 2 all started -> nothing more started, one line, "Job stopped." |
| Chrome closed during the last step, which carries on | client 2 started -> run stops |
| Two clients, voice | "Job completed." twice -> "Job completed. Starting another." then "Job completed." |
| Ledgers on a closed browser | 3 attempts x every tax type, 90 s waits -> stops at the first, wait ends at once |
| Returns message | "Stopped - ..." -> "The browser window was closed - ..." |
Full run: 161 passed (three parts); the run-loop tests passed three times in a row.

## Version 10.15 (8 Oct 2026)
| Change | Test (result before -> after) |
|---|---|
| Ledger check, one PRN in two periods / part not allocated / transfer in Offset column | 'check' -> OK; a real 5,000 difference still 'check' |
| Ledger check on the user's whole live output (34 months with Level 3) | VAT MAY/FEB, LED FEB/JUN flagged by the Principal rule -> all OK |
| Income Tax yearly, test copy of the captured screens | Level 3-5 opened, 30 s timeout x3, nothing saved -> one Level 3-2 per year, Income Tax.xlsx with Summary OK |
| Year with PRN and Tax Credit both 0 | opened -> not opened ('no payment link') |
| One year answered with an error page | whole tax type failed after 45 s x3 -> that year failed after 8 s with page copies, next year saved, fetched by the next run |
| Year inside the period | 3 years for a 1-year period -> 1 |
| WHT certificate vs ledger WHT Credit (return number MTO...) | matched by amount only -> 'Yes - return number'; ASYCUDA listed apart |
| PAYE Search press swallowed | 60 s then whole step failed -> pressed again after 8 s, list read |
| PAYE Search never answered twice | whole step failed -> page copy, reopened, tried again, that year failed, next year saved; second try fetches that year only |
| Step that stops with an error | not tried again -> tried once more after the client's other steps |
| Second try results | (n/a) -> first try's item updated ('saved on second try'), never listed twice |
Full run: 205 passed (run-loop tests with the network check stubbed).

## Version 10.16 (9 Oct 2026)
| Change | Test (result before -> after) |
|---|---|
| Live console test, MAR-2026 direct vs portal Export of 8 Oct | (n/a) -> identical except the time line |
| Live console test, FEB-2026 direct vs Export / PDF | (n/a) -> same 4 lines; two lines of one payment in the other order |
| Auto on the test portal, 3 months | every month click or background -> 1st month proved, 2 months direct, no extra Level 2 rebuild; files identical to the screen way |
| Portal refusing direct requests | (n/a) -> proof 'not possible', all months by the screen way |
| Direct first, refused | (n/a) -> switched at once, all months saved |
| URA answering after 4 s, heartbeat 2 s | (n/a) -> 'URA still working on it', one request only |
| Switch time 2 s, answer after 4 s | (n/a) -> screen way for that month, saved |
| Lines of one payment in another order | background proof 'different' (route switched off) -> 'same lines, different order', both ways stay on |
| Income Tax yearly, Auto | (n/a) -> 1st year proved, 2 years direct, Summary OK |
| PAYE press-again when the search was sent | pressed twice -> once |
Full run: 214 passed (run-loop tests with the network check stubbed).

