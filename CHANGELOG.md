# Changelog

## 10.16 - 9 October 2026 (user: "build direct but also keep current way")

### Level 3 by the direct way, the current (screen) way kept as the backup
- What it is: a payment link on Level 2 only asks URA for one page (openLevelThreeTwo -> POST
  loadLevelThreeTwo?period=...). The program now sends that same request from the Level 2 page itself (same
  session, same form fields) and saves the ledger table (#tblData) of URA's answer - the same table the portal's
  Export To Excel copies into its file. Nothing is drawn, there is no Back, Level 2 is never rebuilt.
- Proof on the live portal (user, 9 Oct 2026, console test from the ledger frame): URA answered 200 with the full
  Level 3-2 page. MAR-2026: equal to the portal's Export of 8 Oct cell for cell (only the 'Tax Position as at'
  time differs). FEB-2026: the same 4 lines and amounts; the two lines of one payment (same PRN and date,
  one 'Original Return', one 'Penalty Assessment') came in the other order - URA's own Export at 02:35 showed the
  first order again, so the order of such lines is not fixed.
- Offline: the user's 4 captured pages identical either way; test portal identical files; with URA answering after
  40 / 75 / 150 s the direct way took 40 / 75 / 150 s, the screen way 117 / 177 / 252 s (first month).
- How it runs (Ledgers page, "How Level 3 is fetched"):
  - **Auto (default)**: the first month (or year) of each tax type by the screen way, then the same month by the
    direct way into a scratch folder and compared. 'identical' or 'same lines, different order' -> the other months
    directly; 'different' -> the screen way for that tax type and a 'check' item on the Review page.
  - **Direct first, no proof**; **Screen way only** (as 10.15).
  - Any clear 'no' (URA error or refusal, login page, no Level 3-2 table, another month) switches that tax type to
    the screen way at once (Level 2 reloaded first). Nothing is lost.
  - The wait follows URA's answer, not a clock: "URA still working on it (N s)..." after 15 s, then every 30 s;
    after 15 min (x Internet speed setting) without any answer, the screen way.
  - Optional: "Switch to the screen way if URA has not answered after N seconds" (0 = never, the default).
- Results: `_route_check.json` has '<tax> - direct' (identical / same lines, different order / different / not
  possible / switched off); the log ends with the median time URA took per direct page.
- Also: the background way's own proof (10.8) now accepts 'same lines, different order' too - a month like VAT
  FEB-2026 would otherwise switch it off.
- Fixed: PAYE 'press Search again' (10.15) could press a second time although the first search had been sent
  (the time was taken after pressing; the browser reports the request during the press). Found by the tests.
- Tests: 214, all passing; new tests/auto/test_ledger_direct.py (9). Test portal: request kind recorded
  ('navigate' / 'cors'), 'refuse_fetch', 'flip_order'.

## 10.15 - 8 October 2026 (user: "Do all 3. Also put Retry if anything else is failed like returns.")

### 1. Ledger check: the Level 2 payment against the right Level 3 figure
- Was (10.13): Level 2 Payment compared with the total **Principal** of the month's Level 3-2 lines. Live output of
  8 Oct 2026 proved it wrong: part of a payment can stay unallocated (Principal less than the Amount) and Level 2
  still counts all of it - e.g. VAT MAY-2026 was flagged although the ledger is right.
- Now: the **total of Level 2's Payment sub-columns** (VAT: PRN | WHT Credit | Offset Received(Other Period);
  Income Tax: PRN | Tax Credit; IT-WHT, PAYE, LED: one column - read from the saved `<tax>_L2.xlsx`) is compared
  with the **Amount of every different reference** on the Level 3-2 lines (a PRN split over two periods is listed
  twice and counts once). Re-run on the whole live output: all 34 months with a Level 3 are OK (VAT, LED, PAYE,
  IT-WHT). The split between the sub-columns is shown but not compared: VAT AUG-2025 and SEP-2025 show an old
  Credit Balance Transfer (effective 01-JUL-2021) inside Level 2's PRN column.
- Summary tab, new columns: Level 3 total, Level 2 split, Level 3 PRN payments, Level 3 credits, Credits by type,
  Not allocated (per PRN, Amount minus Principal, Penalty, Interest, Refund - information only, never a 'check').

### 2. Income Tax and Rental Level 3 (yearly ledgers) work
- Was: live 8 Oct 2026 22:33 and 22:35 - 'Locator.click: Timeout 30000ms exceeded', 3 tries each, so neither tax
  type got a Level 3 and the WHT certificates could not be checked against the ledger. Cause (user's captures):
  the yearly Level 2 has two heading rows; counting headings put 'Payment' over the Interest cell, so the program
  opened Level 3-5 (Interest) and waited for an Export button that page does not have. It also opened '0' amounts
  and counted years by calendar year (3 years fetched for a one-year period).
- Now: only the links that call `openLevelThreeTwo(...)` (PRN and Tax Credit open the same page - opened once per
  year); a year whose PRN and Tax Credit are both 0 is not opened; a year counts when it overlaps the period; the
  Level 3-2 box is found in any frame and saved with the portal's Export (a copy of the on-screen table if the
  portal gives no file), like the monthly ledgers. One failing year is recorded for that year (screenshot and page
  copies in the log folder) and the next year goes on.
- The years go into `Ledgers\<TIN>\Income Tax.xlsx` / `Rental.xlsx` like the monthly tax heads: Summary (one row
  per year, Payment + Tax Credit checked), Level 3-2 (one table), a tab per year.
- A Level 3-2 address answered with an error page now stops after 8 s instead of 45 s (all tax types).

### 3. WHT certificates checked against the Income Tax ledger by return number
- The captured Level 3-2 shows each 'WHT Credit' with the withholding agent's **return number** (MTO...) as its
  reference, not a TCC number. Matching order now: Return/Assessment Number, then TCC No., then amount. The lines
  are read from `Income Tax.xlsx` (and files of earlier versions). 'ASYCUDA Credit' (customs) lines are listed
  apart on the 'Check vs ledger' tab for information, never as a 'check'.

### 4. PAYE certificates: a Search with no answer no longer fails the whole Certificates step
- Was: live 8 Oct 2026 22:40 - 'no results within 60 s after Search', the Certificates step stopped.
- Now: if the Search press sends nothing to the portal within 8 s it is pressed again (the page's own
  send-once guard is cleared first - otherwise a second press does nothing); after 60 s with no answer a page copy
  and screenshot are saved (`paye_search_timeout_<year>`), the portal home page is opened and the year searched
  once more; still nothing = that year 'failed' (the message says whether the portal received the search) and the
  next year goes on.

### 5. Second try at the end of the client for anything that failed (like the slow returns of 10.14)
- After a client's steps, once more, only for what failed: a whole step that stopped with an error; ledger tax
  types / months / years and the profile or TIN certificate; WHT certificate types with a failed month and PAYE
  years with no answer; return types with a return not downloaded (not the ones you skipped). Files already saved
  are skipped, so only the missing work is done.
- The Review page is updated, not doubled: a fixed item reads e.g. 'saved on second try' / 'done on second try';
  a still-failing one keeps 'failed' with the second reason. Nothing is put off a third time.
- Code: core.runner RunContext.final_try / record merge / resolve; app.whole_step_again; LedgerModule.failed_work /
  second_try; ReturnsModule.failed_types; WhtReceivableModule.second_try; paye_engine.run(reopen, failed, years).
- Tests: 205, all passing; new: test_ledger_yearly.py (4), test_second_try.py (5), 3 PAYE, 2 ledger check, 1 WHT
  check, 1 run-loop test. Test copies of the yearly Income Tax screens and of a PAYE Search that is swallowed or
  never answered (invented values only).

## 10.14 - 8 October 2026

### A return the portal will not open no longer holds up the client (live log 8 Oct 2026, 21:29-21:37)
- Was: a VAT return (August 2025) never opened; in 'wait until it downloads' mode the program allowed 15 min,
  nothing showed whether URA was still working, and after 8.5 min the person closed Chrome - the next return,
  Ledgers and Certificates were never done.
- Now, three parts (user's choice):
  1. **First try**: each return gets at most `FIRST_TRY_SECONDS` = 120 s from View to its Print button (live:
     the slowest return that did open took 64 s to open and 70 s to finish loading; normal ones 1-3 s). A slower
     one is **set aside** ('set aside' in the log and the run summary) and the run goes on - the next returns,
     the other return types, Ledgers, Certificates.
  2. **Second try** after the client's other steps ('----- Returns - second try for N slow return(s) -----'):
     only the set-aside returns, with the mode's full wait (15 min in 'wait until it downloads').
  3. During the second try, after `SKIP_OFFER_AFTER_SECONDS` = 60 s, the voice says **'Attention needed.'** and
     the yellow bar shows **Skip this return / Keep waiting**. The wait goes on while the bar shows; nobody at
     the PC = the full wait. Skip ends it: the return is listed for review ('skipped by you while waiting') and
     is fetched on the next run (nothing is marked done unless saved).
- Normal returns are untouched (same 1-3 s). The voice uses the Settings voice choice ('Off' = silent).
- Also: the 'still waiting for the return to finish loading' line is no longer printed twice after a slow check.
- Code: returns_common FIRST_TRY_SECONDS / SKIP_OFFER_AFTER_SECONDS; returns_pw.run_on_page(first_try_s, person),
  result['deferred']; modules/returns.second_try; core.runner RunContext.attention / poll_answer / hide_question /
  after_client; app.py runs ctx.after_client after a client's steps; question-bar kind 'skipwait';
  core.voice.attention_phrase.
- Tests: 189, all passing; new: tests/auto/test_slow_return_set_aside.py (5) + 1 run-loop test in test_window.py; test history page has a
  return whose View opens only after a set delay.

## 10.13 - 8 October 2026 (ideas taken from the ura-returns-tax-statement-download skill, user's choice 1-4)

### Returns named for the audit file
- Monthly returns: `<n>. <Month> <YYYY>.pdf` - n is the month's place in the period asked for (first month = 1),
  the month taken from the row's Return Period From. Amended: `<n>. <Month> <YYYY> - Amended filed <dd-mm-yyyy>.pdf`
  (the filing date keeps several amendments apart and never changes between runs). A branch other than the
  main one: ` - <Branch>`. Yearly returns: `Original Prov 2025-26.pdf`, `Amended Prov 2025-26 filed ....pdf`
  (Income Tax returns 'ITR'). Local Excise Duty keeps .mhtml, same names.
- Original / Amended and the branch are read from the row's own cells (Application Type, Applied For Branch).
- The return number behind each name is kept in `_saved_as.json` in each return-type folder and in a new
  'Saved as' column of Returns_Status.xlsx; 'already saved' is decided by return number, so a name can never
  hide a different return (two returns that would share a name get ' (filed dd-mm-yyyy)').
- Returns saved by 10.12 or earlier (named by return number) are renamed on the next run, not downloaded again.
- A row without a readable period keeps the old name.

### 'No Records Found' screenshot
- When a return type has no returns for the period, the search screen (Form For, Return For = Both, the dates,
  the portal's message) is saved as `<type folder>\No Records Found <2025-26>.png` and linked from the run summary.

### Ledgers: Level 2 payment checked against the month's Level 3-2 lines
- Each tax-type workbook's Summary gains 'Level 3 principal', 'Difference' and 'Check'. Check = OK when the Level 2
  payment equals the month's Level 3-2 Principal total (the skill's rule; `CHECK_COLUMN` in
  engines/ledger_taxhead.py, Amount when a month has no Principal column); otherwise
  'check: ... differs by N', naming any line adjusted against another period. Each 'check' month is also a
  Review item. To be confirmed on the first live run: that Principal is the right column.

### WHT certificates checked against the Income Tax ledger
- The Payee WHT workbook gets a 'Check vs ledger' sheet: every certificate with 'Found in Income Tax ledger'
  (Yes - TCC No. / Yes - amount / No), then the ledger's WHT credit lines no certificate matches. Each gap is a
  Review item. Needs Income Tax with its Level 3 ticked on the Ledgers page (otherwise the sheet says so).
  engines/wht_crosscheck.py; the Income Tax Level 3 layout is not assumed (TCC No. anywhere in a line, else the
  amount).

### Tests
- 183 tests (175 + 8 new), all passing; the new ones fail on 10.12.

## 10.12 - 8 October 2026

### Ledgers: no Excel file per month - one workbook per tax type, like the WHT workbook (user, 8 Oct 2026)
- `Ledgers\<TIN>\<Tax head>.xlsx` now has: **Summary** (one line per month: Period, Status, Payment (Level 2),
  Level 3 rows, Level 3 amount, Tax Position as at, Last checked, Detail - an Excel table), **Level 3-2** (all
  months in one table, unchanged from 10.9) and **one tab per month**, oldest first, holding the portal's Level 3
  exactly as exported (same cells, merged headings, number formats; PRNs stay text, text never a formula).
- The monthly downloads still land in `_months\<Tax head>\` during the step (resume after Stop, the
  background-route proof, the conversion check); at the end of the step each one is removed - with its
  `portal_original` copy - only after the workbook has been written and read back with that month's tab in it.
  Empty `_months` folders are removed. Workbook open in Excel -> nothing removed, folded in on the next run.
- A month inside the workbook counts as downloaded (`month_file_done` -> `ledger_taxhead.months_in_workbook`), so
  removing the file never causes a re-download. Workbook deleted by hand -> the month is fetched again.
- Summary statuses come from the run (nil payment - skipped, not in ledger, no payment link, failed) and are kept
  from earlier runs for months not looked at again.
- Ledgers page option "Also keep a separate Excel file for every Level 3 month" (off; settings key
  `keep_month_files`, engine `KEEP_MONTH_FILES`).
- First run on a 10.9-10.11 folder: the month files already in `_months` become tabs and are removed; a month
  whose file is gone gets a tab rebuilt from its Level 3-2 lines (marked as rebuilt).
- Run summary: each month's row links to the tax-type workbook instead of a file that is removed.

### PAYE certificates - three faults found from the user's capture of 8 Oct 2026 (Recorder JSON + network record)
The capture shows the portal answering the same search the program sends (Return For = Headquarter,
period 2025-2026) in under a second with 75 employees (72 without a TCC No.). The program's own reading code
reads that page correctly; the faults were around it:
- **'No employees listed' (live, 6 Oct)**: after Search the program took ANY frame that had freshly loaded
  as the portal's answer. While the portal took a few seconds, another frame reloading (a chat or tracking
  widget) was taken, 6 s passed without a list in it, and the year was recorded as 'no records'. Reproduced on
  a test copy of the live page (search answered after 3 s, a widget frame re-created every 1.5 s). Now only
  the frame holding the PAYE form counts; once that page is back, a list gets up to 10 s to draw (ends as soon
  as it shows). When the program does conclude 'no employees', the log and Review say why (the portal's
  message, or 'the PAYE page came back with no list') and a page copy and screenshot are saved in the log
  folder (paye_no_employees_<year>.html/.png). The log now says how long the list took after Search.
- **Month-wise details**: 'Get Month-Wise Details' opens a SEPARATE pop-up window (the portal's
  OpenWindowWithPostWithParam -> getMonthsDtls); the 3 Oct recording shows the Close press inside that window
  too. The program looked in the main window, waited 30 s per employee and failed each one. Now read from the
  pop-up, which is then closed (the old same-window box is still understood).
- **Generate missing TCC numbers** (off unless ticked): the page accepts at most 5 ticks per press
  (`var count = 5`; a 6th tick is refused with 'You can generate maximum 5 certificates at a time'). The program
  ticked all and pressed once - 5 made, the rest 'incomplete'. Now in batches of the page's limit, searching
  again after each, until none are left or a batch makes no progress. Rows the portal marks as already
  generated (hidden allowTCC = 'N') are never ticked.
- WHT certificates: the same 'any reloaded frame' loophole closed ('reloaded' counts only for the WHT page
  itself). Not seen in the live logs; how 'No records' is recognised is unchanged.
- Test copy of the live PAYE screens: tests/mock_portal/paye_server.py (+ Portal/paye_outer.html,
  paye_outer_busy.html), invented employees only. Recording kept as tests/live_recording_paye_2026-10-08.json
  (no password or client data in it).

### Returns: Withholding Tax Return now ticked by default
- The portal offers it and the program supported it (Headquarter -> Both box, Print with Schedules), but the
  default list carried over from 4.3 left it unticked, so no WHT return was ever downloaded unless ticked by hand.

### Fixed: `_route_check.json` was written inside `_months\<Tax head>\` (since 10.9)
- It belongs in `Ledgers\<TIN>\` (as the notes say) and kept the `_months` folder alive. Now written there; a copy
  left by 10.9-10.11 is merged into it at the start of each tax type.

### Tests
- 175 tests (161 + 14 new); the new ones and 4 changed ones failed on 10.11 first. All pass. (In a sandbox
  without internet the 3 run-loop tests in test_window.py wait for ura.go.ug; they pass with the network check
  stubbed.) Sample workbook opened in LibreOffice without repair.

## 10.11 - 6 October 2026

### Chrome window closed by hand: the run stops cleanly (live log 6 Oct 2026, 20:36)
- Was: Returns ended with TargetClosedError and the words "Stopped - writing the status sheet"; Ledgers and
  Certificates were still started on the dead browser and each "failed" ("could not return to the portal home
  page"); the run then went on to the next client; "Could not find the Logout button".
- Now treated like Stop: one line "The browser window was closed during Returns - Ledgers, Certificates, 1 more
  client(s) not started. Run again: whatever is already saved is skipped.", then "--- Stopped: the browser window
  was closed. ---", "Job stopped." spoken. History row "Stopped: browser closed"; nothing else started.
- Detected wherever it happens: a step that fails with the browser gone, before each next step, and after a
  client's last step (which may finish without an error). core.runner.BrowserClosed (a kind of Stop) and
  browser_gone(); re-exported by core.session.
- Ledgers: no retries and no further tax types on a closed browser; the Level 1 wait (up to 90 s, three times per
  tax type) ends at once. Returns: "The browser window was closed - writing the status sheet...", or "Interrupted
  by an error - ..."; "Stopped - ..." only after Stop. Logout: "Not logged out of the URA portal: the browser
  window was closed."

### Fixed: "Starting another." was never said (10.10)
- The client counter belonged to an earlier loop, so every client got only "Job completed.". Now counted in the
  loop that runs the clients; tested through the real run loop with two clients.

### Tests
- 161 tests, all passing; new: test_browser_closed.py (4), 3 run-loop tests in test_window.py (real App.start
  with stand-in browser sessions and module work). Each failed first; the run-loop tests reproduced the live log.

## 10.10 - 6 October 2026

### Spoken announcement when a client's browser closes
- "Job completed. Starting another." between clients, "Job completed." after the last; "Job completed, please
  check the errors." when a client ended with errors; "Job stopped." after the Stop button.
- Voice = the user's choice from Voice_Samples.ps1: Sample 1 (default) or Sample 4 (1st / 2nd English Windows
  voice, normal speed), or Off - Settings > Runs, with a Test button. Windows' own speech (core/voice.py),
  runs in the background, never stops a run.

### Window fits every screen size and Windows display scaling
- Was a fixed 1240x820 (never smaller than 1000x700): too tall for 1366x768 and 1280x720, the Run button not
  visible at 1920x1080 / 150 %. Now: Windows told the program is DPI-aware (sharp, correctly sized text);
  window sized and centred inside the visible screen (maximised when smaller than wanted); the main area
  scrolls only when the screen is short; on windows narrower than about 1280 px the Run/Stop buttons move under
  the client list and a sideways scroll bar reaches wide option panels; pixel sizes scale with Windows; the
  left 'This run' list shows at most 12 rows (counts for the rest). core/display.py.

### Bulk client upload from Excel
- Clients > 'Excel template' and 'Import from Excel' (columns Client name | Login ID (TIN) | Password, found by
  heading). Same checks as 'Add client'; rows with problems listed by Excel row number; a Login ID already in
  the list can be updated or left; passwords never shown in messages and saved encrypted (core/client_import.py).

### Back-to-back capacity (measured)
- 40 clients through the real session code: memory +5 MB in total, Chrome processes not accumulating, the same
  time per client from the 2nd to the 40th; the activity log box adds 400 lines in about 50 ms at client 2 and
  at client 50. No limit in the code. The practical limit is the security code typed at each client's login.

### Tests
- 154 tests, all passing; new: test_client_import.py (4), test_voice.py (5), test_screen_sizes.py (7, real
  window in virtual screens 1024x768 to 2560x1440 and 150 %), 3 in test_window.py. Each failed first.

## 10.9 - 6 October 2026 (after the first live run of 10.8)

### Live run of 10.8 (one client, 12 months, all modules): 10 min 30 s in total
- Option C proved on the live portal: 3 of 3 route checks identical; 24 months in the background, median
  1 s to open and 1 s to save Level 3. Ledgers step 3 min 7 s for 27 Level 3 months.
- WHT calendar fixed live: every month "Period set to 01/MM/YYYY"; 24 months in 2 min 25 s.

### One workbook per tax head, one table, updated in place (user request)
- `Ledgers\<TIN>\<Tax head>.xlsx`, sheet "Level 3-2", ONE Excel table: Period, Tax Position as at, then the
  portal's Level 3-2 columns; oldest month first; amounts as numbers, PRNs as text; frozen heading, filter.
- Every run updates the same file: months fetched in the run replace their rows, others are kept
  (engines/ledger_taxhead.py). Written once at the end of the Ledgers step (also after Stop or an error):
  0.25 s for three tax heads x 12 months.
- The portal's monthly files now go to `Ledgers\<TIN>\_months\<Tax head>\` (resume, route proof, evidence);
  files saved by earlier versions are moved there automatically and are not fetched again.
- The period-named `Ledgers_<TIN>_<from>_to_<to>.xlsx` is no longer created. Level 2 files unchanged.
  Yearly tax heads (Income Tax, Rental) are not part of this change yet.
- A tax-head file open in Excel cannot be replaced: the update is saved beside it as
  "<Tax head> (updated <date>).xlsx" and the run reports 'check'.

### Option B check switched off
- Live: 0 of 27 screen copies identical (the screen box holds 4 tables, the portal's file 1). With option C the
  portal's file takes 1 s, so the check (about 1 s per month) is off: `LEVEL3_SCREEN_CHECK = False`.

### Tests
- 135 tests, all passing; new: tests/auto/test_ledger_tax_head_file.py (5, incl. the real monthly routine on
  the test copy of the live Ledger form). Each failed first.

## 10.8 - 6 October 2026

### Option C: Level 3 without going back to Level 2 every month
- From the user's capture of the live portal: the Payment link is `<a onclick="openLevelThreeTwo('AUG-2026')">`;
  `openLevelThreeTwo()` points the POST form `levelTwoOfAR` at `loadLevelThreeTwo?period=<month>` and submits
  it. The answer replaced Level 2, so every month needed Back = the portal rebuilding Level 2 (median 46-68 s
  per month in the 5 Oct logs).
- Now the link's OWN onclick runs with the form's target set to a hidden frame (`download_level3_hidden`):
  the portal receives exactly what a click sends; Level 2 stays on screen; no Back.
- Proof built into every run (`level3_month`): per tax type the first month is saved by the click route
  (today's file), then fetched in the background into a scratch folder and compared cell by cell except
  the "Tax Position as at" download-time line (`_route_check.json`). Identical -> the other months in the
  background. Different, or the portal refuses a background request (e.g. a security token that changes)
  -> Level 2 is reloaded and the rest of that tax type uses the click route. No month is lost.
- Switches: `HIDDEN_L3`, `HIDDEN_L3_VERIFY` in engines/ledger_engine.py. The log shows the route and the
  times for every month; the end summary gives medians per route.

### Fixed: 'Open portal (manual)' - nothing clickable after login, no second frame (10.6 and 10.7)
- Cause, reproduced: with 'Load portal pages faster' ticked, every request waits for the program's filter,
  which only answers while the program works with the browser; a pure wait for the person held every
  request. Manual mode now switches the filter off for the person's browser, and every wait for the person
  (manual mode, `hand_over`, the date question) keeps the browser serviced (`BROWSER_KEEPALIVE`).

### Tests
- 129 tests (122 + 7 new), all passing; each new test failed first. Test copy of the live Ledger form:
  tests/mock_portal/ledger_server.py (same field names and function, POST, session or one-use token, slow
  Level 2 rebuild, changing time line). Click-route Level 3 files identical to 10.5.
- Found while building: Playwright still lists a removed frame (detached) - the hidden frame is taken only
  when live.

## 10.7 - 6 October 2026

### Universal rule: a pop-up that may not appear is never waited for (user, 6 Oct 2026)
- Applies from login to logout, in every module: a pop-up, notice or banner is checked at once and closed
  only if it is showing. Inventory of every pop-up routine: only two broke the rule.
- `dismiss_pre_login_popup` waited 5.00 s (measured) whenever the notice was absent - at login, at the Tax
  Statement and after EVERY Level 3 month. Now checked at once. In the 5 Oct logs the notice was closed 10
  times and was already showing each time, so the same notices are closed. A notice that appears later is
  still closed by `clear_overlays` before the next clicks.
- Test Login pressed the first Logout link and waited 3 s when it was hidden: now the first visible one
  (`core.session.press_logout_link_if_shown`).
- Already following the rule (tested): cookie banner (0.04 s), `clear_overlays`, WHT `dismiss_popups` /
  `_quick_dismiss`, Returns `_skip_popups`.
- NOT covered, on purpose: waits for things the program needs (a Level 3 page, a download, the schedules
  button, the branch box) - checking those at once could lose data.
- Practice page, one Level 3 month: 7.4 s -> 2.4 s. About 5 s saved per Level 3 month on the live portal
  (about 3 minutes for 36 months).

### Ledger log: where each Level 3 month's time goes
- One line per month: "JUN-2026: times - Level 3 opened after N s, file saved N s later (portal export),
  back on Level 2 N s later"; at the end of the step the medians.

### Option B check: could Level 3 be read from the screen instead of the portal's Export?
- For every Level 3 month exported by the portal, the on-screen box is put through the SAME converter
  (`excel_files.html_to_xlsx`) in a scratch folder and compared cell by cell with the saved file
  (`excel_files.compare_cells`). Log line per month + `_screen_check.json` in the client's ledger folder +
  end-of-step count. The saved file is always the portal's; nothing changes until every month is identical
  on the live portal (the user's rule: faster only with 100% the same results). `LEVEL3_SCREEN_CHECK = False`
  switches the check off.

### Tests
- 122 tests (108 + 14 new): tests/auto/test_popup_rule.py, tests/auto/test_ledger_month_checks.py; new
  practice page `Portal/ledger_l2l3.html` (Level 2 -> Level 3 -> Export -> Back). The new tests failed first.
- Same results as 10.5: Level 3 files (cells, merges, files saved incl. portal_original), WHT months and Returns.

## 10.6 - 6 October 2026 (from the live runs of 5 Oct 2026 for two clients)

### Fixed - WHT certificates: no month was searched (every month failed as "portal too slow")
- **Cause, reproduced with the portal's own calendar code:** the URA date selector is "The DHTML Calendar"
  (Dynarch). Its arrows and dates react only to a mouse button going down and up on them; the script clicks
  used since 10.2 (`el.click()`) were ignored ("Stepped 60 month(s) to 'October, 2026'", about 110 s per month).
  The date was then written directly, but the calendar stayed open over the Search button; the forced Search
  click pressed that screen position, picked a date in the calendar (every wrong From date in the live
  screenshots is a Tuesday - the column under Search) and Search was never pressed.
- Calendar cells are now pressed like a mouse button (down, up, click) sent to the cell itself (`PRESS_JS`).
- Walking stops after 2 presses that do not move the month (`DEAD_PRESSES`); the log counts months really moved
  ("Moved 15 month(s) to 'July, 2025' (4 press(es))").
- The week-number column ('wk') is never taken for day 1 (a January view starts with week 1).
- The calendar is closed (its x button, else its box hidden) before a date is written and before Search.
- Search is never force-clicked (`click_any(..., force=False)`); it is not pressed when the From date on screen
  is not the month's first day; a From date that changes when Search is pressed is reported as such; a real
  timeout names the From date on screen instead of "portal too slow".
- Practice page with the real calendar: one month 2.2 s (was about 110 s and a wrong date).

### Fixed - Ledgers: one failure ended the whole step (PAYE and Local Excise Duty never downloaded)
- The Level 3-2 box is looked for in every frame (`find_level3_box`): twice the program waited 45 s in a frame that
  did not hold it while the screen showed it.
- Level 1 gets up to 90 s (x speed) to show (`LEVEL1_PATIENCE_S`, heartbeat every 30 s) instead of 10 s; in every
  live case it was showing 1 s after the failure.
- "Getting the menu back" now uses the portal-home route that worked every time in the logs (`ensure_portal_home`)
  instead of a Portal-link click that clicked nothing 5 times out of 5.
- `click_first_visible` looks at every match, not only the first (a hidden menu entry can come first).
- A failed recovery inside the retry loop no longer escapes and ends the Ledgers step; attempts 2-3 and the next
  tax types run. A tax type that is finished stays finished if the way back to Level 1 fails.

### Fixed - Returns: second page of one client's VAT returns not downloaded (6 of 12)
- A page turn may now take up to 60 s (180 s in "wait until it downloads"), with a progress line every 30 s
  (`NEXT_PAGE_WAIT_S`); before, about 12-13 s whatever the setting. Live page turns took 6-20 s.

### Tests
- 108 tests (95 + 13 new), all passing. The 13 new tests failed on 10.5 first. Old and new engines give identical
  results on every existing practice page (12 WHT months, both Generate For options, and the Returns files).
- New practice pages: `Portal/wht_dynarch.html` with the real calendar files (`Portal/jscalendar/`, LGPL,
  unchanged, checksums in SOURCE.txt), `Portal/level1_late.html`, `Portal/level3_outer.html`,
  `Portal/portal_links.html`; `history.html` (both copies) can turn pages slowly for a test.
- Why the 95 tests missed the WHT fault: the old practice calendar (`wht_slowclick.html`) reacts to plain clicks.

### Not changed (needs a decision or more evidence)
- Ledger folders per tax type, layout of the combined Level 3-2 sheet (Swaraj's points 2-4).
- PAYE certificates "No employees listed" (6 Oct): needs a manual check on the portal.
- Program carrying on after the browser window has closed; the fixed 5 s pop-up wait per Level 3 month.

## 10.5 - 4 October 2026

### Return PDFs - complete and neat (user asked whether they are as neat and machine-readable as 4.3)
- Checked with pdffonts / pdftotext / pdfinfo: return PDFs are machine-readable - real text, fonts embedded
  with Unicode maps (copy, search and data extraction work), no page is a picture, tagged PDF, A4.
- **Fixed: wide schedules lost their right-hand columns.** On a 24-column test page the PDF stopped at column 16.
  The page is now measured with print styles at A4 width: a page that fits prints exactly as before (A4
  portrait, full size); a wider one prints landscape and, if still too wide, scaled down just enough to fit
  (text stays text). Test page: all 24 columns, landscape, 66 %.
- **Fixed: the portal's Print / Print with Schedules / Back / Close buttons were printed** into the PDF. They are
  hidden in print only (they carry no data).
- Not verifiable here: the old 4.3 output itself (it needed Selenium and Chrome's kiosk printing). 4.3 used
  Chrome's default "Save as PDF" settings and did not unbox wide tables before printing.

## 10.4 - 4 October 2026

### Changed - Returns and 'Print with Schedules'
- The PDF is now made **when the portal asks to print** (it does that once its view is complete), instead of a
  fixed ~5.5 s after the click. The old Selenium program 4.3 printed at that same moment. The program had always
  recorded the portal's print request (`window.__botPrint`) but never used it.
- If the portal does not ask to print, the PDF is made once the page has been quiet - nothing loading, nothing
  changing - for 5.5 s (the old fixed wait), so it is never slower or less complete than before; a page still
  working stops at the page limit (60 s x speed; at most 5 min in "wait until it downloads") and is flagged.
- **Monthly VAT Return: URA no longer gives VAT schedules** (confirmed by the user, 4 Oct 2026). 'Print with
  Schedules' is still pressed (if URA brings them back they are captured), and a VAT PDF without schedules is not
  flagged. Other return types without schedules are flagged "check". The list is `NO_SCHEDULES_FROM_URA` in
  engines/returns_pw.py.
- The summary says what was done ("Print with Schedules - printed when the portal asked to print"), never
  "with schedules", which claimed content the program cannot see.
- Mock portal Returns run: 34.2 s -> 22.7 s, same files and statuses. Returns that are ready at once (like PAYE):
  about 1 s instead of 5.5 s each.

## 10.3 - 4 October 2026

### Added
- **One combined Ledgers workbook per client**, written at the end of each Ledgers step:
  `Ledgers\<TIN>\Ledgers_<TIN>_<from>_to_<to>.xlsx` with a **Contents** tab (every tax type and month of the
  period with its status: included and row count, nil payment, not in ledger, failed, not downloaded), then for
  each tax type a **"<tax> - L2"** tab (Level 2 exactly as saved, merged headers kept) and a **"<tax> - All L3"**
  tab (every Level 3 month stacked, with Month and Source file columns; a month whose columns differ gets its own
  header row; yearly Income Tax / Rental items listed by item). Built from the files on disk, so it also holds
  months downloaded in earlier runs. The separate Level 2 / Level 3 files are kept - they decide what is already
  downloaded. No portal time: a full year for 4 tax types (960 Level 3 rows) builds in about 0.3 s.

## 10.2 - 4 October 2026 (from the first live run of 10.1, log run_20261003_235216)

### Fixed
- **Ledger Excel files that Excel refused to open** ("file format or extension is not valid") - every file saved
  from the portal's Export To Excel button, since version 10. The portal sends an HTML table named `.xls`; the
  program saved it as `.xlsx`. Downloads are now checked by their contents and converted to real Excel workbooks
  with exactly the portal's text (only amounts become numbers; TINs, PRNs, return numbers, dates and codes stay
  text; merged headers kept; every cell read back and compared before the file is kept). The portal's own file
  is kept in `portal_original\` as `.xls`. **Files saved by earlier versions are repaired automatically** at
  start-up and at the start of each client's Ledgers step, keeping their dates. A ledger month counts as
  downloaded only if a real workbook is on disk.
- **Logout really logs out.** The portal asks "You have requested to be logged out, do you want to proceed?".
  10.1 answered Cancel (its question policy) and still logged "Logged out"; version 10 also answered Cancel
  whenever the Certificates step had not run (Playwright's default). The question is now answered OK while
  logging out, and "Logged out" is written only when the Logout button has gone; otherwise a warning.
- Run summary log line: missing closing bracket.

### Faster, same results
- **WHT calendar** (the slowest part of the live run): the live portal ignored mouse clicks on the calendar, so
  every arrow and the day-1 click waited 4 s before a script click did the work; opening the calendar took
  6-11 s. Script clicks are now used directly, and each one is still verified (the calendar heading after
  every arrow, the From date after day 1, the To date). The live log shows about 15-25 s of each month's
  ~30 s was this waiting.

### Diagnostics
- When a return's schedules cannot be confirmed, the log adds one line of numbers only (visible text before and
  after, number of frames) so the check can be tuned without sharing client data.

## 10.1 - 3 October 2026 (audit fixes and speed work)

Finding codes (H1, M4, ...) refer to the audit report of 3 October 2026.

### Security and data safety
- **Passwords are never written to logs, history or summaries (H6).** A failed login used to copy the password into
  its error message, which then reached the log file, `history.json` and both summary workbooks. Login errors are now
  re-raised without that text, and every log line, summary row and history entry masks known passwords as `***`.
- **One client's password is never used for another client (H2).** The login falls back to saved credentials only
  for the same Login ID, and they are cleared after each client.
- **Saved passwords are never lost (H1).** A password this Windows account cannot decrypt is kept unchanged until it
  is typed again; a run will not start for that client. A damaged client list is kept as
  `clients.damaged-<date>.json` instead of being overwritten, and `clients.json.bak` is kept before every save.
- **Generate TCC is opt-in (H7).** "Generate missing TCC numbers" (WHT and PAYE) is unticked by default. Portal
  questions are answered OK only for the generate-certificate question while Generate is being pressed; any other
  question is answered Cancel and logged.
- **Login ID** must be digits only and cannot be added twice (L1). The output folder name is always made safe.
- **Excel/CSV**: text copied from the portal is always stored as text, never as a live formula (L2).
- **Log folder clean-up**: screenshots, page copies and logs older than 30 days are deleted at start-up (Settings,
  0 = never). Error screenshots carry the client's TIN and an error while taking one no longer stops the run (M1, M2).
- **Clients > Check old files for passwords** finds and blanks passwords that earlier versions may have written
  into logs, history or summaries (shows file and client names only).
- Unused second password store removed from `engines/wht_outputs.py` (L3).
- Mock PAYE page: real-looking employee names, TINs and certificate numbers replaced with invented ones (H4).

### Nothing marked done unless it was downloaded (H8)
- Ledgers: a month not yet on the ledger, or with no payment link yet, is checked again on a later day.
- Ledgers: a Level 2 ledger saved on an earlier day is fetched again; the old copy goes to `previous`
  (option on the Ledgers page).
- Returns: an empty PDF is deleted and retried; files are written to `.part` and renamed, so a stopped run never
  leaves half a file that later counts as saved.
- Items a person must check (ledger built from the screen, Level 3 page not naming its month, schedules not
  confirmed) appear on the Review page and in the Failed sheet (M6). Run status reads "Completed - N item(s) to check".
- Amounts written as (500), UGX 60,000 or 60,000/= are counted in WHT totals (L6).

### Speed (same results - checked by tests and by comparing every output old vs new)
- Returns: both print buttons are looked for in one pass every 0.25 s instead of 1 s each in turn
  (mock run 52.1 s -> 47.7 s). The window for "Print with Schedules" to appear is never shorter than before.
- WHT: three fixed calendar pauses stop as soon as their check succeeds (mock runs 20-29 % faster).
- No fix adds waiting to a normal run. Only re-running a period on a later day re-checks unposted months and
  refreshes Level 2 ledgers (the latter can be switched off).

### Dependencies (M4)
- Exact versions in `requirements.txt`; `pyinstaller` moved to `requirements-build.txt`; `pytest` in
  `requirements-dev.txt`. pypdf raised 5.9.0 -> 6.19.0 and lxml 6.0.2 -> 6.1.3 because pip-audit reported published
  security flaws in the older versions. `pip-audit -r requirements.txt`: no known vulnerabilities.

### Tests (H5)
- `tests/auto/`: 95 pytest tests with real checks; the mock portal is served automatically. Run `python -m pytest`.

### Not changed (needs a decision or a live run)
- The client briefing (`tests/URA_Automation_Client_Briefing.docx`) - see BRIEFING_CORRECTIONS.md (H3).
- 146 places where an error is silently ignored are not yet reviewed one by one (M3).
- Fixed pauses in the ledger Level 3 loop and PAYE month-wise details (need one timed live run).
