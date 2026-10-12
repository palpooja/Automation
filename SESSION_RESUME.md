# URA Suite - session resume file (saved 4 October 2026)

## How to continue in a new chat
1. Start a new chat with Claude.
2. Upload `URA_Suite_v10.14.zip` (this file is also inside it, at `ura_suite/SESSION_RESUME.md`).
3. Write: *"Read SESSION_RESUME.md, CLAUDE_CONTEXT.md, HANDOVER.md and CHANGELOG.md in the zip, then continue
   as my senior engineering partner."* Add your new run log or question.

## Working rules agreed for the assistant
- Use the senior-engineering-partner approach: read the code and logs first, find the cause before changing
  anything, write a test that fails first, then fix, then run every test (`python -m pytest`).
- User preferences: always write the full form of an abbreviation with a short explanation in brackets; no
  jargon - explain plainly; do not assume - give the source of every conclusion (file and line, test output,
  tool output, or the user's own statement); be precise, never casual.
- Speed rule set by the user: make it faster only with 100% the same results; prove it by comparing old and new
  outputs.
- Never put client data or passwords into examples, test files or summaries.

## The project in one paragraph
URA Suite is a Windows desktop program (Python, Tkinter window, Playwright browser control) that logs into the
URA (Uganda Revenue Authority) portal for each client - the person types the security code (CAPTCHA) - and
downloads: 1 Returns (return copies as PDF; Local Excise Duty as a single-file web page .mhtml), 2 Ledgers
(Level 2 and Level 3 tax statement ledgers in Excel), 3 Certificates (WHT - Withholding Tax - and PAYE - Pay As
You Earn - tax credit certificates in Excel). It writes run summaries for review.

## Current version: 10.16 (all 214 automated tests pass)
| Version | What changed |
|---|---|
| 10.1 | Audit fixes: passwords never written to logs/history/summaries; one client's password never used for another; saved passwords never lost; TCC generation opt-in; nothing marked done unless downloaded; Excel text never becomes a formula; log clean-up; Login ID checks; exact library versions (pypdf 6.19.0, lxml 6.1.3 - older ones had published security flaws); real test suite; "Check old files for passwords" button |
| 10.2 | Ledger "Excel" files from the portal's Export button (really HTML tables) saved as real .xlsx, earlier files repaired automatically (originals in `portal_original`); logout answers the portal's "do you want to proceed?" with OK and is verified; WHT calendar uses verified script clicks (3 months: 133 s -> 13 s on a live-like mock) |
| 10.3 | One combined Ledgers workbook per client: `Ledgers_<TIN>_<from>_to_<to>.xlsx` with Contents, and per tax type an "L2" tab and an "All L3" tab |
| 10.4 | Return PDFs made when the portal asks to print, else after 5.5 s of quiet page (the old fixed wait); Monthly VAT Return not flagged for missing schedules |
| 10.5 | Wide schedules keep every column (landscape / scale only when needed); portal buttons no longer printed in PDFs |
| 10.6 | Live faults of 5 Oct 2026 fixed: WHT calendar pressed like a mouse button, calendar closed before Search, no forced Search click; Ledgers no longer stop after one failure (Level 3-2 box in any frame, 90 s for Level 1, proven home-page route); Returns wait up to 60 s for a page turn |
| 10.7 | Universal rule: pop-ups checked at once, never waited for (5 s saved per Level 3 month); log times per Level 3 month; screen-copy check for a faster Level 3 (option B) |
| 10.8 | Option C: Level 3 in a hidden frame, no Back per month, proved per tax type with automatic fallback; manual mode no longer freezes |
| 10.9 | One workbook per tax head (one table, updated in place); monthly files in _months; option B check off |
| 10.10 | Spoken 'Job completed'; window fits every screen and scaling; bulk client import from Excel |
| 10.11 | Chrome closed by hand stops the run cleanly; 'Starting another.' now said between clients |
| 10.16 | Level 3 by the direct way (URA's answer to the payment link's own request, no screen, no Back), proved per tax type against the screen way, which stays as the automatic backup; Ledgers-page choice Auto / Direct / Screen only and optional switch time |
| 10.15 | Ledger check = total of Level 2 Payment sub-columns vs Level 3 amounts per reference; Income Tax / Rental Level 3 via the openLevelThreeTwo links (yearly workbook); WHT check by return number; PAYE Search pressed again / tried again / per-year failure; second try at the end of the client for anything that failed |
| 10.14 | A return slow to open is set aside after 120 s and tried again at the end of the client; 'Attention needed' + Skip / Keep waiting during that wait |
| 10.13 | Audit-file return names (`3. September 2025.pdf`, '- Amended filed ...'), 'No Records Found' screenshots, Level 2 vs Level 3 payment check, WHT certificates vs Income Tax ledger |
| 10.12 | Ledgers: no Excel per month - <Tax head>.xlsx has Summary, Level 3-2 and a tab per month (like WHT); monthly files removed once inside it |

## Decisions the user made
- (Superseded 8 Oct 2026) Separate Level 3 files per month are no longer kept: each month is a tab in
  <Tax head>.xlsx and counts as downloaded from there. Level 2 stays a separate <tax>_L2.xlsx.
- Combined Ledgers layout: one workbook per client, an L2 tab and an "All L3" tab for each tax type.
- The existing WHT workbook (Summary, Combined, one tab per month) is enough.
- URA no longer gives VAT schedules (user, 4 Oct 2026). "Print with Schedules" is still pressed for VAT so
  schedules are captured if URA brings them back.
- Local Excise Duty Return stays as .mhtml only - no PDF wanted.
- "Generate missing TCC numbers" is off by default; tick it only for clients who agreed.

## What the live run of 3 Oct 2026 showed about the portal
- Logout asks "You have requested to be logged out, do you want to proceed?" (must be answered OK).
- Mouse clicks on the WHT calendar do not land; script clicks do (each arrow cost 4 s before).
  CORRECTED 6 Oct 2026: the 5 Oct log proves script clicks do NOT move the calendar (The DHTML Calendar reacts
  to mouse button down/up only). 10.6 presses cells with mouse events sent to the cell. The 3 Oct log
  (run_20261003_235216.txt) was not available to check why it seemed to work then.
- The portal's "Export To Excel" sends an HTML table named Level-2.xls.
- "Print with Schedules" on returns shows the result on the same page (no new tab).
- After each Level 3 month the program waits up to 5 s for a pop-up that never appeared (not changed - needs more
  live evidence; `ledger_engine.py`, `dismiss_pre_login_popup`).

## Open items - check after the next live run
1. Log says "Logged out of the URA portal." with no warning.
2. WHT months much faster than ~30 s each.
3. VAT returns: about 1 s each = fine; about 6 s each = the portal's VAT button does not call print - the user
   may choose plain "Print" for VAT to save ~5 s per return (trade-off: VAT schedules would no longer be caught
   automatically if URA brings them back).
4. Open the new combined Ledgers workbook and check the Contents tab.
5. If an old 4.3 PDF with a wide schedule exists, compare it with a 10.5 PDF.
6. Not done yet: review of 146 places that silently ignore an error (audit finding M3); the 5 s pop-up wait per
   Level 3 month; corrections to the client briefing document (see BRIEFING_CORRECTIONS.md - the user's decision).

## Running the tests on a PC
```
pip install -r requirements-dev.txt
python -m playwright install chromium
python -m pytest
```

## Open after 10.6 (6 Oct 2026)
1. Live check of 10.6: WHT log should show "Moved N month(s)" and "Period set to 01/MM/YYYY" for every month, and
   results or "no records" instead of "portal too slow". Ledgers should reach every ticked tax type.
2. Swaraj's points 2-4 (separate ledger folders per tax type; layout of the combined Level 3-2 sheet; all months /
   tax types in one sheet) - need his decisions and one real Ledgers_<TIN>_... workbook.
3. PAYE certificates "No employees listed for July 2025 to June 2026" (client A, 6 Oct) conflicts with 12 PAYE
   returns filed - needs one manual check on the portal; the program should save a screenshot when it says "none".
4. Ask the person who ran client A whether they closed the browser at 12:52 and 13:04 on 5 Oct (log cannot tell); the program should stop
   cleanly when the browser has gone instead of failing the next modules.
5. Full client B log (run_20261005_181132.txt) - the pasted copy stops at 18:48:33.

## Open after 10.7 (6 Oct 2026)
6. After the next live run send the log and each client's `Ledgers\<TIN>\_screen_check.json`: if every Level 3
   month is "identical", Level 3 can be read from the screen (no Export wait); if not, the file says what differs.
7. Option C built in 10.8 - after the first live run check `Ledgers\\<TIN>\\_route_check.json` and the log's route per month.
8. Open: PAYE certificates still "No employees listed" on the live run of 6 Oct (client files PAYE monthly) -
   needs a manual check on the portal and a page copy. Level 2 in the tax-head file (second sheet) - user to decide.

9. Done in 10.11: clean stop when the Chrome window is closed by hand.
10. 10.12: PAYE certificates fixed from the user's capture of 8 Oct 2026 (any-frame 'no employees', month-wise
   pop-up window, 5-per-press generate). Check on the next live run: employees listed, month-wise lines for
   every employee, and the log line 'Employees list shown N s after Search'.
