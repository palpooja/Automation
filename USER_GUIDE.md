# URA Suite — User Guide

One program that downloads from the URA (Uganda Revenue Authority) e-Tax portal for many clients in one run:

| Menu item | What it fetches | Where it saves |
|---|---|---|
| **Returns** | Every filed return in a period as PDF (with schedules when the portal offers them); Local Excise Duty returns as a `.mhtml` web copy; plus `Returns_Status.xlsx` listing every return and its status | `<output folder>\Returns\<TIN>\<return type>\` |
| **Ledgers & Certificates** | Tax-statement ledgers at Level 2 (monthly summary) and Level 3 (transaction detail) for VAT, IT-WHT, PAYE, Local Excise Duty, Income Tax, Rental; WHT and PAYE tax-credit certificates; taxpayer profile; TIN/VAT certificate; PRN payment slips | `<output folder>\Ledgers\<TIN>\` |
| **WHT Receivable** | The "Generate Tax Credit Certificate for Withholding" grid for each month: Excel (Summary + Combined + one sheet per month), PDF table, page copies (PDF/HTML), optional CSV | `<output folder>\WHT_Receivable\<TIN>\` |

You log in **once per client**; every ticked module then runs on that login.

## 1. Requirements
- Windows 10/11, Google Chrome or Microsoft Edge installed.
- The portal needs a human to type the **security code** (CAPTCHA — the picture code) at each login. The suite cannot and does not do this for you.

## 2. First start
1. Install (`URA_Suite_Setup.exe`) or unzip and run `URA_Suite.exe` (or `run.bat` from source).
2. **Clients** (left menu) → **Add**: client name, Login ID (TIN — Tax Identification Number), password. Passwords are encrypted with your Windows account (DPAPI — Data Protection API): only you, on this PC, can read them back. If you used the old Returns Downloader, its client list is imported automatically on first start.
   The Login ID must be digits only, and each Login ID can be added once. If a saved password cannot be decrypted (for example the client list was copied from another PC or another Windows account), the program says so at start-up, marks the client "password must be typed again" in Clients, keeps the old encrypted text untouched, and will not start a run for that client until you type the password again (Clients > Edit). A copy of the client list before each change is kept as `clients.json.bak`.
3. **Output folder** (left menu): choose where everything is saved. Default: `Downloads\URA_Suite`.
4. **Settings**: leave the defaults unless you are behind a company proxy (fill in Proxy) or the portal is slow (raise "Wait for a page element").

## 3. Running
1. Tick the modules you want in the left menu (the box beside each name).
2. Click a module name to open its page and set its options (period, tax types, files to save).
3. In the client list, tick the clients to include (click the box in the first column).
4. Press **Run selected modules** (all ticked modules, all ticked clients) or **Run '…' only** (just the open page).
5. When the **yellow bar** appears: look at the browser, type the security code, then press **Log In** in the bar (or Enter, or click Log In in the browser yourself). The bar names the client it is asking for.
6. Watch the **Activity** log. Green = saved / logged in, amber = warning or skipped, red = error.
7. **Stop** ends the run after the current step and logs out.

### Module options
**Returns**: period as dd/mm/yyyy; tick return types (Select all / Clear all). "Skip returns already saved" makes re-runs fast.
"Wait until each return downloads" is **ticked by default** (as in the original Returns Downloader): each return may take up to 15 minutes and a failed one is retried 3 times, so nothing is skipped just because the portal is slow. The log prints "still waiting… N s" every 30 seconds so you can see it is working. The cost: a return that is genuinely broken on the portal can hold the run for up to an hour before it is recorded as failed. Untick it for normal waits (60 seconds, 1 retry) when you want speed; anything that fails is listed in the Failed sheet and a re-run picks it up.

**Ledgers & Certificates**: period as year-month (e.g. 2025-07); "Last 18 months" fills it.
**PAYE (Pay As You Earn) tax credit certificates**: the portal asks for a financial year (July to June). The program
works out every financial year your period touches and searches each one - e.g. 2025-05 to 2026-06 searches
"July 2024 to June 2025" and "July 2025 to June 2026" (shown under the tick box as you type). For each year it can
generate missing TCC numbers (then searches again to read them) and open every employee's "Get Month-Wise Details".
One Excel workbook per year: `Ledgers\<TIN>\PAYE_certificates\PAYE_TCC_<TIN>_<year>.xlsx`, sheet *Employees* (with TCC
numbers) and sheet *Month-wise details* (every employee's months, with the employee's TIN and name on each line). Tick the ledgers and other documents. "Skip Level 3 for months with nil payment" saves one page per nil month. The engine itself skips files that already exist, so a run can be resumed.

**WHT Receivable**: months as year-month, e.g. 2026-07 (defaults to the current Ugandan financial year, July–June). Under *Generate for* tick **Payee (Deductee)** — tax others withheld from this client — **Payer (Deductor)** — tax this client withheld from others — or both. Each ticked type is fetched separately and saved with the type in the file name, e.g. `WHT_<TIN>_Payee_202607-202706.xlsx` and `WHT_<TIN>_Payer_...`. "Carry on if a month fails" moves on to the next month after an error.
**Generate missing TCC (Tax Credit Certificate) numbers** (**not** ticked by default - tick it only for clients who have
agreed, because it creates certificates on the client's URA account): for each month, the rows that have no
TCC number are ticked, *Generate TCC* is pressed (the portal shows a summary, no PDF), *Back* is pressed and the month
is searched again, so the Excel file contains the new TCC numbers. Rows that already have a number are left alone.
If a number cannot be generated, the month's data is still saved and the reason is listed in the summary.
The portal's question "Are you sure you want to generate Tax Credit Certificate?" is answered OK only while the program
is pressing Generate. Any other question the portal asks is answered **Cancel** and written to the log. The PAYE
"generate missing TCC numbers" option works the same way and is also off by default.

## 4. After the run
**Ledgers - one workbook per tax type (version 10.12), laid out like the WHT workbook:**
`Ledgers\<TIN>\<tax type>.xlsx` (e.g. `Value Added Tax.xlsx`, `IT-WHT.xlsx`, `PAYE.xlsx`) with
- a *Summary* tab: one line per month - status (Level 3 saved, nil payment - skipped, not in ledger, no payment
  link, failed), the Level 2 payment, the number of Level 3 lines and their total amount, when it was checked;
- a *Level 3-2* tab: every month's Level 3 lines in one Excel table (filter and sort it like any table);
- one tab per month (JUL-2025, AUG-2025, ...): that month's Level 3 exactly as the portal exported it.

The same file is updated by every run: months fetched again replace their lines and tab, the others are kept.
There is no separate Excel file per month any more; tick "Also keep a separate Excel file for every Level 3
month" on the Ledgers page if you want them (they go to `_months\<tax type>\`). The Level 2 ledger stays as
`<tax type>_L2.xlsx` beside the workbook. If the workbook is open in Excel during a run, the update is saved as
`<tax type> (updated <date>).xlsx` and the monthly files are kept until a later run can update the workbook.

**Returns (version 10.13):** `Returns\<TIN>\<return type>\` holds `1. July 2025.pdf` ... `12. June 2026.pdf`
(the number is the month's place in the period you chose), an amended return as
`8. February 2026 - Amended filed 10-06-2026.pdf`, a branch return with ` - <branch>` added, and yearly returns as
`Original Prov 2025-26.pdf`. Returns_Status.xlsx shows each return's number and the file it was saved as
('Saved as'). A return type with nothing filed in the period gets `No Records Found 2025-26.png`, a picture of the
portal's search screen.

**Ledger check (10.15):** on each tax type's Summary tab, 'Check' compares the month's (or year's) Level 2
Payment - all its columns added up (for VAT: PRN, WHT Credit and Offset Received; for Income Tax: PRN and Tax
Credit) - with the Level 3-2 lines, each payment or credit counted once even when the portal lists it twice:
OK, or 'check: ... differs by ...' (these are also on the Review page). 'Not allocated' shows the part of a payment
the portal has not yet set against tax, penalty or interest - for information only.

**Income Tax and Rental (10.15):** with "also download their Level 3" ticked, `Income Tax.xlsx` and `Rental.xlsx`
are laid out like the monthly ones, with one row and one tab per year (JUL-2025 to JUN-2026). A year with nothing
paid is not opened.

**WHT check (10.13, 10.15):** the Payee workbook's 'Check vs ledger' tab shows whether each certificate appears in
the Income Tax ledger (found by the agent's return number, else the TCC number, else the amount), which ledger WHT
credits have no certificate, and - for information - other credits such as ASYCUDA (customs). It needs Income Tax
and its Level 3 ticked on the Ledgers page, run before Certificates.

**WHT (Withholding Tax) - one workbook per type and period:** `Certificates\<TIN>\WHT\WHT_<TIN>_<Payee|Payer>_<from>-<to>.xlsx`
with a *Summary* tab (each month, rows, totals), a *Combined* tab (all months) and a tab for each month with
data. The `.html` files in `pages\` are only copies of the portal screens; untick "Portal page copies (HTML)" on
the Certificates page if you do not want them.

Two Excel files appear in the output folder:
- `Run_Summary_<date_time>.xlsx` — this run: **Overview** (client × module), **Files** (every file produced, with a clickable link), **Failed** (every item that failed, came back incomplete, or is marked **check** - for example a ledger rebuilt from the screen, or a return whose schedules could not be confirmed), **Other outcomes** (no records, skipped). Red rows = needs a person, amber = skipped/stopped. The **Review** page shows the same items.
- Text copied from the portal (names, references) is always stored as text, never as an Excel formula.
- `URA_Suite_Summary.xlsx` — the same rows appended after every run, with a Run column: the history across all runs.

Also: **History** in the left menu, and a text log per run in `%LOCALAPPDATA%\URA_Suite\logs\`.

## 5. Resume after a problem
With **Resume** ticked in Settings (default), a client + module pair already completed today is skipped on the next run, so a batch that stopped at client 12 restarts from client 13. Within a client, Returns and Ledgers also skip files already on disk.
- A return PDF that came back empty is deleted and tried again on the next run (it no longer counts as saved).
- Ledgers: the portal's "Export To Excel" sends a web-page table, not a real Excel file. The program converts it
  to a real `.xlsx` with the same text and keeps the portal's own file in `portal_original\` beside it.
- Ledgers: a month that was not yet on the portal's ledger, or had no payment link yet, is looked at again on a **later day** (on the same day it is not, so Resume stays fast).
- Ledgers: a Level 2 ledger saved on an earlier day is fetched again, because a ledger changes as payments and assessments are posted. The older copy is moved to the `previous` folder beside it. Untick "Fetch a fresh Level 2 ledger…" on the Ledgers page to keep the old behaviour.

## 6. What "slow site" does
Each item gets a time limit, then a retry; only after the retries fail is that one item skipped and written to the Failed sheet. Nothing is dropped silently. Returns additionally compare the portal's "Total Records" count with the files on disk and flag the category as "incomplete" if they differ.

## 7. Common messages
| Message | Meaning / what to do |
|---|---|
| "No browser could be started" | Install Google Chrome or Microsoft Edge |
| "Could not load the login page" | Open ura.go.ug in normal Chrome; if it works there, set the Proxy in Settings |
| "Login was rejected / Refilling your details" | Wrong security code — type it again |
| "The output folder cannot be written to" | Pick another folder (not inside Program Files) |
| "already completed today" (amber) | Resume skipped it; untick Resume in Settings to force a re-run |
| "portal reported N record(s) but only M on disk" | Re-run Returns for that client; the missing ones are retried |
| "Passwords needed" / "must be typed again" | The saved password cannot be decrypted on this Windows account: Clients > Edit, type it again |
| "The saved client list could not be read" | The file was damaged; it was kept as `clients.damaged-<date>.json` and `clients.json.bak` holds the copy before the last change |
| "The portal asked '…' - answered Cancel" | An unexpected portal question; check what it was before re-running |
| "check: 'Print with Schedules' pressed, but the portal neither asked to print…" | Open the PDF and check the schedules are there (not shown for Monthly VAT Returns: URA no longer gives VAT schedules) |
| "still waiting for the portal to prepare the schedules" | The portal is still loading the schedules; the program waits (at most 60 s, or 5 min when "wait until each return downloads" is ticked) |
| "Repaired N ledger Excel file(s)" | Files saved before 10.2 that Excel could not open were converted; nothing to do |
| "Logout was pressed but the portal did not confirm it" | The session may stay open on the portal until it times out; nothing else to do |
| "built from the on-screen table" | The portal's Export gave no file; the Excel was made from the screen - compare it with the portal |

## 8. Where things live
- Clients, settings, history, logs: `%LOCALAPPDATA%\URA_Suite\`
- Downloads: the output folder you chose.
- The program sends nothing anywhere except to the URA portal. Files that **you** choose to share (logs, screenshots, page captures) contain client data - see sections 9 and 10.
- Passwords are never written to the logs, the history or the summary workbooks (they are replaced by `***`).
  Versions before 10.1 could write one there after a failed login: **Clients > Check old files for passwords** finds
  such files (it shows file and client names, never the password) and can replace the passwords with `***`. If any
  of those files was ever sent to someone, change that client's password.
- Screenshots, page copies and logs in the log folder are deleted after 30 days (Settings: "Delete screenshots, page copies and logs older than", 0 = never). Error screenshots start with the client's TIN.

## 9. Running several modules for one client
After each module the program returns the portal to its home page before starting the next one (log line:
"Returning to the portal home page… / Portal home page is open."). It stays logged in. If it ever cannot get back,
the log names the screen it was on and a screenshot `home_failed_<TIN>.png` is saved in the log folder — send both
to whoever maintains the program, remembering that the screenshot shows the client's portal screen.

## 10. Open portal (manual) - and recording your steps
Left menu > **Open portal (manual)**: logs in as the client selected in the list, then leaves the browser to you. The
program clicks and downloads nothing; the browser stays open until you press **Close browser** in the yellow bar.
- **Save this page for Claude**: saves what the browser shows (each frame's HTML + a screenshot). On closing, all saved
  screens are zipped into the log folder as `page_captures_<TIN>_<date>.zip`. It contains the client data on screen,
  but no password. **Sharing it outside the firm (for example with an AI service) sends that client data outside the
  firm** - check this is allowed for that client first.
- **Recording with Chrome**: in that browser press F12 > the ">>" menu > **Recorder** > Create a new recording > do your
  steps > End recording > Export > JSON. Because the program already logged in, the recording has no password.
  (In your own Chrome, start the recording AFTER logging in, or delete the password from the JSON before sharing.)

## 11. "Already completed today"
When Resume is on and something ticked was completed today without failures, the program lists it and asks:
Yes = skip it, No = run it again, Cancel = do not start.

## 12. Layout from version 10 (follows the client briefing)
- **Period** (year-month, from / to) sits above the client list and applies to every step. "This financial year" and
  "Last 18 months" fill it. Returns use the first day of the From month to the last day of the To month; PAYE
  works out the July-June financial years itself.
- Left menu = the steps for each client: **1 Returns, 2 Ledgers, 3 Certificates (WHT and PAYE), 4 Review**. Tick the
  steps to run; click a step's name to see its options.
- **Review** opens by itself when a run ends: each step's result per client, and "Needs your review" - every failed or
  incomplete item with its reason. Buttons open the run summary (Excel) and the output folder.
- Output: `Returns\<TIN>\...`, `Ledgers\<TIN>\...`, `Certificates\<TIN>\WHT\...`, `Certificates\<TIN>\PAYE\...`.
- No PDFs are made for WHT or PAYE (Excel and portal page copies only).
- Settings > **Internet speed**: Normal / Slow (every wait x2) / Very slow (x4). A dropped connection is always waited
  out before a retry ("The connection to ura.go.ug is down - waiting...").
- Return file names now use fixed facts: `<return number> <from> to <to> filed <date>.pdf`. Files saved under the old
  names are recognised by their return number, so nothing is downloaded twice.

## 13. Adding many clients from Excel (version 10.10)
Clients > **Excel template** saves a workbook with the columns Client name, Login ID (TIN) and Password.
Fill one client per row, then Clients > **Import from Excel**. The program checks every row like
"Add client" and lists any row it cannot use by its Excel row number. A Login ID already in the list can be
updated (new name or password) or left as it is. Passwords are saved encrypted with your Windows account.
Afterwards delete the Excel file, or keep it somewhere safe: it holds the passwords in plain text.

## 14. "Job completed" announcement and screen sizes (version 10.10)
When a client's browser closes the program says "Job completed. Starting another." - or "Job completed."
after the last client. Choose the voice (Sample 1, Sample 4 or Off) under Settings > Runs and press Test.
The window adapts to the screen and to Windows display scaling; on a short screen a scroll bar appears on
the right. On screens narrower than about 1280 pixels the Run and Stop buttons move under the client list.

## 15. How many clients in one run
There is no limit in the program; memory does not build up from client to client. Each client still needs
the security code typed at its log-in, so stay within reach of the computer: the voice tells you when the next
client starts. If nobody types the code within the waiting time (Settings, default 7 minutes) that client is
marked failed and the run moves on to the next one.

## 17. A return that will not open (version 10.14)
Some returns open slowly on the portal (old VAT returns can take a minute) and now and then one never opens.
The program gives each return up to 2 minutes the first time. A return that is not ready by then is set aside,
and everything else for that client is done first. At the end of the client the set-aside returns are tried again
with the full wait. If one is still not open after a minute, you hear "Attention needed" and the yellow bar asks
**Skip this return** or **Keep waiting**. If you are not there, it simply keeps waiting. A skipped return is listed
on the Review page and is fetched on the next run. Please do not close Chrome to stop a wait - use Skip.

## 18. Second try at the end of a client (version 10.15)
Anything that failed for a client is tried once more after that client's other steps - a whole step that stopped
with an error, a ledger month or year, a WHT certificate type, a PAYE year, a return not downloaded. Work already
saved is skipped, so only the missing part is done. The Review page then says e.g. "saved on second try"; if it
still fails it stays "failed" with the reason, and the next run tries it again.
For PAYE certificates: if the Search press does not reach the portal, it is pressed again after 8 seconds; if the
portal does not answer within a minute, a copy of the page is saved in the log folder and the year is searched once
more from the portal home page.

## 19. How Level 3 is fetched (version 10.16)
On the Ledgers page, "How Level 3 is fetched":
- **Auto (recommended)** - the first month fetched for a client goes the usual way (click, Export To Excel) and
  is also asked of URA directly; the two are compared. If they match, every other month of that client - its other
  tax types too (version 10.17; until 10.16 each tax type was proved on its own) - is asked of URA directly: no
  screen, no Back, so each month takes only as long as URA takes to answer. If they differ, that tax type stays
  on the usual way and the Review page shows what differed; a clear 'no' from URA on a later tax type switches
  that tax type alone to the usual way.
- **Direct request first, no proof** - the direct way from the first month (the usual way takes over if URA says no).
- **Screen way only** - exactly as version 10.15.
Lines of one payment that URA lists in another order (same date) count as the same; any other difference does not.
While URA is working on a page the log shows "URA still working on it (N s)". The box "Switch to the screen way if
URA has not answered after N seconds" is 0 (wait for URA's answer) unless you set it.

## 16. Closing the Chrome window during a run (version 10.11)
Closing the Chrome window is treated like pressing Stop: the run ends at once, nothing else is started, and
the Activity box says which steps were not started. Run again later: whatever is already saved is skipped.
The Stop button is still the better way - it lets the current step finish and logs out of the URA portal.


## 20. PAYE certificates that did not search, and other fixes (version 10.17)
- If the log says "Search sent nothing to the portal", the next line now describes the button that was pressed and
  the second press goes another way. After 16 s with nothing sent the year is recorded as failed with that reason
  and is not tried a third time at the end of the client. Send `logs\paye_search_timeout_<year>.png` (every attempt
  keeps its own copy: `_2`, `_3` ...) together with that log line.
- A PAYE year whose page comes back showing another period than the one asked for is a "check" item on the Review
  page, not "no records".
- Ledgers: the first Back after a Level 3 month is no longer pressed twice while URA is still rebuilding Level 2;
  an empty ledger (for example Local Excise Duty with no periods) is saved in seconds.
- The run summary's Overview counts the same files as its Files sheet, also for a step that ended with an error.
- Logout is pressed from the portal home page when the last step ended on another screen.
