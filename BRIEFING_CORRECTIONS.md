# Client briefing - statements that do not match the program (audit finding H3)

The document `tests/URA_Automation_Client_Briefing.docx` (dated 22 September 2026) was NOT changed, because it is a
statement to clients and its wording is the firm's decision. Below: each statement, what version 10.1 actually does,
and suggested replacement wording.

| Briefing says | What the program does | Suggested wording |
|---|---|---|
| Step 4: the tool reads registration details to decide which tax types apply | Not implemented. The person ticks the tax types and return types. | "Our team selects the tax types and returns that apply to you." |
| Step 8: it collects payment reference numbers and reprints each payment slip | Not implemented in version 10 (the old code for this is present but never runs). | Remove step 8, or mark it "planned". |
| Step 9: it compares downloads against a checklist and confirms payment totals match payment slips | Not implemented. The program lists every failed or doubtful item on its Review page and in a summary workbook. | "The tool lists anything it could not download or could not confirm, and our team reviews that list." |
| §6 and Q&A: it only views and downloads; there is no way for it to take an action that affects your account | It can generate missing tax credit certificates on the URA portal. Since 10.1 this happens only when the team ticks that option for a run. | "It does not file returns or make payments. If we ask it to, it generates missing tax credit certificates on the URA portal - only when we have agreed this with you." |
| §6: it never sends your information anywhere outside our own computer | The program itself sends nothing except to the URA portal. Its manual mode can save screens (with client data) for troubleshooting, and the team may choose to share them with an outside service (for example an AI assistant). | Either stop sharing such captures outside the firm, or say: "For troubleshooting we may share screen captures with [named service] under [terms]." |
| "It never guesses ... it stops completely and waits" | It carries on after many uncertain steps and lists what needs checking (Review page). | "When something looks unexpected, the tool records it for our team to check before your file is finalised." |
| §7: the password is never written into a plain file or a spreadsheet | True since 10.1 (before 10.1 a failed login could write it into the log, history and summary files). | Keep. If earlier versions were used, search old log/history/summary files for passwords and change any found. |

Whether sharing client data with an outside service is permitted for a given client is a legal/compliance question
(for example under Uganda's Data Protection and Privacy Act, 2019); this review did not assess it.
