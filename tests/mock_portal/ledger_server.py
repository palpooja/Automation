"""Test copy of the portal's Ledger screens that answer the REAL form (user's capture, 6 Oct 2026):
  <form name="levelTwoOfAR" id="levelTwoOfAR" method="post" action=".../accountReportDetailedController/loadLevelTwo">
  hidden fields: _csrf ViewLedger taxType reportCode screenWidth hParaShowString fieldsToSkip prm_TaxType
                 prm_datefrom prm_dateto prm0_reportCode lastEventTS periodFromTo periodList isDebitJVZero
                 isRefundZero isAmountOtherPeriodZero isCreditJVZero isTaxWaiverZero isInstallmentZero
                 isDisputedTaxZero
  link:          <a onclick="openLevelThreeTwo('AUG-2026')" style="cursor:pointer;">1,329,127</a>
  function:      openLevelThreeTwo(prd) - copied below word for word.
STATE (set by a test): csrf 'session' (one token for the login session) or 'rotate' (a new token on
every page, each usable once - then a second request from an unchanged Level 2 is refused, 403);
level2_delay (s) imitates the slow rebuild of Level 2 on Back.  Every Level 3 page has a different
'Tax Position as at' line, like the live portal.

kind='yearly' (10.15): the Income Tax / Rental screens as captured by the user on 8 Oct 2026 - Level 2 with TWO
heading rows (Return [Provisional, Final] ... Payment [PRN, Tax Credit] ...), every amount a link: Return,
Assessment, Penalty -> openLevelThreeOne, Interest -> openLevelThreeFive, PRN and Tax Credit -> openLevelThreeTwo
(same page), Amount Allocated -> openLevelThreeThree, Tax waiver -> openLevelThreeFour.  Level 3-2 of a year:
form#levelThreeTwoOfAR, div#tab > table#tblData, lines 'Payment (PRN)', 'WHT Credit', 'ASYCUDA Credit', the
portal's exportTableToExcel() word for word (a <button> with no type inside the form, so it also submits the
form), Back = goBack().  Every amount, PRN and return number is invented.  STATE 'fail_years': years whose
Level 3-2 request gets a server error page; 'other_requests' records any Level 3-1/3-3/3-4/3-5 request."""
import itertools
import time
from urllib.parse import parse_qs, urlparse

BASE = "/Portal/accountReportDetailedController/"
STATE = {"csrf": "session", "level2_delay": 1.5, "level3_delay": 0.4, "level2_loads": 0,
         "level3_requests": [], "valid": set()}
_counter = itertools.count(1)
MONTHS = [("JUN-2026", "1,329,127"), ("MAY-2026", "8,068,063"), ("APR-2026", "9,704,197")]
FIELDS = ("ViewLedger taxType reportCode screenWidth hParaShowString fieldsToSkip prm_TaxType prm_datefrom "
          "prm_dateto prm0_reportCode lastEventTS periodFromTo periodList isDebitJVZero isRefundZero "
          "isAmountOtherPeriodZero isCreditJVZero isTaxWaiverZero isInstallmentZero isDisputedTaxZero").split()
FUNCTION = '''function openLevelThreeTwo(prd){
	var period = prd;
	document.levelTwoOfAR.action = "/Portal/accountReportDetailedController/loadLevelThreeTwo?period="+period+"";
	document.levelTwoOfAR.submit();
	return true;
}'''


def reset(**kw):
    STATE.update({"csrf": "session", "level2_delay": 1.5, "level3_delay": 0.4, "level2_loads": 0,
                  "level3_requests": [], "valid": set(), "kind": "monthly", "fail_years": set(),
                  "other_requests": [], "level3_modes": [], "refuse_fetch": False, "flip_order": False})
    STATE.update(kw)


# ---------------- yearly (Income Tax) - invented amounts
YEARS = {   # period: (provisional, final, PRN, Tax Credit, lines)
    "JUL-2026 to JUN-2027": ("0", "0", "350,000", "460,000", [
        ("Payment (PRN)", "11-JUN-2026", "2260000000201", "100,000"),
        ("Payment (PRN)", "20-JUL-2026", "2270000000202", "250,000"),
        ("WHT Credit", "15-AUG-2026", "MTO000000000301", "60,000"),
        ("ASYCUDA Credit", "02-SEP-2026", "C 0000 2026", "400,000")]),
    "JUL-2025 to JUN-2026": ("9,000,000", "0", "7,500,000", "1,200,000", [
        ("Payment (PRN)", "30-SEP-2025", "2250000000401", "2,500,000"),
        ("Payment (PRN)", "30-DEC-2025", "2250000000402", "2,500,000"),
        ("Payment (PRN)", "30-MAR-2026", "2260000000403", "2,500,000"),
        ("WHT Credit", "10-NOV-2025", "MTO000000000501", "450,000"),
        ("WHT Credit", "12-FEB-2026", "MTO000000000502", "750,000")]),
    "JUL-2024 to JUN-2025": ("0", "0", "0", "0", []),
    "JUL-2023 to JUN-2024": ("0", "6,100,000", "6,100,000", "0", [
        ("Payment (PRN)", "30-JUN-2024", "2240000000601", "6,100,000")]),
}
YEAR_FUNCTIONS = "\n".join(f'''function openLevelThree{n}(prd){{
	var period = prd;
	document.levelTwoOfAR.action = "/Portal/accountReportDetailedController/loadLevelThree{n}?period="+period+"";
	document.levelTwoOfAR.submit();
	return true;
}}''' for n in ("One", "Two", "Three", "Four", "Five"))
EXPORT_JS = """function exportTableToExcel(tableID, filename = ''){
    var downloadLink;
    var dataType = 'application/vnd.ms-excel';
    var tableSelect = document.getElementById(tableID);
    var tableHTML = tableSelect.outerHTML.replace(/ /g, '%20');
    filename = filename?filename+'.xls':'excel_data.xls';
    downloadLink = document.createElement("a");
    document.body.appendChild(downloadLink);
    downloadLink.href = 'data:' + dataType + ', ' + tableHTML;
    downloadLink.download = filename;
    downloadLink.click();
}
function goBack(){
	document.levelThreeTwoOfAR.action = "/Portal/accountReportDetailedController/loadLevelTwo";
	document.levelThreeTwoOfAR.submit();
	return true;
}"""


def _a(fn, period, amount):
    return f'<td><a onclick="openLevelThree{fn}(\'{period}\')" style="cursor:pointer;">{amount}</a></td>'


def yearly_level2_page():
    STATE["level2_loads"] += 1
    time.sleep(STATE["level2_delay"])
    tok = _token()
    hidden = "".join(f'<input type="hidden" name="{f}" value="x">' for f in FIELDS)
    rows = ""
    for i, (period, (prov, final, prn, credit, _)) in enumerate(YEARS.items(), 1):
        rows += (f"<tr><td>{i}</td><td>{period}</td>" + _a("One", period, prov) + _a("One", period, final)
                 + _a("One", period, "0") + _a("One", period, "0") + _a("Five", period, "0")
                 + _a("Two", period, prn) + _a("Two", period, credit) + _a("Three", period, "0")
                 + _a("Four", period, "0") + "<td>0</td></tr>")
    return (f'<html><body><form name="levelTwoOfAR" id="levelTwoOfAR" method="post" action="{BASE}loadLevelTwo">'
            f'<input type="hidden" name="_csrf" value="{tok}">{hidden}'
            '<h4>Period-wise Taxpayer Ledger -&gt; Level - 2 (Period Wise Position)</h4>'
            '<table id="tblData"><tr><td colspan="13">Tax Position as at 08/10/2026 10:35:08</td></tr>'
            '<tr><th rowspan="2">S/N</th><th rowspan="2">Period</th><th colspan="2">Return</th>'
            '<th rowspan="2">Assessment</th><th rowspan="2">Penalty</th><th rowspan="2">Interest</th>'
            '<th colspan="2">Payment</th><th rowspan="2">Amount Allocated to Other Periods</th>'
            '<th rowspan="2">Tax waiver</th><th rowspan="2">Net Position</th></tr>'
            '<tr><th>Provisional</th><th>Final</th><th>PRN</th><th>Tax Credit</th></tr>'
            f'{rows}</table><button onclick="exportTableToExcel(\'tblData\', \'Level-2\')">Export To Excel</button>'
            f'</form><script>\n{YEAR_FUNCTIONS}\n{EXPORT_JS.split("function goBack")[0]}\n</script></body></html>')


def yearly_level3_page(period, body):
    STATE["level3_requests"].append(period)
    if period in STATE["fail_years"]:
        return 500, "<html><head><title>Error</title></head><body><h1>Internal Server Error</h1></body></html>"
    tok = _token()
    time.sleep(STATE["level3_delay"])
    stamp = f"08/10/2026 10:{next(_counter) % 60:02d}:{int(time.time()) % 60:02d}"
    lines = "".join(f"<tr><td>{n}</td><td>{b}</td><td>{d}</td><td>{ref}</td><td>{amt}</td><td>NA</td><td></td>"
                    "<td>0</td><td>0</td><td>0</td><td>0</td></tr>"
                    for n, (b, d, ref, amt) in enumerate(YEARS.get(period, (0, 0, 0, 0, []))[4], 1))
    heads = ("Sr No.", "Payment Basis", "Effective Date", "Payment Registration Number", "Amount",
             "Adjusted Against Period", "Liability Basis", "Principal", "Penalty", "Interest", "Refund")
    table = ('<div style="overflow-x: scroll;" id="tab"><table class="FORM_TAB_BORDER" border="1" id="tblData">'
             f'<tbody><tr><td align="center" colspan="11">Tax Position as at {stamp}</td></tr>'
             f'<tr><th colspan="11">Period: {period}</th></tr>'
             "<tr>" + "".join(f"<th>{h}</th>" for h in heads) + f"</tr>{lines}</tbody></table></div>")
    return 200, (f'<html><head><script>\n{EXPORT_JS}\n</script></head><body>'
                 '<h2 class="page-title">Period-wise Taxpayer Ledger -&gt; Level - 3-2 (Details Of Payment)</h2>'
                 '<form name="levelThreeTwoOfAR" id="levelThreeTwoOfAR" method="POST">'
                 f'<input type="hidden" name="_csrf" value="{tok}">{table}'
                 '<table><tfoot><tr><td><input type="button" name="backbtn" value="Back" class="submit" '
                 'onclick="goBack()" id="backbtn"> <button onclick="exportTableToExcel(\'tblData\', \'Level-3_2\')">'
                 'Export To Excel</button></td></tr></tfoot></table></form></body></html>')


def other_level3_page(name, period):
    STATE["other_requests"].append((name, period))
    return 200, (f'<html><body><form name="level{name}OfAR" id="level{name}OfAR" method="POST">'
                 f'<h4>Period-wise Taxpayer Ledger -&gt; Level - 3</h4><p>Period: {period}</p>'
                 '<table id="tblData"><tr><th>Sr No.</th><th>Particulars</th></tr></table>'
                 '<input type="button" id="backbtn" value="Back"></form></body></html>')


def _token():
    if STATE["csrf"] == "session":
        STATE["valid"] = {"tok-session"}
        return "tok-session"
    tok = f"tok-{next(_counter)}"
    STATE["valid"] = {tok}                      # only the newest token works, once
    return tok


def level2_page():
    STATE["level2_loads"] += 1
    time.sleep(STATE["level2_delay"])
    tok = _token()
    hidden = "".join(f'<input type="hidden" name="{f}" value="x">' for f in FIELDS)
    rows = "".join(f'<tr><td>{i}</td><td>{m}</td><td>100</td><td>0</td><td>0</td><td>0</td>'
                   f'<td><a onclick="openLevelThreeTwo(\'{m}\')" style="cursor:pointer;">{amt}</a></td><td>0</td></tr>'
                   for i, (m, amt) in enumerate(MONTHS, 1))
    return (f'<html><body><form name="levelTwoOfAR" id="levelTwoOfAR" method="post" action="{BASE}loadLevelTwo">'
            f'<input type="hidden" name="_csrf" value="{tok}">{hidden}'
            '<h4>Period-wise Taxpayer Ledger -&gt; Level - 2 (Period Wise Position)</h4>'
            '<table id="tblData"><tr><th>S/N</th><th>Period</th><th>Return</th><th>Assessment</th><th>Penalty</th>'
            f'<th>Interest</th><th>Payment</th><th>Net Position</th></tr>{rows}</table></form>'
            f'<script>\n{FUNCTION}\n</script></body></html>')


def _split_lines():
    """flip_order (10.16, live VAT FEB-2026 9 Oct 2026): one payment split over two lines with the same date;
    the portal lists them in either order - here the order changes on every request."""
    if not STATE.get("flip_order"):
        return ""
    STATE["flips"] = STATE.get("flips", 0) + 1
    a = ("Original Return", "1,000")
    b = ("Penalty Assessment", "27")
    first, second = (a, b) if STATE["flips"] % 2 else (b, a)
    return "".join(f'<tr><td>{n}</td><td>Payment (PRN)</td><td>13-MAR-2026</td><td>2260000000999</td>'
                   f'<td>1,027</td><td>{basis} {amt}</td></tr>' for n, (basis, amt) in ((2, first), (3, second)))


def level3_page(period, body):
    STATE["level3_requests"].append(period)
    sent = parse_qs(body).get("_csrf", [""])[0]
    if sent not in STATE["valid"]:
        return 403, "<html><head><title>Error</title></head><body><h1>Access Denied</h1></body></html>"
    if STATE["csrf"] == "rotate":
        STATE["valid"] = set()                  # used up
    tok = _token()
    time.sleep(STATE["level3_delay"])
    stamp = f"06/10/2026 04:{next(_counter) % 60:02d}:{int(time.time()) % 60:02d}"
    amt = dict(MONTHS).get(period, "0")
    table = (f'<table class="FORM_TAB_BORDER"><tr><td colspan="6">Tax Position as at {stamp}</td></tr>'
             f'<tr><td colspan="6">Period: {period}</td></tr>'
             '<tr><th>Sr No.</th><th>Payment Basis</th><th>Effective Date</th><th>Payment Registration Number</th>'
             '<th>Amount</th><th>Adjusted Against Period</th></tr>'
             f'<tr><td>1</td><td>Payment (PRN)</td><td>14-SEP-2026</td><td>2270004432422</td><td>{amt}</td>'
             f'<td>{period}</td></tr>{_split_lines()}</table>')
    return 200, (f'<html><body><form name="levelThreeTwoOfAR" id="levelThreeTwoOfAR" method="post" '
                 f'action="{BASE}loadLevelTwo"><input type="hidden" name="_csrf" value="{tok}">'
                 '<h4>Period-wise Taxpayer Ledger -&gt; Level - 3-2 (Details Of Payment)</h4>'
                 f'{table}<input type="button" id="backbtn" value="Back" '
                 'onclick="document.levelThreeTwoOfAR.submit()">'
                 '<input type="button" value="Export To Excel" onclick="exportFile()"></form>'
                 '<script>function exportFile(){var t=document.querySelector("table.FORM_TAB_BORDER");'
                 'var a=document.createElement("a");a.href=URL.createObjectURL(new Blob(["<html><body><table>"+'
                 't.innerHTML+"</table></body></html>"],{type:"application/vnd.ms-excel"}));'
                 'a.download="Level3.xls";document.body.appendChild(a);a.click();}</script></body></html>')


def handle(handler, method):
    """Answer the Ledger addresses; returns False for any other path."""
    url = urlparse(handler.path)
    if not url.path.startswith(BASE):
        return False
    length = int(handler.headers.get("Content-Length") or 0)
    body = handler.rfile.read(length).decode() if (method == "POST" and length) else ""
    period = parse_qs(url.query).get("period", [""])[0]
    yearly = STATE.get("kind") == "yearly"
    # 10.16: how the Level 3-2 request was made - 'navigate' (a click / the background frame) or 'cors' (the
    # direct way's fetch).  refuse_fetch imitates a portal that answers only page requests.
    mode = handler.headers.get("Sec-Fetch-Mode", "")
    if url.path.endswith("loadLevelThreeTwo") and method == "POST":
        STATE.setdefault("level3_modes", []).append(mode)
        if STATE.get("refuse_fetch") and mode != "navigate":
            code, html = 403, "<html><head><title>Error</title></head><body><h1>Access Denied</h1></body></html>"
            data = html.encode()
            handler.send_response(code)
            handler.send_header("Content-Type", "text/html; charset=utf-8")
            handler.send_header("Content-Length", str(len(data)))
            handler.end_headers()
            handler.wfile.write(data)
            return True
    if url.path.endswith("loadLevelTwo"):
        code, html = 200, (yearly_level2_page() if yearly else level2_page())
    elif url.path.endswith("loadLevelThreeTwo") and method == "POST":
        code, html = (yearly_level3_page(period, body) if yearly else level3_page(period, body))
    elif yearly and method == "POST" and url.path[len(BASE):].startswith("loadLevelThree"):
        code, html = other_level3_page(url.path[len(BASE) + len("loadLevel"):], period)
    else:
        code, html = 404, "not found"
    data = html.encode()
    handler.send_response(code)
    handler.send_header("Content-Type", "text/html; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)
    return True
