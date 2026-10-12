"""Test copy of the live PAYE certificate screens (user's capture of 8 Oct 2026: Recorder JSON + HAR).

What the live portal does, and this copy does too (structure only - every name, TIN and amount is invented):
* loadPayePage      the form: Return For (#retBrnch, onchange showBrtab), #branchListTable, Period (#period,
                    values '2025-2026' ...), Search button -> searchPayeDtls() -> submitform() -> a real POST
                    that REPLACES the page (not drawn by script) to searchPayeEntry.
* searchPayeEntry   the same form again + the employees grid: outer table FORM_TAB_BORDER > td > table with a
                    <thead> row of 9 <th> (Select All | Sr No. | Employee Tin | Employee Name | Chargeable Income
                    | Tax Deducted | TCC No. | (link) | hidden allowTCC) and one <tr> per employee, then a row
                    with the GenerateTCC button.  Every page also has an <iframe id="chatbot_widget" src="">
                    and window.onload appends a second one (live: both about:blank).
* Get Month-Wise Details -> OpenWindowWithPostWithParam(): window.open('post.htm', 'Uganda_Revenue_Authority')
                    and a form POST to getMonthsDtls targeted at that window - a SEPARATE pop-up window, where
                    #divPostContent holds the months and a Close button (3 Oct recording, steps 28-35).
"""
import time
from urllib.parse import parse_qs, urlparse

BASE = "/Portal/genTaxCreditCertController/"
EMPLOYEES = {      # invented: (TIN, name, chargeable income, tax deducted, TCC No.)
    "2025-2026": [("1000000011", "MR. TEST EMPLOYEE ONE", "36,000,000.00", "9,300,000.00", ""),
                  ("1000000022", "MS. TEST EMPLOYEE TWO", "12,000,000.00", "2,100,000.00", ""),
                  ("1000000033", "MR. TEST EMPLOYEE THREE", "6,000,000.00", "540,000.00", "TCC000000000001")],
    "2026-2027": [("1000000011", "MR. TEST EMPLOYEE ONE", "9,000,000.00", "2,325,000.00", "")],
    "2024-2025": [(f"10000001{n:02d}", f"MR. TEST WORKER {n}", "4,800,000.00", "300,000.00", "") for n in range(1, 8)],
}
STATE = {"search_delay": 0.6, "searches": [], "monthwise": [], "generated": [], "drop_presses": 0,
         "hang": {}}
# 10.15 (live 8 Oct 2026, 22:40: 'no results within 60 s after Search'):
#   drop_presses  the next page's Search press is swallowed this many times (the page's frmSubmitFlag is set
#                 but nothing is sent) - a press that never reaches the portal;
#   hang          {period: n} - the next n searches for that period get no answer (the server sits on them).


def reset(**kw):
    STATE.update({"search_delay": 0.6, "searches": [], "monthwise": [], "generated": [], "drop_presses": 0,
                  "hang": {}})
    STATE.update(kw)


PERIODS = ["2023-2024", "2024-2025", "2025-2026", "2026-2027"]


def _label(fy):
    a, b = fy.split("-")
    return f"July {a} to June {b}"


SCRIPT = """
<script>
var frmSubmitFlag = 0; var sCount = 0; var count = 5;      /* live page: at most 5 certificates per press */
function submitform(f) { if (frmSubmitFlag == 0) { frmSubmitFlag = 1; f.submit(); } }
function showBrtab() { var o = document.getElementById('brtrid');
  if (o) o.style.display = document.getElementById('retBrnch').value == 'B' ? '' : 'none'; }
function validate() { if (!document.getElementById('period').value) { alert('Please select Period'); return false; }
  return true; }
function searchPayeDtls() { if (dropPress > 0) { dropPress--; frmSubmitFlag = 1; return false; }
  if (validate()) {
  document.TaxCreditCertForm.action = '/Portal/genTaxCreditCertController/searchPayeEntry';
  submitform(document.TaxCreditCertForm); } else { return false; } }
function generateTcc() { var b = document.getElementsByName('selectedEntry'), any = false;
  for (var j = 0; j < b.length; j++) if (b[j].checked) any = true;
  if (!any) { alert('Please select atleast one Employee to generate Tax Credit Certificate'); return false; }
  document.TaxCreditCertForm.action = '/Portal/genTaxCreditCertController/generatePAYETCC';
  submitform(document.TaxCreditCertForm); return true; }
function selectRow(ele, lastEle, tccNo) { if (lastEle.innerHTML.trim() == 'N') {
  alert('A TCC with Reference Number: ' + tccNo.innerHTML.trim() + ' has already been generated for this period.');
  ele.checked = false; return false; }
  if (ele.checked && sCount < count) { sCount++; return true; }
  else if (!ele.checked) { sCount--; return true; }
  else { ele.checked = false; alert('You can generate maximum ' + count + ' certificates at a time.'); } }
function OpenWindowWithPostWithParam(url, windowoption, name, params) {
  var form = document.createElement('form'); form.setAttribute('method', 'post');
  form.setAttribute('action', url); form.setAttribute('target', name);
  for (var i in params) { var inp = document.createElement('input'); inp.type = 'hidden';
    inp.name = i; inp.value = params[i]; form.appendChild(inp); }
  document.body.appendChild(form); window.open('post.htm', name, windowoption); form.submit();
  document.body.removeChild(form); }
function getMonthsDtls(tinId, nameId) {
  var p = {tin: document.getElementById(tinId).innerHTML.trim(), name: document.getElementById(nameId).innerHTML.trim(),
           retFor: document.getElementById('retBrnch').value, branchId: '', period: document.getElementById('period').value};
  OpenWindowWithPostWithParam('/Portal/genTaxCreditCertController/getMonthsDtls',
     'width=900,height=600', 'Uganda_Revenue_Authority', p); }
window.onload = function () { var f = document.createElement('iframe'); f.setAttribute('id', 'chatbot_widget');
  f.setAttribute('width', '0'); f.setAttribute('height', '0'); f.setAttribute('src', ''); document.body.appendChild(f); };
</script>"""


def _form(period="", grid=""):
    opts = "".join(f'<option value="{p}"{" selected" if p == period else ""}>{_label(p)}</option>'
                   for p in PERIODS)
    rows = grid and ('<table class="FORM_TAB_BORDER"><tr><td>' + grid + '</td></tr></table>')
    drop, STATE["drop_presses"] = STATE["drop_presses"], 0
    return (f"""<html><head><title>Generate PAYE Tax Credit Certificate | eTax</title>
<script>var dropPress = {drop};</script>{SCRIPT}</head><body>
<form name="TaxCreditCertForm" id="TaxCreditCertForm" method="POST">
<input type="hidden" name="actionCode" value="TAXCREDITCIRTVO"><input type="hidden" name="selectedBranchId" id="selectedBranchId" value="">
<table class="FORM_TAB_BORDER"><thead><tr><th>Generate PAYE Tax Credit Certificate</th></tr></thead>
<tr><td>Return For*</td><td><select id="retBrnch" name="retBrnch" onchange="showBrtab();">
<option value="H" selected>Headquarter</option><option value="B">Branch</option></select></td></tr>
<tr id="brtrid" style="display:none"><td><table class="FORM_TAB_BORDER" id="branchListTable"><thead><tr>
<td>Select</td><td>Sr No.</td><td>Applied For Branch</td><td>Branch Address</td></tr></thead></table></td></tr>
<tr><td>Period*</td><td><select class="form101_textf" name="period" id="period" onchange=""><option value="">--Select--</option>{opts}</select></td></tr>
<tr><td><button type="button" value="Search" class="submit" onclick="return searchPayeDtls();">Search</button>
<input type="button" value="Back" class="submit"></td></tr></table>
{rows}
</form>
<form name="footerForm" method="POST"><iframe id="chatbot_widget" width="325px" height="65px" src="" style="border:0"></iframe></form>
</body></html>""")


def _grid(period):
    emps = EMPLOYEES.get(period, [])
    if not emps:
        return ""
    head = "".join(f"<th><label>{h}</label></th>" for h in
                   ("Sr No.", "Employee Tin", "Employee Name", "Chargeable Income", "Tax Deducted", "TCC No.", ""))
    body = ""
    for n, (tin, name, inc, tax, tcc) in enumerate(emps, 1):
        tcc = tcc or next((t for p, w, t in STATE["generated"] if p == period and w == tin), "")
        allow = "N" if tcc else "Y"
        body += (f'<tr><td align="center"><input type="checkbox" name="selectedEntry" id="listPaye" '
                 f'onclick="selectRow(this,document.getElementById(\'allowTCC_{n}\'),document.getElementById(\'tccNo_{n}\'));" '
                 f'value="{tin}" /></td><td>{n}</td><td id="tin_{n}"> {tin}</td><td id="name_{n}"> {name}</td>'
                 f'<td align="right">{inc}</td><td align="right">{tax}</td><td id="tccNo_{n}"> {tcc}</td>'
                 f'<td><a href="javascript:getMonthsDtls(\'tin_{n}\',\'name_{n}\')">Get Month-Wise Details</a></td>'
                 f'<td id="allowTCC_{n}" style="display: none">{allow}</td></tr>')
    return ('<table class="FORM_TAB_BORDER"><thead><tr><th><input type="checkbox" name="selectAll" id="selectAll" />'
            f' &nbsp; Select All</th>{head}<th style="display: none"><label>A</label></th></tr></thead><tbody>{body}'
            '<tr><td align="center" colspan="9"><input type="button" value="GenerateTCC" class="submit" '
            'onclick="return generateTcc();" /></td></tr></tbody></table>')


def _monthwise(form):
    tin, period = form.get("tin", [""])[0], form.get("period", ["2025-2026"])[0]
    STATE["monthwise"].append(tin)
    y = int(period[:4])
    months = [("JUL", y), ("AUG", y), ("SEP", y), ("OCT", y), ("NOV", y), ("DEC", y),
              ("JAN", y + 1), ("FEB", y + 1), ("MAR", y + 1), ("APR", y + 1), ("MAY", y + 1), ("JUN", y + 1)]
    rows = "".join(f"<tr><td>{m}-{yy}</td><td>3,000,000.00</td><td>775,000.00</td></tr>" for m, yy in months)
    return (f"""<html><head><title>Month-Wise Details</title></head><body><div id="divPostContent">
<table><tr><th>Month</th><th>Chargeable Income</th><th>Tax Deducted</th></tr>{rows}
<tr><td colspan="3"><input type="button" value="Close" onclick="window.close()"></td></tr></table></div></body></html>""")


def handle(handler, method):
    url = urlparse(handler.path)
    if not url.path.startswith(BASE):
        return False
    length = int(handler.headers.get("Content-Length") or 0)
    form = parse_qs(handler.rfile.read(length).decode()) if (method == "POST" and length) else {}
    name = url.path[len(BASE):]
    if name == "loadPayePage":
        html = _form()
    elif name == "searchPayeEntry" and method == "POST":
        period = form.get("period", [""])[0]
        STATE["searches"].append((form.get("retBrnch", [""])[0], period))
        if STATE["hang"].get(period, 0) > 0:
            STATE["hang"][period] -= 1
            time.sleep(20)                       # no answer while the program waits (its wait is shortened)
        time.sleep(STATE["search_delay"])
        html = _form(period, _grid(period))
    elif name == "generatePAYETCC" and method == "POST":
        period = form.get("period", [""])[0]
        for i, tin in enumerate(form.get("selectedEntry", [])):
            STATE["generated"].append((period, tin, f"TCC9{len(STATE['generated']) + 1:011d}"))
        html = _form(period, _grid(period))
    elif name == "getMonthsDtls" and method == "POST":
        html = _monthwise(form)
    elif name == "post.htm":
        html = "<html><body></body></html>"
    else:
        handler.send_error(404)
        return True
    data = html.encode()
    handler.send_response(200)
    handler.send_header("Content-Type", "text/html; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)
    return True
