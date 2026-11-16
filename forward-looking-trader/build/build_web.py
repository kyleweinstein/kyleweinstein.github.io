"""Assembles the browser edition (single HTML page) from the document fragments + Source_Register.xlsx.
Usage: FLT_HTML=<dir> node doc_*.js ; python3 build_web.py <fragdir> <out.html>"""
import sys, html, re
from openpyxl import load_workbook
frag, out = sys.argv[1], sys.argv[2]
def rd(n, swap=True):
    s = open(f"{frag}/{n}.html").read()
    for a, b in [] if not swap else [("Open Forecast_Registry.xlsx at the <em>Scenario_Exercise</em> tab. Every number below is in that sheet and every result is a live formula.", "Open the <a href=\"#lab\">Scenario Lab</a> tab. Every number below is in that calculator and every result updates live."),
                 ("Forecast_Registry.xlsx (tabs <em>Registry</em>, <em>Revisions</em>, <em>Calibration</em>, <em>Scenario_Exercise</em>) and Industry_Foresight_Template.docx", "the Scenario Lab, Registry and Template tabs on this page"),
                 ("In the <em>Scenario_Exercise</em> tab", "In the Scenario Lab tab"), ("<em>Scenario_Exercise</em>", "the Scenario Lab"),
                 ("Forecast_Registry.xlsx → Registry", "Registry tab"), ("Forecast_Registry.xlsx → Revisions", "Registry tab → Revise"),
                 ("Industry_Foresight_Template.docx", "the Template tab"), ("Source_Register.xlsx", "the Sources tab"), ("Project_Progress_Log.docx", "the Sources tab"),
                 ("by the formulas in Forecast_Registry.xlsx", "by the formulas in the course workbook"), ("Forecast_Registry.xlsx", "the Registry tab"),
                 ("with the spreadsheet open", "with the Scenario Lab open"), ("If the table of contents is empty, right-click it in Word and choose Update Field.", "")]:
        s = s.replace(a, b)
    return s
e = html.escape
def xl(sheet, cols):
    ws = load_workbook("../Source_Register.xlsx")[sheet]
    heads = [ws.cell(row=4, column=c).value for c in cols]
    rows = []
    for r in range(5, ws.max_row + 1):
        v = [ws.cell(row=r, column=c).value for c in cols]
        if v[0] is None: continue
        def cell(x):
            x = e(str(x or ""))
            return re.sub(r"(https://[^\s;]+)", r'<a href="\1" target="_blank" rel="noopener">link</a>', x)
        rows.append("<tr>" + "".join(f"<td>{cell(x)}</td>" for x in v) + "</tr>")
    return '<div class="tw"><table class="wide"><thead><tr>' + "".join(f"<th>{e(h)}</th>" for h in heads) + "</tr></thead><tbody>" + "".join(rows) + "</tbody></table></div>"
sources = "<h2>Source register</h2><p>Verification grades describe what was actually checked when this release was built on 16 Nov 2026 (UTC).</p>" + xl("Sources", [1, 2, 4, 5, 6, 8, 10]) + "<h2>Claim-level map</h2>" + xl("Claims", [1, 2, 3, 4, 5]) + "<h2>Bibliography</h2>" + xl("Bibliography", [1, 2, 3])
page = open("web_shell.html").read()
for k, v in {"LESSON": rd("01_Module_1_Learner_Guide"), "KEY": rd("01_Module_1_Instructor_Key"), "COURSE": rd("00_Curriculum_Blueprint"),
             "TEMPLATE": rd("Industry_Foresight_Template"), "SOURCES": sources + rd("Project_Progress_Log", swap=False)}.items():
    page = page.replace("<!--" + k + "-->", v)
open(out, "w").write(page); print("wrote", out, len(page))
