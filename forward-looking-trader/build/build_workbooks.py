"""Builds Forecast_Registry.xlsx and Source_Register.xlsx (openpyxl).
Run: python3 build_workbooks.py ; then recalculate in Excel/LibreOffice."""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter as L

OUT = ".."
F = "Arial"
BLUE = Font(name=F, size=10, color="0000FF")
BLK = Font(name=F, size=10)
GRN = Font(name=F, size=10, color="008000")
BOLD = Font(name=F, size=10, bold=True)
H1 = Font(name=F, size=14, bold=True, color="1F3A5F")
H2 = Font(name=F, size=11, bold=True, color="FFFFFF")
NOTE = Font(name=F, size=9, italic=True, color="555555")
HDR = PatternFill("solid", fgColor="1F3A5F")
SUB = PatternFill("solid", fgColor="DCE6F1")
INP = PatternFill("solid", fgColor="FFF9DB")
thin = Side(style="thin", color="BBBBBB")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")

def hdr(ws, row, labels, col=1):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=t)
        c.font = H2; c.fill = HDR; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BOX

def put(ws, ref, v, font=BLK, fmt=None, fill=None, wrap=False):
    c = ws[ref]; c.value = v; c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if wrap: c.alignment = WRAP
    return c

# ====================================================================== Forecast registry
wb = Workbook()
rd = wb.active; rd.title = "README"
reg = wb.create_sheet("Registry")
rev = wb.create_sheet("Revisions")
cal = wb.create_sheet("Calibration")
sc = wb.create_sheet("Scenario_Exercise")

# ---------------- README
rd.column_dimensions["A"].width = 26; rd.column_dimensions["B"].width = 110
put(rd, "A1", "01_Forecast_Registry — The Forward-Looking Trader", H1)
put(rd, "A2", "Version 1.0 · built 16 Nov 2026 · Excel-first; no macros; recalculates in Excel 2010+ and LibreOffice.", NOTE)
rows = [
 ("Purpose", "Record dated, measurable, probabilistic predictions; log every revision (including decisions not to act); score them with the Brier score; and check calibration. Forecast accuracy is scored here independently of P&L."),
 ("Sheets", "Registry = one row per prediction.  Revisions = one row per probability update.  Calibration = scoring summary (all formulas).  Scenario_Exercise = the Lesson 1 worked 18–24-month scenario exercise (hypothetical company)."),
 ("Colour legend", "Blue text on pale-yellow fill = cells you type into.  Black = formulas (do not overwrite).  Green = formula pulling from another sheet."),
 ("How to add a prediction", "Registry: fill the next empty row — ID, data type, as-of date, a question with ONE unambiguous yes/no answer, the resolution criteria and source, the event (resolution) date, a base rate, and your initial probability. Record the probability BEFORE looking at anyone else's view."),
 ("How to revise", "Never overwrite the initial probability. Add a row on Revisions (ID, date, new probability, what changed, action taken — 'no action' is a valid entry). The Registry's 'Current probability' picks up the most recent (lowest-down) revision row for that ID automatically."),
 ("How to resolve", "When the event date passes, type 1 (happened) or 0 (did not happen) in Outcome. Leave blank until then. Rows past their event date without an outcome show OVERDUE."),
 ("Brier score", "For one binary prediction: (p − o)², where p is the probability assigned to 'yes' and o is 1 or 0. Range 0 (perfect) to 1 (certain and wrong). Always saying 50% scores 0.25. Lower is better. Mean Brier = average across resolved predictions. (Brier 1950, one-sided binary form.)"),
 ("Skill vs base rate", "Brier skill = 1 − (your mean Brier ÷ reference mean Brier). Positive = you beat the reference. Two references are shown: your own recorded base rates, and the 'climatology' forecast that always states the observed frequency of yes outcomes."),
 ("Calibration table", "Resolved predictions are grouped into ten probability bins on the final (current) probability. If you are well calibrated, things you call 70% happen about 70% of the time. With fewer than ~10 predictions per bin the table is noise: read it as a habit-builder, not a verdict."),
 ("Decomposition", "Murphy (1973): Brier = Reliability − Resolution + Uncertainty. This holds exactly only when all forecasts inside a bin are identical; the 'Decomposition residual' cell shows the within-bin difference."),
 ("Definitions kept distinct", "Probability = your numeric estimate that a stated event occurs by a stated date.  Forecast confidence = how much evidence stands behind that probability (how far it would move on new data).  Conviction = a feeling; never recorded here as a number.  Position size = a risk-budget decision made later (Module 5); it is not a probability."),
 ("Sample rows", "Rows R-001…R-005 are a HISTORICAL REPLAY on Micron Technology: outcomes are real (SEC-filed figures, see Source_Register.xlsx M1–M13); the probabilities are illustrative values written for teaching, not forecasts anyone actually recorded in 2022. Rows H-001…H-007 are HYPOTHETICAL demo rows (invented outcomes) so the Calibration sheet has something to compute. Rows L-001…L-005 are the open prediction ladder for the hypothetical Calder Forge exercise. Delete or overwrite all sample rows before real use; base rates in sample rows are placeholders, not researched frequencies."),
 ("Limits", "Binary questions only (multi-outcome and numeric-range scoring arrive in Week 2). 200 Registry rows and 400 Revision rows are wired; extend ranges if you need more. Scoring uses the final probability, which rewards late updating; the initial-probability Brier is shown alongside for that reason. Nothing here is investment advice."),
]
for i, (a, b) in enumerate(rows, start=4):
    put(rd, f"A{i}", a, BOLD).alignment = WRAP
    put(rd, f"B{i}", b, BLK, wrap=True)
    rd.row_dimensions[i].height = max(30, 15 * (len(b) // 105 + 1))

# ---------------- Registry
RH = 5; R0 = 6; RN = 205
put(reg, "A1", "Forecast Registry", H1)
put(reg, "A2", "Scoring date (used for OVERDUE flag):", BOLD)
put(reg, "D2", "=TODAY()", BLK, "dd-mmm-yyyy")
put(reg, "E2", "Type a fixed date over this formula to score as of a past date.", NOTE)
put(reg, "A3", "Blue/yellow = type here. Black = formula. Never overwrite the initial probability — log changes on the Revisions sheet.", NOTE)
cols = ["ID", "Data type", "As-of date", "Question (one unambiguous yes/no answer)", "Resolution criteria & source", "Event (resolution) date",
        "Horizon (months)", "Horizon bucket", "Base rate", "Initial probability", "Evidence for", "Evidence against",
        "Current probability", "# revisions", "Outcome (1 = yes, 0 = no)", "Status", "Brier — initial", "Brier — final",
        "Brier — base rate", "Linked thesis / module", "Notes"]
hdr(reg, RH, cols)
widths = [8, 16, 12, 46, 40, 12, 9, 9, 9, 10, 34, 34, 10, 9, 10, 10, 9, 9, 9, 18, 30]
for i, w in enumerate(widths, 1): reg.column_dimensions[L(i)].width = w
reg.row_dimensions[RH].height = 44
reg.freeze_panes = "E6"
inputs = [1, 2, 3, 4, 5, 6, 9, 10, 11, 12, 15, 20, 21]
for r in range(R0, RN + 1):
    for c in inputs:
        cell = reg.cell(row=r, column=c); cell.font = BLUE; cell.fill = INP; cell.border = BOX; cell.alignment = WRAP
    for c in (3, 6): reg.cell(row=r, column=c).number_format = "dd-mmm-yyyy"
    for c in (9, 10): reg.cell(row=r, column=c).number_format = "0%"
    f = {
     7: f'=IF(OR(C{r}="",F{r}=""),"",ROUND((F{r}-C{r})/30.4375,1))',
     8: f'=IF(G{r}="","",IF(G{r}<=4.5,3,IF(G{r}<=9,6,IF(G{r}<=15,12,IF(G{r}<=21,18,24)))))',
     13: f'=IF(J{r}="","",IFERROR(INDEX(Revisions!$C$6:$C$405,SUMPRODUCT(MAX((Revisions!$A$6:$A$405=A{r})*ROW(Revisions!$A$6:$A$405)))-5),J{r}))',
     14: f'=IF(A{r}="","",COUNTIF(Revisions!$A$6:$A$405,A{r}))',
     16: f'=IF(D{r}="","",IF(ISNUMBER(O{r}),"Resolved",IF(F{r}<$D$2,"OVERDUE","Open")))',
     17: f'=IF(AND(ISNUMBER(O{r}),ISNUMBER(J{r})),(J{r}-O{r})^2,"")',
     18: f'=IF(AND(ISNUMBER(O{r}),ISNUMBER(M{r})),(M{r}-O{r})^2,"")',
     19: f'=IF(AND(ISNUMBER(O{r}),ISNUMBER(I{r})),(I{r}-O{r})^2,"")',
    }
    for c, v in f.items():
        cell = reg.cell(row=r, column=c, value=v); cell.font = GRN if c in (13, 14) else BLK; cell.border = BOX
    reg.cell(row=r, column=13).number_format = "0%"
    for c in (17, 18, 19): reg.cell(row=r, column=c).number_format = "0.000"
dvp = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True,
                     error="Enter a probability between 0 and 1 (e.g. 0.35).", errorTitle="Probability")
dvo = DataValidation(type="list", formula1='"0,1"', allow_blank=True, error="Outcome must be 1 (yes) or 0 (no). Leave blank until resolved.")
reg.add_data_validation(dvp); reg.add_data_validation(dvo)
dvp.add(f"I{R0}:J{RN}"); dvo.add(f"O{R0}:O{RN}")
reg.conditional_formatting.add(f"P{R0}:P{RN}", CellIsRule(operator="equal", formula=['"OVERDUE"'], fill=PatternFill("solid", bgColor="F8CBAD", fgColor="F8CBAD")))

D = dt.date
REPLAY = "Historical replay (real outcome; illustrative probability)"
HYP = "Hypothetical demo"
OPEN = "Hypothetical exercise (open)"
A0 = D(2022, 12, 22)
sample = [
 ("R-001", REPLAY, A0, "Micron reports positive GAAP gross margin in at least one fiscal quarter ending on or before 30 Jun 2024.",
  "GAAP gross margin from 10-Q/10-K. Resolved YES: FQ2 FY24 (ended 29 Feb 2024) gross margin 18.5% (Source Register M7).", D(2024, 6, 30), 0.70, 0.75,
  "Company states it is reducing wafer starts and capex (10-Q filed 22 Dec 2022); supply response under way.",
  "Company expects industry profitability 'challenged throughout calendar 2023'; inventories elevated.", 1, "L1 case: memory cycle", "Base rate is a placeholder."),
 ("R-002", REPLAY, A0, "Micron quarterly revenue exceeds $6.64B (the FQ4 FY22 level) in a fiscal quarter ending on or before 30 Jun 2024.",
  "Revenue from 10-Q/10-K. Resolved YES: FQ3 FY24 (ended 30 May 2024) revenue $6.81B (M8).", D(2024, 6, 30), 0.35, 0.40,
  "Supply cuts plus eventual restocking; revenue only needs to return to a level seen one quarter before the trough began.",
  "Revenue had fallen 39% q/q; both volume and price would need to recover within six quarters.", 1, "L1 case: memory cycle", ""),
 ("R-003", REPLAY, A0, "Micron quarterly capex (PP&E expenditures) is below $2.0B in each of FQ2, FQ3 and FQ4 of FY2023.",
  "Cash-flow statement, expenditures for PP&E. Resolved NO: FQ2 FY23 capex $2.21B (M3).", D(2023, 8, 31), 0.50, 0.55,
  "Explicit statement of capex reductions; FQ1 FY23 capex already down to $2.45B from $3.61B.",
  "Committed equipment deliveries and construction in progress cannot be cancelled immediately — spending lags decisions.", 0, "L1 case: memory cycle", "Teaching point: decisions lead cash outlays."),
 ("R-004", REPLAY, A0, "Micron reports positive GAAP net income for FQ4 FY2023 (quarter ending 31 Aug 2023).",
  "GAAP net income, 10-K FY2023. Resolved NO: FQ4 FY23 net loss $1.43B (M5).", D(2023, 8, 31), 0.15, 0.20,
  "Cost actions announced.", "Management itself guided to a challenged calendar 2023; under-utilisation charges ahead.", 0, "L1 case: memory cycle", ""),
 ("R-005", REPLAY, A0, "Micron quarterly GAAP net income exceeds $1.49B (the FQ4 FY22 level) in a fiscal quarter ending on or before 31 Dec 2024.",
  "GAAP net income, 10-Q. Resolved YES: FQ1 FY25 (ended 28 Nov 2024) net income $1.87B (M10).", D(2024, 12, 31), 0.30, 0.35,
  "If supply discipline holds, operating leverage on recovering prices is large.", "Requires full recovery of both price and volume within 24 months.", 1, "L1 case: memory cycle", ""),
 ("H-001", HYP, D(2025, 1, 15), "HYPOTHETICAL: Industry forging capacity is lower at the 6-month check than at the as-of date.", "Hypothetical data.", D(2025, 7, 15), 0.80, 0.90, "Two closures already announced.", "Announcements can be reversed.", 1, "Demo", "Invented outcome."),
 ("H-002", HYP, D(2025, 1, 15), "HYPOTHETICAL: Quoted lead time for large forgings exceeds 30 weeks at the 6-month check.", "Hypothetical data.", D(2025, 7, 15), 0.70, 0.85, "Order books lengthening.", "Single-supplier survey.", 1, "Demo", "Invented outcome."),
 ("H-003", HYP, D(2025, 1, 15), "HYPOTHETICAL: Contract price index is up at least 5% at the 12-month check.", "Hypothetical data.", D(2026, 1, 15), 0.45, 0.70, "Utilisation rising.", "Customers have annual contracts.", 0, "Demo", "Invented outcome."),
 ("H-004", HYP, D(2025, 3, 1), "HYPOTHETICAL: The largest competitor announces no new capacity in the next 9 months.", "Hypothetical data.", D(2025, 12, 1), 0.60, 0.60, "Balance sheet stretched.", "Subsidy programme pending.", 1, "Demo", "Invented outcome."),
 ("H-005", HYP, D(2025, 3, 1), "HYPOTHETICAL: Company gross margin exceeds 25% in any quarter within 12 months.", "Hypothetical data.", D(2026, 3, 1), 0.25, 0.30, "Operating leverage.", "Price increases lag by contract.", 0, "Demo", "Invented outcome."),
 ("H-006", HYP, D(2025, 3, 1), "HYPOTHETICAL: A substitute technology wins a major customer qualification within 12 months.", "Hypothetical data.", D(2026, 3, 1), 0.10, 0.15, "Pilot programmes exist.", "Qualification cycles take years.", 0, "Demo", "Invented outcome."),
 ("H-007", HYP, D(2025, 6, 1), "HYPOTHETICAL: Industry utilisation is above 85% at the 3-month check.", "Hypothetical data.", D(2025, 9, 1), 0.50, 0.65, "Closures effective.", "Demand soft in one end market.", 1, "Demo", "Invented outcome."),
]
A1 = D(2026, 11, 16)
ladder = [
 ("L-001", 3, D(2027, 2, 16), "CALDER EXERCISE (hypothetical): At least one of the two announced competitor plant closures is confirmed as physically shut.", 0.75, 0.80),
 ("L-002", 6, D(2027, 5, 16), "CALDER EXERCISE (hypothetical): Quoted industry lead times for large forgings are longer than at the as-of date.", 0.55, 0.60),
 ("L-003", 12, D(2027, 11, 16), "CALDER EXERCISE (hypothetical): Industry contract price is at least 5% above the as-of price of $2,000/t.", 0.40, 0.45),
 ("L-004", 18, D(2028, 5, 16), "CALDER EXERCISE (hypothetical): Industry utilisation is at or above 85%.", 0.35, 0.40),
 ("L-005", 24, D(2028, 11, 16), "CALDER EXERCISE (hypothetical): Calder's trailing-four-quarter EPS is at or above $0.80.", 0.25, 0.30),
]
for (i, m, ed, q, br, p) in ladder:
    sample.append((i, OPEN, A1, q, "Hypothetical: resolves from the exercise data pack.", ed, br, p,
                   "See Scenario_Exercise scenario A.", "See Scenario_Exercise scenarios B and C.", None, "L1 worked exercise", f"{m}-month rung of the prediction ladder."))
for k, s in enumerate(sample):
    r = R0 + k
    for c, v in zip([1, 2, 3, 4, 5, 6, 9, 10, 11, 12, 15, 20, 21], s):
        if v is not None and v != "": reg.cell(row=r, column=c).value = v
    reg.row_dimensions[r].height = 62

# ---------------- Revisions
put(rev, "A1", "Revision Log", H1)
put(rev, "A2", "One row per update. Add new rows BELOW existing ones: the Registry reads the lowest row for each ID as the current probability.", NOTE)
put(rev, "A3", "Log non-actions too: 'new data, probability unchanged, no action' is a legitimate and useful entry.", NOTE)
hdr(rev, 5, ["Forecast ID", "Revision date", "New probability", "What changed (evidence, with source/date)", "Why it moves the probability this much", "Action taken (or deliberate non-action)", "Initial p (lookup)", "Change vs initial"])
for i, w in enumerate([11, 13, 11, 52, 44, 36, 10, 10], 1): rev.column_dimensions[L(i)].width = w
rev.freeze_panes = "A6"
for r in range(6, 406):
    for c in range(1, 7):
        cell = rev.cell(row=r, column=c); cell.font = BLUE; cell.fill = INP; cell.border = BOX; cell.alignment = WRAP
    rev.cell(row=r, column=2).number_format = "dd-mmm-yyyy"; rev.cell(row=r, column=3).number_format = "0%"
    g = rev.cell(row=r, column=7, value=f'=IF(A{r}="","",IFERROR(INDEX(Registry!$J$6:$J$205,MATCH(A{r},Registry!$A$6:$A$205,0)),"ID?"))'); g.font = GRN; g.number_format = "0%"; g.border = BOX
    h = rev.cell(row=r, column=8, value=f'=IF(OR(A{r}="",NOT(ISNUMBER(G{r})),NOT(ISNUMBER(C{r}))),"",C{r}-G{r})'); h.font = BLK; h.number_format = '+0%;-0%;0%'; h.border = BOX
dv2 = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True); rev.add_data_validation(dv2); dv2.add("C6:C405")
revs = [
 ("R-002", D(2023, 3, 29), 0.35, "FQ2 FY23 10-Q filed 29 Mar 2023: revenue $3.69B, gross margin −32.7%, net loss $2.31B (Source Register M3).",
  "Worse than the December snapshot, but the loss includes inventory write-downs and the supply response is confirmed. Small move down: the path is worse, the mechanism is intact.", "No action. Noted that the share price is ABOVE its 22 Dec 2022 level despite the worst quarter — price is not waiting for earnings."),
 ("R-002", D(2024, 1, 15), 0.70, "FQ1 FY24 10-Q (quarter ended 30 Nov 2023): revenue $4.73B, third consecutive sequential increase; gross margin −0.7% vs −32.7% three quarters earlier (M6).",
  "Leading indicators that were forecast (margin turning before revenue fully recovers) are now observed. Remaining gap to $6.64B is ~40% with two reported quarters left in the window.", "Illustrative: raise probability; review whether price already reflects recovery (Proposition 2)."),
 ("R-004", D(2023, 6, 29), 0.03, "FQ3 FY23 10-Q filed 29 Jun 2023: net loss $1.90B, gross margin −17.8% (M4).", "One quarter left in window; loss too large to close.", "No action."),
]
for k, s in enumerate(revs):
    for c, v in enumerate(s, 1): rev.cell(row=6 + k, column=c).value = v
    rev.row_dimensions[6 + k].height = 62

# ---------------- Calibration
for i, w in enumerate([34, 12, 12, 12, 14, 14, 12, 16, 16], 1): cal.column_dimensions[L(i)].width = w
put(cal, "A1", "Calibration & Scoring", H1)
put(cal, "A2", "All cells on this sheet are formulas reading the Registry. Scored on the final (current) probability of resolved predictions.", NOTE)
RG = "Registry!$%s$6:$%s$205"
M, O, Hh, Rr, Q, S, I, J = (RG % (x, x) for x in "MOHRQSIJ")
put(cal, "A4", "Summary", BOLD, fill=SUB); put(cal, "B4", None, fill=SUB)
summ = [
 ("Predictions recorded", f'=COUNTA(Registry!$D$6:$D$205)', "0"),
 ("Resolved", f'=COUNT({O})', "0"),
 ("Open / overdue", '=B5-B6', "0"),
 ("Mean Brier — final probability", f'=IF(B6=0,"",AVERAGE({Rr}))', "0.000"),
 ("Mean Brier — initial probability", f'=IF(B6=0,"",AVERAGE({Q}))', "0.000"),
 ("Mean Brier — your base rates", f'=IF(COUNT({S})=0,"",AVERAGE({S}))', "0.000"),
 ("Observed frequency of YES (ō)", f'=IF(B6=0,"",AVERAGE({O}))', "0.0%"),
 ("Mean final probability", f'=IF(B6=0,"",SUMPRODUCT(--ISNUMBER({O}),{M}*1)/B6)', "0.0%"),
 ("Bias (mean probability − ō)", '=IF(B6=0,"",B12-B11)', "+0.0%;-0.0%;0.0%"),
 ("Uncertainty ō(1−ō)", '=IF(B6=0,"",B11*(1-B11))', "0.000"),
 ("Reliability (lower is better)", '=IF(B6=0,"",SUM(H22:H31)/B6)', "0.000"),
 ("Resolution (higher is better)", '=IF(B6=0,"",SUM(I22:I31)/B6)', "0.000"),
 ("Decomposition residual (within-bin)", '=IF(B6=0,"",B8-(B15-B16+B14))', "0.000"),
 ("Brier skill vs climatology", '=IF(OR(B6=0,B14=0),"",1-B8/B14)', "0.0%"),
 ("Brier skill vs your base rates", '=IF(OR(B6=0,B10="",B10=0),"",1-B8/B10)', "0.0%"),
]
for k, (a, f, fmt) in enumerate(summ):
    put(cal, f"A{5+k}", a); put(cal, f"B{5+k}", f, GRN, fmt)
# the mean-final-prob formula: M holds "" text in empty rows -> M*1 errors. Use SUMIFS instead.
cal["B12"].value = f'=IF(B6=0,"",SUMIFS({M},{O},">=0")/B6)'
hdr(cal, 21, ["Probability bin (final p)", "Bin low", "Bin high", "N resolved", "Mean forecast", "Observed frequency", "Gap (obs − fcst)", "N × (fcst − obs)²", "N × (obs − ō)²"])
for k in range(10):
    r = 22 + k; lo = k / 10; hi = (k + 1) / 10 if k < 9 else 1.0001
    put(cal, f"A{r}", f"{k*10}% – {(k+1)*10}%" + (" (incl. 100%)" if k == 9 else ""))
    put(cal, f"B{r}", lo, BLK, "0%"); put(cal, f"C{r}", hi, BLK, "0%")
    crit = f'{M},">="&B{r},{M},"<"&C{r},{O},">=0"'
    put(cal, f"D{r}", f'=COUNTIFS({crit})', GRN, "0")
    put(cal, f"E{r}", f'=IF(D{r}=0,"",SUMIFS({M},{crit})/D{r})', GRN, "0.0%")
    put(cal, f"F{r}", f'=IF(D{r}=0,"",SUMIFS({O},{crit})/D{r})', GRN, "0.0%")
    put(cal, f"G{r}", f'=IF(D{r}=0,"",F{r}-E{r})', BLK, "+0.0%;-0.0%;0.0%")
    put(cal, f"H{r}", f'=IF(D{r}=0,0,D{r}*(E{r}-F{r})^2)', BLK, "0.000")
    put(cal, f"I{r}", f'=IF(D{r}=0,0,D{r}*(F{r}-$B$11)^2)', BLK, "0.000")
put(cal, "A32", "Total", BOLD); put(cal, "D32", "=SUM(D22:D31)", BOLD, "0")
put(cal, "E32", '=IF(D32=B6,"OK: all resolved rows binned","CHECK: binned N ≠ resolved N")', NOTE)
cal.conditional_formatting.add("G22:G31", CellIsRule(operator="greaterThan", formula=["0.1"], fill=PatternFill("solid", bgColor="C6E0B4", fgColor="C6E0B4")))
cal.conditional_formatting.add("G22:G31", CellIsRule(operator="lessThan", formula=["-0.1"], fill=PatternFill("solid", bgColor="F8CBAD", fgColor="F8CBAD")))
put(cal, "A33", "Gap > +10 pts (green): events happened more often than you said — under-confident in that bin. Gap < −10 pts (red): over-confident. Fewer than ~10 per bin: do not over-read.", NOTE)
hdr(cal, 36, ["Horizon bucket (months)", "N resolved", "Mean Brier — final", "Mean Brier — initial"])
for k, b in enumerate([3, 6, 12, 18, 24]):
    r = 37 + k
    put(cal, f"A{r}", b, BLK, "0")
    put(cal, f"B{r}", f'=COUNTIFS({Hh},A{r},{O},">=0")', GRN, "0")
    put(cal, f"C{r}", f'=IF(B{r}=0,"",SUMIFS({Rr},{Hh},A{r},{O},">=0")/B{r})', GRN, "0.000")
    put(cal, f"D{r}", f'=IF(B{r}=0,"",SUMIFS({Q},{Hh},A{r},{O},">=0")/B{r})', GRN, "0.000")
put(cal, "A43", "Reference points: always-50% forecaster scores 0.250. A forecaster who always states the observed frequency scores the Uncertainty value in B14.", NOTE)

# ---------------- Scenario exercise
for i, w in enumerate([46, 16, 16, 16, 20, 58], 1): sc.column_dimensions[L(i)].width = w
put(sc, "A1", "Lesson 1 worked exercise — Calder Forge (HYPOTHETICAL company and industry)", H1)
put(sc, "A2", "An 18–24-month future-state exercise. Every number is invented for teaching. Blue/yellow cells are inputs; overwrite them (or copy this tab) to build your own industry forecast.", NOTE)
def sec(r, t):
    for c in range(1, 7): sc.cell(row=r, column=c).fill = HDR
    put(sc, f"A{r}", t, H2, fill=HDR)
def inp(r, label, v, fmt, unit="", note=""):
    put(sc, f"A{r}", label); c = put(sc, f"B{r}", v, BLUE, fmt, INP); c.border = BOX
    put(sc, f"C{r}", unit, NOTE); put(sc, f"D{r}", note, NOTE)
def calc(r, label, f, fmt, unit="", note="", bold=False):
    put(sc, f"A{r}", label, BOLD if bold else BLK); put(sc, f"B{r}", f, BOLD if bold else BLK, fmt); put(sc, f"C{r}", unit, NOTE); put(sc, f"D{r}", note, NOTE)
sec(4, "1. The three clocks (dates)")
inp(5, "As-of date (observe the present)", A1, "dd-mmm-yyyy", "", "Every input below is 'as of' this date.")
inp(6, "Forecast horizon used for the future state", 21, "0", "months", "Midpoint of the 18–24-month window.")
calc(7, "Forecast date", "=EDATE(B5,B6)", "dd-mmm-yyyy")
calc(8, "Forecast window opens", "=EDATE(B5,18)", "dd-mmm-yyyy")
calc(9, "Forecast window closes", "=EDATE(B5,24)", "dd-mmm-yyyy", "", "The holding period is NOT set here: it is decided by evidence and price (Section 4 of the lesson).")
sec(10, "2. Present state")
inp(11, "Industry effective capacity", 1000, "#,##0", "kt / year")
inp(12, "Industry demand", 780, "#,##0", "kt / year")
calc(13, "Industry utilisation", "=B12/B11", "0.0%")
inp(14, "Industry contract price", 2000, "$#,##0", "$ / t")
inp(15, "Company capacity", 200, "#,##0", "kt / year")
calc(16, "Company volume", "=B15*MIN(B13,B33)", "#,##0.0", "kt / year", "Assumes the company runs at industry utilisation, capped at the practical maximum.")
inp(17, "Variable cost", 1450, "$#,##0", "$ / t")
inp(18, "Fixed cost", 70, "$#,##0.0", "$ m / year")
inp(19, "Net debt", 200, "$#,##0", "$ m")
inp(20, "Interest rate on net debt", 0.06, "0.0%")
inp(21, "Tax rate (on positive pre-tax income only)", 0.25, "0.0%", "", "No tax credit is booked on losses (conservative simplification).")
inp(22, "Diluted shares", 50, "#,##0.0", "m")
inp(23, "Share price today", 4.5, "$#,##0.00", "$")
put(sc, "A24", "Present economics (what the screen shows today)", BOLD, fill=SUB)
calc(25, "Revenue", "=B16*B14/1000", "$#,##0.0", "$ m")
calc(26, "EBIT", "=B16*(B14-B17)/1000-B18", "$#,##0.0;($#,##0.0)", "$ m")
calc(27, "Interest", "=B19*B20", "$#,##0.0", "$ m")
calc(28, "Pre-tax income", "=B26-B27", "$#,##0.0;($#,##0.0)", "$ m")
calc(29, "Net income", "=B28-MAX(0,B28)*B21", "$#,##0.0;($#,##0.0)", "$ m")
calc(30, "EPS", "=B29/B22", "$0.000;($0.000)", "$")
calc(31, "P/E on present earnings", '=IF(B30>0,B23/B30,"n.m.")', "0.0x", "", "The 'present' view: the stock looks very expensive on trough earnings.")
put(sc, "A32", "Global mechanism assumptions", BOLD, fill=SUB)
inp(33, "Practical maximum utilisation", 0.95, "0.0%")
inp(34, "Price sensitivity k", 1.5, "0.00", "", "% change in price per 1 percentage point change in utilisation. The key causal assumption — stress it.")
inp(35, "Variable-cost inflation to forecast date", 0.04, "0.0%")
inp(36, "Fixed-cost inflation to forecast date", 0.03, "0.0%")
sec(38, "3. Future states at the forecast date")
hdr(sc, 39, ["", "A · Tightening", "B · Slow heal", "C · Secular fade", "Check / probability-weighted", "What it means"])
def srow(r, label, vals, fmt, note="", e=None, efmt=None):
    put(sc, f"A{r}", label)
    for col, v in zip("BCD", vals):
        c = put(sc, f"{col}{r}", v, BLUE, fmt, INP); c.border = BOX
    if e: put(sc, f"E{r}", e, BOLD, efmt or fmt)
    put(sc, f"F{r}", note, NOTE)
def crow(r, label, tmpl, fmt, note="", e=None, efmt=None, bold=False):
    put(sc, f"A{r}", label, BOLD if bold else BLK)
    for col in "BCD": put(sc, f"{col}{r}", tmpl.replace("#", col), BOLD if bold else BLK, fmt)
    if e: put(sc, f"E{r}", e, BOLD, efmt or fmt)
    put(sc, f"F{r}", note, NOTE)
srow(40, "Probability", [0.35, 0.40, 0.25], "0%", "Must sum to 100%. Record these in the Registry before modelling prices.", '=IF(ABS(SUM(B40:D40)-1)<0.0001,"OK: sums to 100%","ERROR: probabilities must sum to 100%")')
srow(41, "Demand change to forecast date", [0.06, 0.02, -0.08], "+0.0%;-0.0%", "A: grid/power orders grow. C: substitution takes volume.")
srow(42, "Capacity change to forecast date", [-0.08, -0.03, -0.04], "+0.0%;-0.0%", "Closures already announced vs. actually executed; no additions possible inside the lead time.")
srow(43, "Multiple the market pays on that EPS in that state", [10, 13, 8], "0.0x", "Cyclicals usually get LOWER multiples on high earnings. An explicit judgement — Week 11 replaces it with a reverse DCF.")
srow(44, "Floor value per share if EPS ≤ 0", [1.5, 1.5, 1.5], "$0.00", "Distress/asset-value placeholder used when earnings cannot be capitalised.")
put(sc, "A45", "Mechanics", BOLD, fill=SUB)
crow(46, "Industry demand (kt)", "=$B$12*(1+#41)", "#,##0.0")
crow(47, "Industry capacity (kt)", "=$B$11*(1+#42)", "#,##0.0")
crow(48, "Industry utilisation", "=#46/#47", "0.0%", "Proposition 1 lives here: demand can grow and utilisation still fall if capacity grows faster.")
crow(49, "Change in utilisation (pts)", "=#48-$B$13", "+0.0%;-0.0%")
crow(50, "Contract price ($/t)", "=$B$14*(1+$B$34*#49)", "$#,##0")
crow(51, "Variable cost ($/t)", "=$B$17*(1+$B$35)", "$#,##0")
crow(52, "Company volume (kt)", "=$B$15*MIN(#48,$B$33)", "#,##0.0")
crow(53, "Revenue ($m)", "=#52*#50/1000", "$#,##0.0")
crow(54, "Fixed cost ($m)", "=$B$18*(1+$B$36)", "$#,##0.0")
crow(55, "EBIT ($m)", "=#52*(#50-#51)/1000-#54", "$#,##0.0;($#,##0.0)")
crow(56, "Interest ($m)", "=$B$27", "$#,##0.0", "Net debt held constant (simplification: ignores cash build/burn on the path).")
crow(57, "Pre-tax income ($m)", "=#55-#56", "$#,##0.0;($#,##0.0)")
crow(58, "Net income ($m)", "=#57-MAX(0,#57)*$B$21", "$#,##0.0;($#,##0.0)")
crow(59, "EPS ($)", "=#58/$B$22", "$0.00;($0.00)", "", "=SUMPRODUCT(B40:D40,B59:D59)", bold=True)
crow(60, "Value per share at forecast date ($)", "=IF(#59>0,MAX(#59*#43,#44),#44)", "$0.00", "Undiscounted value at the forecast date.", "=SUMPRODUCT(B40:D40,B60:D60)", bold=True)
crow(61, "Return vs today's price", "=#60/$B$23-1", "+0.0%;-0.0%", "", "=E60/$B$23-1", bold=True)
crow(62, "P/E at TODAY's price on FUTURE EPS", '=IF(#59>0,$B$23/#59,"n.m.")', "0.0x", "The 'future' view of the same stock.")
sec(64, "4. From forecast to price: is there a gap?")
calc(65, "Expected value per share", "=SUMPRODUCT(B40:D40,B60:D60)", "$0.00", "", "Σ probability × value.", True)
calc(66, "Expected return to forecast date", "=B65/B23-1", "+0.0%;-0.0%")
calc(67, "Annualised expected return", "=(1+B66)^(12/B6)-1", "+0.0%;-0.0%", "", "If recognition takes the whole horizon. Earlier recognition raises it; that is a timing question, not a forecast question.")
calc(68, "Probability the position loses money", "=SUMPRODUCT((B60:D60<B23)*B40:D40)", "0%", "", "A positive expected value can coexist with a better-than-even chance of loss.")
calc(69, "Return in the single most likely state", "=INDEX(B61:D61,MATCH(MAX(B40:D40),B40:D40,0))", "+0.0%;-0.0%")
calc(70, "Probability-weighted upside per share", "=SUMPRODUCT((B60:D60>B23)*B40:D40*(B60:D60-B23))", "$0.00")
calc(71, "Probability-weighted downside per share", "=SUMPRODUCT((B60:D60<B23)*B40:D40*(B23-B60:D60))", "$0.00")
calc(72, "Upside / downside ratio", '=IF(B71=0,"n.m.",B70/B71)', "0.00x")
calc(73, "Value if A does NOT happen (B and C at your relative odds)", "=(C40*C60+D40*D60)/(C40+D40)", "$0.00")
calc(74, "Market-implied (break-even) probability of A", "=(B23-B73)/(B60-B73)", "0.0%", "", "The probability of A at which today's price equals expected value, taking your state values as given.", True)
calc(75, "Your probability of A minus market-implied", "=B40-B74", "+0.0%;-0.0%", "", "Your entire edge in this framing. Is your evidence worth this many points?", True)
sec(77, "5. Sensitivity: move probability between A and C")
hdr(sc, 78, ["Points moved from C to A", "P(A)", "P(C)", "Expected value", "Expected return", ""])
for k, sh in enumerate([-0.15, -0.10, -0.05, 0, 0.05, 0.10]):
    r = 79 + k
    c = put(sc, f"A{r}", sh, BLUE, "+0%;-0%;0%", INP); c.border = BOX
    put(sc, f"B{r}", f"=$B$40+A{r}", BLK, "0%"); put(sc, f"C{r}", f"=$D$40-A{r}", BLK, "0%")
    put(sc, f"D{r}", f"=$B$65+A{r}*($B$60-$D$60)", BLK, "$0.00"); put(sc, f"E{r}", f"=D{r}/$B$23-1", BLK, "+0.0%;-0.0%")
    put(sc, f"F{r}", f'=IF(OR(B{r}<0,C{r}<0),"invalid: negative probability","")', NOTE)
sec(86, "6. Your written answers (type here)")
qs = ["Proposition 1 — What would make the demand forecast right but industry profitability wrong?",
      "Proposition 2 — What does today's price already assume? (Use B74.)",
      "Proposition 3 — What could make the value right but the trade fail (timing, liquidity, leverage, path)?",
      "Leading indicators you will watch, with Registry IDs (L-001 … L-005):",
      "Evidence that would cut P(A) in half:"]
for k, q in enumerate(qs):
    r = 87 + k
    put(sc, f"A{r}", q, BLK, wrap=True); sc.merge_cells(f"B{r}:F{r}")
    c = sc[f"B{r}"]; c.fill = INP; c.font = BLUE; c.alignment = WRAP; sc.row_dimensions[r].height = 48
put(sc, "A93", "Model limits: one-period snapshot at the forecast date; no discounting; net debt constant; linear price rule; company share of industry volume constant; multiples are judgements. Each is revisited in Modules 2–3.", NOTE)
sc.freeze_panes = "B5"

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
wb.save(f"{OUT}/Forecast_Registry.xlsx")
print("registry written")
