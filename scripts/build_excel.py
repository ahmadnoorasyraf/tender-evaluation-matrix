# Tender Evaluation Matrix - Excel workbook builder
# MIT License - Copyright (c) 2026 Asyraf Abdul. See LICENSE.
# Usage: pip install openpyxl ; python scripts/build_excel.py

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as L

NAVY = "1F3864"; TEAL = "2E75B6"; INPUT = "FFF2CC"; GREY = "F2F2F2"
F = "Arial"
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

def font(**k):
    k.setdefault("name", F); k.setdefault("size", 10); return Font(**k)

def title(ws, text, sub=None):
    ws["A1"] = text; ws["A1"].font = font(size=16, bold=True, color=NAVY)
    if sub:
        ws["A2"] = sub; ws["A2"].font = font(size=9, italic=True, color="595959")
    ws.sheet_view.showGridLines = False

def hdr(c, fill=NAVY):
    c.font = font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX

def inp(c, fmt=None, align="left"):
    c.fill = PatternFill("solid", fgColor=INPUT); c.font = font(color="0000FF"); c.border = BOX
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if fmt: c.number_format = fmt

def calc(c, fmt=None, bold=False, align="center", fill=GREY):
    c.font = font(bold=bold); c.border = BOX
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if fill: c.fill = PatternFill("solid", fgColor=fill)
    if fmt: c.number_format = fmt

def lab(c, bold=False):
    c.font = font(bold=bold); c.border = BOX
    c.alignment = Alignment(vertical="center", wrap_text=True)

GREEN = PatternFill("solid", fgColor="C6EFCE"); RED = PatternFill("solid", fgColor="FFC7CE"); AMBER = PatternFill("solid", fgColor="FFEB9C")
def status_cf(ws, rng):
    for txt, fill, col in [("PASS", GREEN, "006100"), ("Yes", GREEN, "006100"), ("FAIL", RED, "9C0006"),
                           ("No", RED, "9C0006"), ("INCOMPLETE", AMBER, "9C5700")]:
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{txt}"'], fill=fill, font=Font(name=F, bold=True, color=col)))

NB = 6
wb = Workbook()

# ---------------- Start Here ----------------
s = wb.active; s.title = "Start Here"
title(s, "Tender Evaluation Matrix", "Compliance gate → technical scoring → evaluated price → weighted award recommendation")
s.column_dimensions["A"].width = 4; s.column_dimensions["B"].width = 26; s.column_dimensions["C"].width = 90
rows = [
    ("HOW IT WORKS", None),
    ("1. Setup", "Enter project details, technical/commercial weighting, technical pass mark, budget estimate and up to 6 bidder names. Names flow to every tab."),
    ("2. Compliance", "List mandatory requirements (pass/fail). Mark each bidder Yes / No / N/A. Any 'No' disqualifies the bidder; blanks show INCOMPLETE."),
    ("3. Technical", "List criteria and weights (must total 100%). Bidders below the pass mark are not eligible for award."),
    ("4. Eval 1 – Eval 5", "Each evaluator scores every bidder 0–5 on their own tab, independently. Unused tabs are ignored. "
                           "Technical combines them (average or median — choose on Setup)."),
    ("5. Panel Review", "Shows each evaluator's scores side by side, flags criteria where evaluators differ by 2+ points, "
                        "and holds the conflict-of-interest declarations. Use it in the moderation meeting."),
    ("6. Clarifications", "Log every clarification. Enter a cost impact for excluded scope or deviations — it is added automatically to that bidder's evaluated price."),
    ("7. Commercial", "Enter bid prices as submitted plus any arithmetic corrections and other adjustments. The tool builds an evaluated (like-for-like) price."),
    ("8. Summary", "Eligible bidders are price-scored (lowest evaluated price ÷ bidder's price), weighted, totalled and ranked. Add your justification and sign-offs."),
    ("9. Sensitivity", "Shows who wins at different technical/commercial weightings — use it to show management your recommendation is robust."),
    ("", None),
    ("COLOUR LEGEND", None),
    ("Yellow cells, blue text", "Your inputs. Type here."),
    ("Grey cells, black text", "Formulas. Do not overwrite."),
    ("Green / Red / Amber", "PASS / FAIL / INCOMPLETE status flags."),
    ("", None),
    ("EXAMPLE DATA", "The file opens with a worked example (chiller plant tender, 4 bidders) so you can see how it behaves. "
                     "Overwrite the yellow cells with your own tender, or clear them to start fresh."),
    ("Price scoring method", "Price score = lowest evaluated price among ELIGIBLE bidders ÷ this bidder's evaluated price (lowest scores 100%). "
                             "Total = Technical % × technical weight + Price score × commercial weight."),
    ("Works in", "Microsoft Excel 2019 / 365 and Google Sheets (File → Import). No macros."),
]
r = 4
for a, b in rows:
    if b is None and a:
        s.cell(r, 2, a).font = font(bold=True, color="FFFFFF", size=11)
        for c in (2, 3): s.cell(r, c).fill = PatternFill("solid", fgColor=NAVY)
    elif a:
        s.cell(r, 2, a).font = font(bold=True); s.cell(r, 3, b).font = font()
        s.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top"); s.cell(r, 2).alignment = Alignment(vertical="top")
        if a.startswith("Yellow"): inp(s.cell(r, 2)); s.cell(r, 2).font = font(bold=True, color="0000FF")
        if a.startswith("Grey"): calc(s.cell(r, 2), align="left", bold=True)
        s.row_dimensions[r].height = 30
    r += 1

# ---------------- Setup ----------------
su = wb.create_sheet("Setup"); title(su, "Setup", "Fill the yellow cells. Everything else is calculated.")
for col, w in zip("ABCDEFG", [3, 34, 38, 3, 8, 60, 3]): su.column_dimensions[col].width = w
info = [("Tender / project title", "Supply, Installation & Commissioning of 2 x 500RT Chiller Plant", None),
        ("Tender reference", "TND-2026-014", None),
        ("Client / owner", "Example Manufacturing Co.", None),
        ("Evaluation date", "2026-11-11", None),
        ("Currency", "USD", None),
        ("Budget / engineer's estimate", 1200000, "#,##0"),
        (None, None, None),
        ("Technical weight", 0.7, "0%"),
        ("Commercial weight", "=1-C11", "0%"),
        ("Technical pass mark (minimum %)", 0.6, "0%")]
r = 4
for a, v, fmt in info:
    if a:
        lab(su.cell(r, 2, a), bold=True); c = su.cell(r, 3, v)
        if isinstance(v, str) and v.startswith("="): calc(c, fmt, align="left")
        else: inp(c, fmt)
    r += 1
dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", showErrorMessage=True,
                    error="Enter a percentage between 0% and 100%"); su.add_data_validation(dv); dv.add("C11"); dv.add("C13")
su["B16"] = "Bidders (up to 6)"; hdr(su["B16"]); su["C16"] = "Bidder name"; hdr(su["C16"])
bidders = ["Northbridge Engineering", "Coastal M&E Contractors", "Summit Cooling Systems", "Vertex Building Services", "", ""]
for i in range(NB):
    lab(su.cell(17 + i, 2, f"Bidder {i+1}")); inp(su.cell(17 + i, 3, bidders[i] or None))
BID = [f"Setup!$C${17+i}" for i in range(NB)]
NE = 5
lab(su["B14"], bold=True); su["B14"] = "Combine panel scores using"
inp(su["C14"], align="left"); su["C14"] = "Average"
dvm = DataValidation(type="list", formula1='"Average,Median"', allow_blank=False); su.add_data_validation(dvm); dvm.add("C14")
su["B24"] = "Evaluation panel (up to 5)"; hdr(su["B24"]); su["C24"] = "Evaluator name / role"; hdr(su["C24"])
evs = ["A. Rahman – Project Engineering", "S. Lee – Maintenance", "K. Kumar – HSE", "", ""]
for i in range(NE):
    lab(su.cell(25 + i, 2, f"Evaluator {i+1}  →  tab 'Eval {i+1}'")); inp(su.cell(25 + i, 3, evs[i] or None))
su["B30"] = "Leave unused evaluators blank. One evaluator = just use Eval 1."
su["B30"].font = font(size=9, italic=True, color="595959")
EV = [f"Setup!$C${25+i}" for i in range(NE)]
EVN = [f'IF({EV[i]}="","Evaluator {i+1}",{EV[i]})' for i in range(NE)]
su["E3"] = "Score"; hdr(su["E3"]); su["F3"] = "Scoring scale (technical criteria)"; hdr(su["F3"])
scale = ["Not addressed / unacceptable", "Major gaps or serious risk", "Partly meets requirement; notable weaknesses",
         "Meets requirement", "Meets requirement with some added value", "Exceeds requirement; clear added value, low risk"]
for i, t in enumerate(scale):
    c = su.cell(4 + i, 5, i); calc(c, fill=None); inp(su.cell(4 + i, 6, t))
su["F11"] = "Edit the descriptions to suit your organisation's procurement policy."; su["F11"].font = font(size=9, italic=True, color="595959")

# ---------------- Compliance ----------------
co = wb.create_sheet("Compliance"); title(co, "Stage 1 — Mandatory Compliance (Pass / Fail)", "Any 'No' = disqualified. N/A = not applicable.")
co.column_dimensions["A"].width = 5; co.column_dimensions["B"].width = 48
for j in range(NB): co.column_dimensions[L(3 + j)].width = 17
co.column_dimensions[L(3 + NB)].width = 36
hdr(co.cell(4, 1, "No")); hdr(co.cell(4, 2, "Mandatory requirement"))
for j in range(NB): hdr(co.cell(4, 3 + j, f'=IF({BID[j]}="","Bidder {j+1}",{BID[j]})'))
hdr(co.cell(4, 3 + NB, "Notes"))
reqs = ["Bid submitted before closing date/time", "Form of Tender signed by authorised person", "Bid bond / tender security provided",
        "Valid company registration & required licences", "Bid validity at least 90 days", "Price schedule fully completed",
        "Acceptance of contract conditions (or deviations listed)", "HSE policy and method statement submitted"]
ans = {0: ["Yes"] * 8, 1: ["Yes"] * 8, 2: ["Yes"] * 8, 3: ["Yes", "Yes", "No", "Yes", "Yes", "Yes", "Yes", "Yes"]}
R0, R1 = 5, 16
dvc = DataValidation(type="list", formula1='"Yes,No,N/A"', allow_blank=True); co.add_data_validation(dvc)
for k, r in enumerate(range(R0, R1 + 1)):
    calc(co.cell(r, 1, k + 1), fill=None); inp(co.cell(r, 2, reqs[k] if k < len(reqs) else None))
    for j in range(NB):
        c = co.cell(r, 3 + j, ans[j][k] if (j in ans and k < len(reqs)) else None); inp(c, align="center")
    inp(co.cell(r, 3 + NB, "No bid bond enclosed with submission" if k == 2 else None))
dvc.add(f"C{R0}:{L(2+NB)}{R1}")
labels = ["Non-compliances ('No')", "Unanswered", "COMPLIANCE STATUS"]
for i, t in enumerate(labels):
    lab(co.cell(R1 + 1 + i, 2, t), bold=True)
for j in range(NB):
    col = L(3 + j); rng = f"{col}{R0}:{col}{R1}"
    calc(co.cell(R1 + 1, 3 + j, f'=COUNTIF({rng},"No")'))
    calc(co.cell(R1 + 2, 3 + j, f'=COUNTA($B${R0}:$B${R1})-COUNTIF({rng},"Yes")-COUNTIF({rng},"No")-COUNTIF({rng},"N/A")'))
    calc(co.cell(R1 + 3, 3 + j, f'=IF({BID[j]}="","",IF({col}{R1+1}>0,"FAIL",IF({col}{R1+2}>0,"INCOMPLETE","PASS")))'), bold=True)
CSTAT = R1 + 3
status_cf(co, f"C{R0}:{L(2+NB)}{CSTAT}")
co.freeze_panes = "C5"

# ---------------- Technical ----------------
te = wb.create_sheet("Technical"); title(te, "Stage 2 — Technical Evaluation (combined panel score)",
                                         "Set criteria & weights here (must total 100%). Scores are combined automatically from the Eval tabs.")
for col, w in zip("ABCD", [5, 22, 46, 10]): te.column_dimensions[col].width = w
for j in range(NB): te.column_dimensions[L(5 + j)].width = 17
te.column_dimensions[L(5 + NB)].width = 34
for i, t in enumerate(["No", "Category", "Criterion", "Weight"]): hdr(te.cell(4, 1 + i, t))
for j in range(NB): hdr(te.cell(4, 5 + j, f'=IF({BID[j]}="","Bidder {j+1}",{BID[j]})'))
hdr(te.cell(4, 5 + NB, "Evaluator notes"))
crit = [("Technical", "Compliance with specification & scope understanding", .15),
        ("Technical", "Equipment make/model & energy efficiency", .15),
        ("Delivery", "Method statement & installation approach", .10),
        ("Delivery", "Programme & lead time", .10),
        ("Capability", "Relevant experience (similar projects, last 5 years)", .10),
        ("Capability", "Key personnel & site team", .08),
        ("HSE / Quality", "HSE plan & safety record", .08),
        ("HSE / Quality", "Testing & commissioning plan (FAT/SAT)", .07),
        ("After-sales", "Warranty & after-sales service", .07),
        ("After-sales", "Spare parts & local support", .05),
        ("HSE / Quality", "Quality plan / ISO 9001", .05)]
# example scores: EVSC[evaluator][bidder][criterion]
base = [[4, 4, 4, 3, 4, 4, 4, 4, 3, 4, 4], [4, 3, 3, 4, 3, 3, 4, 3, 4, 3, 3],
        [2, 3, 2, 3, 2, 2, 3, 2, 3, 2, 3], [4, 4, 3, 4, 4, 3, 3, 3, 4, 3, 4]]
EVSC = [base,
        [[4, 5, 4, 3, 4, 4, 4, 4, 3, 4, 4], [4, 3, 4, 4, 3, 3, 4, 3, 4, 3, 3],
         [2, 3, 2, 3, 2, 2, 3, 2, 3, 2, 3], [4, 4, 3, 4, 4, 3, 3, 3, 4, 3, 4]],
        [[3, 4, 4, 3, 4, 4, 4, 4, 4, 4, 4], [4, 3, 2, 4, 3, 3, 4, 3, 4, 3, 3],
         [3, 3, 2, 3, 2, 2, 3, 2, 3, 2, 3], [4, 3, 3, 4, 4, 3, 3, 3, 4, 3, 4]]]
T0, T1 = 5, 19
for k, r in enumerate(range(T0, T1 + 1)):
    calc(te.cell(r, 1, k + 1), fill=None)
    has = k < len(crit)
    inp(te.cell(r, 2, crit[k][0] if has else None)); inp(te.cell(r, 3, crit[k][1] if has else None))
    inp(te.cell(r, 4, crit[k][2] if has else None), "0%", "center")
    for j in range(NB):
        col = L(5 + j); args = ",".join(f"'Eval {e+1}'!{col}{r}" for e in range(NE))
        calc(te.cell(r, 5 + j, f'=IF(COUNT({args})=0,"",IF(Setup!$C$14="Median",MEDIAN({args}),AVERAGE({args})))'), "0.0")
    inp(te.cell(r, 5 + NB))
te.cell(3, 5, '="Combined score = "&LOWER(Setup!$C$14)&" of evaluators who scored"').font = font(size=9, italic=True, color="595959")
TW, TS, TP, TC = T1 + 1, T1 + 2, T1 + 3, T1 + 4
lab(te.cell(TW, 3, "Total weight (must be 100%)"), bold=True)
calc(te.cell(TW, 4, f"=SUM(D{T0}:D{T1})"), "0%", bold=True)
te.conditional_formatting.add(f"D{TW}", FormulaRule(formula=[f"ROUND(D{TW},4)<>1"], fill=RED, font=Font(name=F, bold=True, color="9C0006")))
te.conditional_formatting.add(f"D{TW}", FormulaRule(formula=[f"ROUND(D{TW},4)=1"], fill=GREEN, font=Font(name=F, bold=True, color="006100")))
lab(te.cell(TS, 3, "TECHNICAL SCORE (%)"), bold=True)
lab(te.cell(TP, 3, "Meets pass mark?"), bold=True)
lab(te.cell(TC, 3, "Criteria scored"), bold=True)
for j in range(NB):
    col = L(5 + j); rng = f"{col}{T0}:{col}{T1}"
    calc(te.cell(TS, 5 + j, f'=IF(OR({BID[j]}="",$D${TW}=0,COUNT({rng})=0),"",SUMPRODUCT($D${T0}:$D${T1},{rng})/5/$D${TW})'), "0.0%", bold=True)
    calc(te.cell(TP, 5 + j, f'=IF({col}{TS}="","",IF(COUNT({rng})<COUNT($D${T0}:$D${T1}),"INCOMPLETE",IF({col}{TS}>=Setup!$C$13,"PASS","FAIL")))'), bold=True)
    calc(te.cell(TC, 5 + j, f'=IF({BID[j]}="","",COUNT({rng})&" / "&COUNT($D${T0}:$D${T1}))'))
status_cf(te, f"E{TP}:{L(4+NB)}{TP}")
for r in range(T0, T1 + 1): te.row_dimensions[r].height = 28
te.freeze_panes = "E5"

# ---------------- Evaluator sheets ----------------
dvs = DataValidation(type="whole", operator="between", formula1="0", formula2="5", showErrorMessage=True, error="Score must be a whole number 0–5")
EVS = []
for e in range(NE):
    ev = wb.create_sheet(f"Eval {e+1}"); EVS.append(ev)
    title(ev, "x", "Score each criterion 0–5 independently (scale on Setup). Criteria and weights come from the Technical tab.")
    ev["A1"] = f'="Technical Scoring — "&{EVN[e]}'
    for col, w in zip("ABCD", [5, 22, 46, 10]): ev.column_dimensions[col].width = w
    for j in range(NB): ev.column_dimensions[L(5 + j)].width = 17
    ev.column_dimensions[L(5 + NB)].width = 34
    for i, t in enumerate(["No", "Category", "Criterion", "Weight"]): hdr(ev.cell(4, 1 + i, t))
    for j in range(NB): hdr(ev.cell(4, 5 + j, f'=IF({BID[j]}="","Bidder {j+1}",{BID[j]})'))
    hdr(ev.cell(4, 5 + NB, "Evaluator comments / justification"))
    d = DataValidation(type="whole", operator="between", formula1="0", formula2="5", showErrorMessage=True, error="Score must be a whole number 0–5")
    ev.add_data_validation(d)
    for k, r in enumerate(range(T0, T1 + 1)):
        calc(ev.cell(r, 1, k + 1), fill=None)
        for c in (2, 3): calc(ev.cell(r, c, f'=IF(Technical!{L(c)}{r}="","",Technical!{L(c)}{r})'), align="left")
        calc(ev.cell(r, 4, f'=IF(Technical!D{r}="","",Technical!D{r})'), "0%")
        for j in range(NB):
            v = EVSC[e][j][k] if (e < len(EVSC) and j < 4 and k < len(crit)) else None
            inp(ev.cell(r, 5 + j, v), align="center")
        note = "Coastal method statement lacks lifting plan detail" if (e == 2 and k == 2) else \
               ("Coastal method statement well sequenced" if (e == 1 and k == 2) else None)
        inp(ev.cell(r, 5 + NB, note)); ev.row_dimensions[r].height = 28
    d.add(f"E{T0}:{L(4+NB)}{T1}")
    lab(ev.cell(TS, 3, "THIS EVALUATOR'S TECHNICAL SCORE (%)"), bold=True)
    lab(ev.cell(TC, 3, "Criteria scored"), bold=True)
    for j in range(NB):
        col = L(5 + j); rng = f"{col}{T0}:{col}{T1}"
        calc(ev.cell(TS, 5 + j, f'=IF(OR({BID[j]}="",Technical!$D${TW}=0,COUNT({rng})=0),"",SUMPRODUCT(Technical!$D${T0}:$D${T1},{rng})/5/Technical!$D${TW})'), "0.0%", bold=True)
        calc(ev.cell(TC, 5 + j, f'=IF({BID[j]}="","",COUNT({rng})&" / "&COUNT(Technical!$D${T0}:$D${T1}))'))
    ev.freeze_panes = "E5"

# ---------------- Panel Review ----------------
pr = wb.create_sheet("Panel Review"); title(pr, "Panel Review — Evaluator Consistency",
                                            "Use in the moderation meeting: where do evaluators disagree, and does each evaluator rank the bidders the same way?")
pr.column_dimensions["A"].width = 5; pr.column_dimensions["B"].width = 46
for j in range(NB + 2): pr.column_dimensions[L(3 + j)].width = 17
hdr(pr.cell(4, 2, "A. Technical score by evaluator"), TEAL)
for j in range(NB): hdr(pr.cell(4, 3 + j, f'=IF({BID[j]}="","Bidder {j+1}",{BID[j]})'), TEAL)
hdr(pr.cell(4, 3 + NB, "Evaluator's top bidder"), TEAL); hdr(pr.cell(4, 4 + NB, "Scores entered"), TEAL)
for e in range(NE):
    r = 5 + e; lab(pr.cell(r, 2, f"={EVN[e]}"), True)
    for j in range(NB):
        calc(pr.cell(r, 3 + j, f"='Eval {e+1}'!{L(5+j)}{TS}"), "0.0%")
    rr = f"C{r}:{L(2+NB)}{r}"
    calc(pr.cell(r, 3 + NB, f'=IF(COUNT({rr})=0,"Not used",INDEX($C$4:${L(2+NB)}$4,MATCH(MAX({rr}),{rr},0)))'), bold=True)
    calc(pr.cell(r, 4 + NB, f"=SUMPRODUCT(--ISNUMBER('Eval {e+1}'!$E${T0}:${L(4+NB)}${T1}))"))
    pr.conditional_formatting.add(rr, FormulaRule(formula=[f'AND(C{r}<>"",C{r}=MAX(${"C"}{r}:${L(2+NB)}{r}))'], fill=GREEN, font=Font(name=F, bold=True, color="006100")))
PC = 5 + NE
lab(pr.cell(PC, 2, '="COMBINED ("&UPPER(Setup!$C$14)&")"'), True)
for j in range(NB): calc(pr.cell(PC, 3 + j, f"=Technical!{L(5+j)}{TS}"), "0.0%", bold=True)
lab(pr.cell(PC + 1, 2, "Score range across evaluators (max − min)"))
for j in range(NB):
    col = L(3 + j)
    calc(pr.cell(PC + 1, 3 + j, f'=IF(COUNT({col}5:{col}{PC-1})<2,"",MAX({col}5:{col}{PC-1})-MIN({col}5:{col}{PC-1}))'), "0.0%")
lab(pr.cell(PC + 2, 2, "Do all evaluators pick the same top bidder?"), True)
tb = f"{L(3+NB)}5:{L(3+NB)}{PC-1}"
cw = f'INDEX($C$4:${L(2+NB)}$4,MATCH(MAX($C${PC}:${L(2+NB)}${PC}),$C${PC}:${L(2+NB)}${PC},0))'
calc(pr.cell(PC + 2, 3, f'=IF(OR(COUNTIF({tb},"<>Not used")=0,COUNT($C${PC}:${L(2+NB)}${PC})=0),"",IF(COUNTIF({tb},{cw})=COUNTIF({tb},"<>Not used"),"Yes","No – discuss at moderation"))'), bold=True)
pr.merge_cells(start_row=PC + 2, start_column=3, end_row=PC + 2, end_column=5)
status_cf(pr, f"C{PC+2}")

B0 = PC + 5
hdr(pr.cell(B0, 2, "B. Score spread by criterion (max − min of evaluator scores)"), TEAL)
for j in range(NB): hdr(pr.cell(B0, 3 + j, f"={L(3+j)}4"), TEAL)
pr.cell(B0 - 1, 2, "Amber = spread of 2 points, Red = 3+ points. Discuss these items and record the agreed reasoning.").font = font(size=9, italic=True, color="595959")
for k, r0 in enumerate(range(T0, T1 + 1)):
    r = B0 + 1 + k
    calc(pr.cell(r, 2, f'=IF(Technical!C{r0}="","",Technical!C{r0})'), align="left")
    for j in range(NB):
        col = L(5 + j); args = ",".join(f"'Eval {e+1}'!{col}{r0}" for e in range(NE))
        calc(pr.cell(r, 3 + j, f'=IF(COUNT({args})<2,"",MAX({args})-MIN({args}))'))
BE = B0 + (T1 - T0 + 1)
gr = f"C{B0+1}:{L(2+NB)}{BE}"
pr.conditional_formatting.add(gr, FormulaRule(formula=[f'AND(ISNUMBER(C{B0+1}),C{B0+1}>=3)'], fill=RED, font=Font(name=F, bold=True, color="9C0006")))
pr.conditional_formatting.add(gr, FormulaRule(formula=[f'AND(ISNUMBER(C{B0+1}),C{B0+1}>=2)'], fill=AMBER, font=Font(name=F, bold=True, color="9C5700")))
lab(pr.cell(BE + 1, 2, "Items to discuss (spread ≥ 2)"), True)
calc(pr.cell(BE + 1, 3, f'=COUNTIF({gr},">=2")'), bold=True)

D0 = BE + 4
hdr(pr.cell(D0, 2, "C. Evaluator declarations"), TEAL)
for i, t in enumerate(["No conflict of interest declared?", "Signature", "Date"]): hdr(pr.cell(D0, 3 + i * 2, t), TEAL); pr.merge_cells(start_row=D0, start_column=3 + i * 2, end_row=D0, end_column=4 + i * 2)
dvy = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True); pr.add_data_validation(dvy)
for e in range(NE):
    r = D0 + 1 + e; lab(pr.cell(r, 2, f"={EVN[e]}"), True); pr.row_dimensions[r].height = 26
    for i in range(3):
        c = pr.cell(r, 3 + i * 2, ("Yes" if e < 3 else None) if i == 0 else None); inp(c, align="center")
        pr.merge_cells(start_row=r, start_column=3 + i * 2, end_row=r, end_column=4 + i * 2)
    dvy.add(f"C{r}")
status_cf(pr, f"C{D0+1}:C{D0+NE}")
pr.cell(D0 + NE + 1, 2, "I confirm I have no financial or personal interest in any bidder and scored independently before moderation.").font = font(size=9, italic=True, color="595959")

# ---------------- Clarifications ----------------
cl = wb.create_sheet("Clarifications"); title(cl, "Clarifications Log", "Cost impact (+) is added to that bidder's evaluated price on the Commercial tab. Use (−) for credits.")
for col, w in zip("ABCDEFGHI", [5, 26, 13, 44, 12, 44, 12, 16, 10]): cl.column_dimensions[col].width = w
for i, t in enumerate(["No", "Bidder", "Type", "Clarification question", "Date sent", "Bidder response", "Date received", "Cost impact", "Status"]):
    hdr(cl.cell(4, 1 + i, t))
cl["H4"] = '="Cost impact ("&Setup!$C$8&")"'; hdr(cl["H4"])
ex = [("Coastal M&E Contractors", "Commercial", "Price schedule excludes incoming electrical connection to chiller MCC. Please confirm.", "2026-10-28",
       "Confirmed excluded. Can add for 45,000.", "2026-10-30", 45000, "Closed"),
      ("Northbridge Engineering", "Technical", "Confirm chiller part-load efficiency (IPLV) at AHRI conditions.", "2026-10-28",
       "IPLV 0.42 kW/RT confirmed; datasheet attached.", "2026-10-31", 0, "Closed"),
      ("Summit Cooling Systems", "Technical", "Method statement for rigging and lifting not submitted.", "2026-10-28", None, None, None, "Open")]
C0, C1 = 5, 54
dvb = DataValidation(type="list", formula1=f"=Setup!$C$17:$C${16+NB}", allow_blank=True)
dvt = DataValidation(type="list", formula1='"Technical,Commercial,Compliance"', allow_blank=True)
dvst = DataValidation(type="list", formula1='"Open,Closed"', allow_blank=True)
for d in (dvb, dvt, dvst): cl.add_data_validation(d)
for k, r in enumerate(range(C0, C1 + 1)):
    calc(cl.cell(r, 1, k + 1), fill=None)
    row = ex[k] if k < len(ex) else (None,) * 8
    for i, v in enumerate(row):
        c = cl.cell(r, 2 + i, v)
        inp(c, "#,##0" if i == 6 else None, "center" if i in (1, 3, 5, 7) else "left")
dvb.add(f"B{C0}:B{C1}"); dvt.add(f"C{C0}:C{C1}"); dvst.add(f"I{C0}:I{C1}")
cl.conditional_formatting.add(f"I{C0}:I{C1}", CellIsRule(operator="equal", formula=['"Open"'], fill=AMBER, font=Font(name=F, bold=True, color="9C5700")))
cl.freeze_panes = "A5"

# ---------------- Commercial ----------------
cm = wb.create_sheet("Commercial"); title(cm, "Stage 3 — Commercial Evaluation", "Build a like-for-like evaluated price for each bidder.")
cm.column_dimensions["A"].width = 46
for j in range(NB): cm.column_dimensions[L(2 + j)].width = 18
cm["A4"] = '="Item ("&Setup!$C$8&")"'; hdr(cm["A4"])
for j in range(NB): hdr(cm.cell(4, 2 + j, f'=IF({BID[j]}="","Bidder {j+1}",{BID[j]})'))
price = [1250000, 1120000, 980000, 1090000]; corr = [0, 0, 0, 3200]; oth = [0, 0, 0, 0]
lab(cm["A5"], True); cm["A5"] = "Bid price as submitted"
lab(cm["A6"]); cm["A6"] = "Arithmetic corrections (+/−)"
lab(cm["A7"]); cm["A7"] = "Add-backs from Clarifications log (auto)"
lab(cm["A8"]); cm["A8"] = "Other adjustments (missing scope, deviations, +/−)"
lab(cm["A9"], True); cm["A9"] = "EVALUATED PRICE"
lab(cm["A10"]); cm["A10"] = "Variance vs budget estimate"
lab(cm["A11"]); cm["A11"] = "Price rank (all bidders, lowest = 1)"
for j in range(NB):
    col = L(2 + j)
    inp(cm.cell(5, 2 + j, price[j] if j < 4 else None), "#,##0", "right")
    inp(cm.cell(6, 2 + j, corr[j] if j < 4 else None), "#,##0;(#,##0);-", "right")
    calc(cm.cell(7, 2 + j, f'=IF({BID[j]}="","",SUMIFS(Clarifications!$H${C0}:$H${C1},Clarifications!$B${C0}:$B${C1},{BID[j]}))'), "#,##0;(#,##0);-", align="right")
    inp(cm.cell(8, 2 + j, oth[j] if j < 4 else None), "#,##0;(#,##0);-", "right")
    calc(cm.cell(9, 2 + j, f'=IF(OR({BID[j]}="",NOT(ISNUMBER({col}5))),"",{col}5+N({col}6)+N({col}7)+N({col}8))'), "#,##0", bold=True, align="right")
    calc(cm.cell(10, 2 + j, f'=IF(OR({col}9="",NOT(ISNUMBER(Setup!$C$9)),N(Setup!$C$9)=0),"",{col}9/Setup!$C$9-1)'), "+0.0%;-0.0%;0.0%")
    calc(cm.cell(11, 2 + j, f'=IF({col}9="","",COUNTIF($B$9:${L(1+NB)}$9,"<"&{col}9)+1)'))
cm.conditional_formatting.add(f"B10:{L(1+NB)}10", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=F, color="C00000")))
hdr(cm["A13"], TEAL); cm["A13"] = "Other commercial terms (for information — not scored)"
for j in range(NB): hdr(cm.cell(13, 2 + j, f"={L(2+j)}4"), TEAL)
terms = [("Completion / delivery (weeks)", [16, 14, 18, 15]), ("Payment terms", ["30/60/10", "20/70/10", "40/50/10", "30/60/10"]),
         ("Bid validity (days)", [90, 120, 90, 90]), ("Warranty (months)", [24, 18, 12, 24]), ("Liquidated damages accepted?", ["Yes", "Yes", "With cap", "Yes"]),
         ("Commercial deviations / qualifications", ["None", "Electrical connection excluded", "Retention reduced to 5%", "None"])]
for i, (t, vals) in enumerate(terms):
    r = 14 + i; lab(cm.cell(r, 1, t))
    for j in range(NB): inp(cm.cell(r, 2 + j, vals[j] if j < 4 else None), align="center")
    cm.row_dimensions[r].height = 28
cm.freeze_panes = "B5"

# ---------------- Summary ----------------
sm = wb.create_sheet("Summary"); title(sm, "Evaluation Summary & Award Recommendation")
widths = [26, 13, 12, 12, 10, 16, 11, 12, 12, 11, 7, 34]
for i, w in enumerate(widths): sm.column_dimensions[L(1 + i)].width = w
meta = [("Tender", "=Setup!C4"), ("Reference", "=Setup!C5"), ("Client", "=Setup!C6"), ("Evaluation date", "=Setup!C7"),
        ("Weighting (Tech / Comm)", '=TEXT(Setup!C11,"0%")&" / "&TEXT(Setup!C12,"0%")'), ("Technical pass mark", '=TEXT(Setup!C13,"0%")'),
        ("Evaluation panel", f'=_xlfn.TEXTJOIN("; ",TRUE,Setup!C25:C29)&"  ("&Setup!C14&" of scores)"')]
for i, (a, f) in enumerate(meta):
    r = 3 + i; lab(sm.cell(r, 1, a), True); sm.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    calc(sm.cell(r, 2, f), align="left")
heads = ["Bidder", "Compliance", "Technical %", "Technical result", "Eligible?", "Evaluated price", "Price score",
         "Weighted technical", "Weighted commercial", "TOTAL SCORE", "Rank", "Outcome"]
H = 11
for i, t in enumerate(heads): hdr(sm.cell(H, 1 + i, t))
sm.cell(H, 6).value = '="Evaluated price ("&Setup!$C$8&")"'
S0 = H + 1; S1 = S0 + NB - 1
LOW = f"$C${S1+3}"
for j in range(NB):
    r = S0 + j; ccol = L(3 + j); tcol = L(5 + j); mcol = L(2 + j)
    calc(sm.cell(r, 1, f'=IF({BID[j]}="","",{BID[j]})'), align="left", bold=True)
    calc(sm.cell(r, 2, f'=IF(A{r}="","",Compliance!{ccol}{CSTAT})'))
    calc(sm.cell(r, 3, f'=IF(A{r}="","",Technical!{tcol}{TS})'), "0.0%")
    calc(sm.cell(r, 4, f'=IF(A{r}="","",Technical!{tcol}{TP})'))
    calc(sm.cell(r, 6, f'=IF(A{r}="","",Commercial!{mcol}9)'), "#,##0")
    calc(sm.cell(r, 5, f'=IF(A{r}="","",IF(AND(B{r}="PASS",D{r}="PASS",ISNUMBER(F{r})),"Yes","No"))'), bold=True)
    calc(sm.cell(r, 7, f'=IF(E{r}<>"Yes","",{LOW}/F{r})'), "0.0%")
    calc(sm.cell(r, 8, f'=IF(E{r}<>"Yes","",C{r}*Setup!$C$11)'), "0.0%")
    calc(sm.cell(r, 9, f'=IF(E{r}<>"Yes","",G{r}*Setup!$C$12)'), "0.0%")
    calc(sm.cell(r, 10, f'=IF(E{r}<>"Yes","",H{r}+I{r})'), "0.0%", bold=True)
    calc(sm.cell(r, 11, f'=IF(E{r}<>"Yes","",COUNTIFS($E${S0}:$E${S1},"Yes",$J${S0}:$J${S1},">"&J{r})+1)'), bold=True)
    calc(sm.cell(r, 12, f'=IF(A{r}="","",IF(E{r}="Yes",IF(K{r}=1,"RECOMMENDED FOR AWARD","Ranked "&K{r}),'
                         f'"Not eligible – "&IF(B{r}<>"PASS","compliance "&LOWER(B{r}),IF(D{r}<>"PASS","technical "&LOWER(D{r}),"no evaluated price"))))'),
         align="left")
    sm.row_dimensions[r].height = 22
status_cf(sm, f"B{S0}:E{S1}")
sm.conditional_formatting.add(f"A{S0}:L{S1}", FormulaRule(formula=[f"$K{S0}=1"], fill=PatternFill("solid", fgColor="DDEBF7"), font=Font(name=F, bold=True, color=NAVY)))
k = S1 + 2
hdr(sm.cell(k, 1, "Key results"), TEAL); sm.merge_cells(start_row=k, start_column=1, end_row=k, end_column=3)
res = [("Lowest eligible evaluated price", f'=IF(COUNTIF($E${S0}:$E${S1},"Yes")=0,"",_xlfn.MINIFS($F${S0}:$F${S1},$E${S0}:$E${S1},"Yes"))', "#,##0"),
       ("Recommended bidder", f'=IFERROR(INDEX($A${S0}:$A${S1},MATCH(1,$K${S0}:$K${S1},0)),"No eligible bidder")', None),
       ("Recommended bidder's evaluated price", f'=IFERROR(INDEX($F${S0}:$F${S1},MATCH(1,$K${S0}:$K${S1},0)),"")', "#,##0"),
       ("Variance vs budget estimate", f'=IF(OR(C{k+3}="",N(Setup!$C$9)=0),"",C{k+3}/Setup!$C$9-1)', "+0.0%;-0.0%;0.0%"),
       ("Eligible bidders", f'=COUNTIF($E${S0}:$E${S1},"Yes")&" of "&COUNTIF($A${S0}:$A${S1},"?*")', None),
       ("Winner holds across all sensitivity cases?", '=Sensitivity!C14', None)]
for i, (a, f, fmt) in enumerate(res):
    r = k + 1 + i; lab(sm.cell(r, 1, a), True); sm.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    calc(sm.cell(r, 3, f), fmt, bold=True, align="left"); sm.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
assert LOW == f"$C${k+1}"
status_cf(sm, f"C{k+6}")
j0 = k + 8
hdr(sm.cell(j0, 1, "Recommendation & justification"), TEAL); sm.merge_cells(start_row=j0, start_column=1, end_row=j0, end_column=12)
sm.merge_cells(start_row=j0 + 1, start_column=1, end_row=j0 + 4, end_column=12)
inp(sm.cell(j0 + 1, 1, "Northbridge Engineering is recommended: fully compliant, highest combined technical score from the 3-person panel (77.1%) and within 4.2% of budget. "
                       "Coastal M&E is 6.8% cheaper after add-backs but scored lower on equipment efficiency and method statement. "
                       "Summit (lowest price) did not meet the 60% technical pass mark; Vertex did not provide a bid bond. Note: Coastal would win only if commercial weight were raised to 60%."))
sm.cell(j0 + 1, 1).alignment = Alignment(wrap_text=True, vertical="top")
a0 = j0 + 6
for i, t in enumerate(["", "Name", "Designation", "Signature", "Date"]):
    c = sm.cell(a0, 1 + i * (2 if i else 1) if False else [1, 2, 5, 8, 11][i], t); hdr(c, TEAL)
for c0, c1 in [(2, 4), (5, 7), (8, 10), (11, 12)]: sm.merge_cells(start_row=a0, start_column=c0, end_row=a0, end_column=c1)
for i, t in enumerate(["Prepared by", "Reviewed by", "Approved by"]):
    r = a0 + 1 + i; lab(sm.cell(r, 1, t), True); sm.row_dimensions[r].height = 30
    for c0, c1 in [(2, 4), (5, 7), (8, 10), (11, 12)]:
        inp(sm.cell(r, c0)); sm.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
ch = BarChart(); ch.type = "bar"; ch.grouping = "stacked"; ch.overlap = 100
ch.title = "Total score — eligible bidders"; ch.y_axis.number_format = "0%"; ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1
ch.add_data(Reference(sm, min_col=8, max_col=9, min_row=H, max_row=S1), titles_from_data=True)
ch.set_categories(Reference(sm, min_col=1, min_row=S0, max_row=S1))
ch.height = 7.5; ch.width = 16; ch.legend.position = "b"
sm.add_chart(ch, f"A{a0+6}")

# ---------------- Sensitivity ----------------
se = wb.create_sheet("Sensitivity"); title(se, "Weighting Sensitivity", "Who wins if the technical / commercial split changes? Edit the yellow weights.")
se.column_dimensions["A"].width = 26
for j in range(8): se.column_dimensions[L(2 + j)].width = 14
se["A4"] = "Technical weight →"
ws_ = [0.4, 0.5, 0.6, 0.7, 0.8, 0.9]
for i, w in enumerate(ws_): inp(se.cell(4, 2 + i, w), "0%", "center")
hdr(se.cell(5, 1, "Bidder"))
for i in range(len(ws_)): hdr(se.cell(5, 2 + i, f'=TEXT({L(2+i)}4,"0%")&" / "&TEXT(1-{L(2+i)}4,"0%")'))
for j in range(NB):
    r = 6 + j; sr = S0 + j
    calc(se.cell(r, 1, f'=Summary!A{sr}'), align="left", bold=True)
    for i in range(len(ws_)):
        col = L(2 + i)
        calc(se.cell(r, 2 + i, f'=IF(Summary!E{sr}<>"Yes","",Summary!C{sr}*{col}$4+Summary!G{sr}*(1-{col}$4))'), "0.0%")
W = 6 + NB
lab(se.cell(W, 1, "Winner"), True)
for i in range(len(ws_)):
    col = L(2 + i)
    calc(se.cell(W, 2 + i, f'=IF(COUNT({col}6:{col}{W-1})=0,"",INDEX($A$6:$A${W-1},MATCH(MAX({col}6:{col}{W-1}),{col}6:{col}{W-1},0)))'), bold=True)
    se.conditional_formatting.add(f"{col}6:{col}{W-1}", FormulaRule(formula=[f'AND({col}6<>"",{col}6=MAX({col}$6:{col}${W-1}))'], fill=GREEN, font=Font(name=F, bold=True, color="006100")))
assert W + 2 == 14
lab(se.cell(14, 1, "Recommendation robust?"), True); se.merge_cells("A14:B14")
calc(se["C14"], bold=True)
se["C14"] = (f'=IF(Summary!C{k+2}="No eligible bidder","",IF(COUNTIF(B{W}:{L(1+len(ws_))}{W},Summary!C{k+2})=COUNTIF(B{W}:{L(1+len(ws_))}{W},"?*"),'
             f'"Yes","No – changes at some weightings"))')
se.merge_cells("C14:F14")
se["A16"] = ("If the winner changes within a reasonable weighting range, say so in the justification. "
             "Fix the weighting BEFORE bids are opened — never adjust it to favour a bidder.")
se["A16"].font = font(size=9, italic=True, color="595959")
status_cf(se, "C14")

# tab colours + print setup
order = ["Start Here", "Setup", "Compliance"] + [f"Eval {e+1}" for e in range(NE)] + \
        ["Technical", "Panel Review", "Clarifications", "Commercial", "Summary", "Sensitivity"]
wb._sheets = [wb[n] for n in order]
for ws, colr in [(s, NAVY), (su, "FFC000"), (co, TEAL), (te, TEAL), (pr, TEAL), (cl, "A6A6A6"), (cm, TEAL), (sm, "00B050"), (se, "00B050")] + [(x, "FFC000") for x in EVS]:
    ws.sheet_properties.tabColor = colr
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
import os
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "excel", "Tender_Evaluation_Matrix.xlsx")
wb.save(out)
print("saved")
