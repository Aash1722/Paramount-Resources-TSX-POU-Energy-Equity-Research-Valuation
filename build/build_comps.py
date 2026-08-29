"""
Paramount Resources (TSX:POU) - Module 5: Comparable Company & Transaction Analysis
Trading comps, M&A comps, and a bridge from peer multiples to an implied share price.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
F = "Arial"
BLUE  = Font(name=F, size=10, color="0000FF")
BLACK = Font(name=F, size=10)
GREEN = Font(name=F, size=10, color="008000")
BOLD  = Font(name=F, size=10, bold=True)
BOLDW = Font(name=F, size=10, bold=True, color="FFFFFF")
RED   = Font(name=F, size=10, italic=True, color="C00000")
REDB  = Font(name=F, size=10, bold=True, color="C00000")
TITLE = Font(name=F, size=14, bold=True, color="1F3864")
SUB   = Font(name=F, size=10, italic=True, color="595959")
HDR   = PatternFill("solid", fgColor="1F3864")
SECF  = PatternFill("solid", fgColor="D9E2F3")
YEL   = PatternFill("solid", fgColor="FFFF00")
GRN   = PatternFill("solid", fgColor="E2EFDA")
ORG   = PatternFill("solid", fgColor="FCE4D6")
TOP   = Border(top=Side(style="thin", color="404040"))
CUR   = '$#,##0;($#,##0);-'
CUR1  = '$#,##0.00;($#,##0.00);-'
NUM   = '#,##0;(#,##0);-'
NUM1  = '#,##0.0;(#,##0.0);-'
PCT   = '0%'
MULT  = '0.0"x"'

def title(ws, n, s):
    ws["A1"] = n; ws["A1"].font = TITLE
    ws["A2"] = s; ws["A2"].font = SUB
def sec(ws, r, t, w=10):
    ws.cell(r, 1, t).font = Font(name=F, size=10, bold=True, color="1F3864")
    for c in range(1, w+1): ws.cell(r, c).fill = SECF
    return r+1
def W(ws, s):
    for k, v in s.items(): ws.column_dimensions[k].width = v

# =====================================================================
# README
# =====================================================================
ws = wb.active; ws.title = "README"
title(ws, "Paramount Resources — Module 5: Comparable Company & Transaction Analysis",
      "Where Paramount trades against its peers, and what acquirers have actually paid for Montney and Duvernay assets.")
W(ws, {"A": 30, "B": 104})
r = 4
r = sec(ws, r, "WHY THIS MODULE EXISTS", 2)
for k, v in [
    ("The problem with NAV alone", "Module 4 produced a NAV of C$31.65 per share, but that number is only as good as its assumptions. Comparables test it against what the market and actual acquirers are paying, which is evidence rather than assumption."),
    ("Two kinds of comparable", "Trading comps show what public investors pay for a minority stake. Transaction comps show what a strategic buyer pays for control, which normally carries a premium. Both matter and they answer different questions."),
    ("The metrics that matter here", "Canadian producers are compared on EV to adjusted funds flow, enterprise value per flowing barrel, and enterprise value per barrel of 2P reserves. Price-to-earnings is close to useless in this sector because depletion accounting swamps it."),
]:
    ws.cell(r, 1, k).font = BOLD
    c = ws.cell(r, 2, v); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 44; r += 1
r += 1
r = sec(ws, r, "READ THIS BEFORE USING THE NUMBERS", 2)
c = ws.cell(r, 1, "DATA FRESHNESS")
c.font = REDB
v = ("Comps go stale within days. Share prices, share counts and net debt for the peer group are marked in the Data status column. "
     "Anything flagged VERIFY is my estimate and must be refreshed from a terminal, SEDAR+ filing or the company's latest release before you quote a single multiple. "
     "Production, cash flow and reserve figures are sourced from Q2 2026 releases and are reliable. This is normal — a live comps sheet is refreshed constantly, and knowing which cells are soft is part of the discipline.")
c2 = ws.cell(r, 2, v); c2.alignment = Alignment(wrap_text=True, vertical="top"); c2.font = RED
ws.row_dimensions[r].height = 60; r += 2

r = sec(ws, r, "THE HEADLINE FINDING", 2)
for v in [
    "The Canadian gas peer group is disappearing. ARC Resources agreed to be acquired by Shell in April 2026 for roughly C$22 billion, and Ovintiv closed its acquisition of NuVista in February 2026 for roughly C$3.8 billion. Two of the most obvious Paramount comparables were removed from the market inside four months.",
    "That consolidation is itself an argument. Scarcity value accrues to the remaining independent Montney and Duvernay operators with owned infrastructure and long inventory — a description that fits Paramount closely.",
    "Shell explicitly noted that around 40% of ARC's production was liquids but that those liquids generated about 70% of revenue. That is external, third-party confirmation of the six-to-one energy versus value distinction running through this entire project.",
]:
    c = ws.cell(r, 1, "•  " + v); c.font = BLACK
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 46; r += 1
r += 1
ws.cell(r, 1, "Hypothetical case study. Not investment advice.").font = SUB

# =====================================================================
# TRADING COMPS
# =====================================================================
ws = wb.create_sheet("Trading_Comps")
title(ws, "Trading Comparables — Canadian Gas-Weighted Producers",
      "Blue cells are inputs. All multiples are formulas. Check the Data status column before quoting anything.")
W(ws, {"A": 22, "B": 8, "C": 10, "D": 10, "E": 11, "F": 10, "G": 11, "H": 11, "I": 11,
       "J": 8, "K": 11, "L": 11, "M": 9, "N": 11, "O": 11, "P": 9, "Q": 16})

hdrs = ["Company", "Ticker", "Price (C$)", "Shares (mm)", "Mkt cap (C$mm)", "Net debt (C$mm)",
        "EV (C$mm)", "Q2'26 (Boe/d)", "2026E (Boe/d)", "Liq %", "AFF ann. (C$mm)",
        "2P (MBoe)", "EV/AFF", "EV/flow Q2", "EV/flow 26E", "EV/2P", "Data status"]
r = 4
for i, h in enumerate(hdrs):
    c = ws.cell(r, 1+i, h); c.font = BOLDW; c.fill = HDR
    c.alignment = Alignment(horizontal="right" if i > 1 else "left", wrap_text=True, vertical="center")
ws.row_dimensions[r].height = 30
r += 1
first = r

# name, tkr, price, shares, netdebt, q2prod, prod26, liq, aff_ann, twop, status
comps = [
    ("Paramount Resources", "POU", 32.65, 145.8, -537.1, 47279, 52133, 0.49, 702.4, 521518,
     "Verified — my model"),
    ("Tourmaline Oil", "TOU", 62.00, 389.0, 1500.0, 594198, 630000, 0.13, 3144.4, None,
     "VERIFY price & shares"),
    ("Peyto Exploration", "PEY", 24.00, 198.0, 1250.0, 145320, 148000, 0.11, 910.8, None,
     "VERIFY price, shares, debt"),
    ("Birchcliff Energy", "BIR", 5.93, 276.0, 250.0, 77562, 83500, 0.16, 374.8, None,
     "Price from Q2 buyback; VERIFY debt"),
    ("Kelt Exploration", "KEL", 6.50, 190.0, 215.1, 48098, 51000, 0.35, 400.0, None,
     "VERIFY price & shares"),
    ("Advantage Energy", "AAV", None, None, None, None, None, None, None, None,
     "NOT POPULATED — add from filings"),
]
for nm, tk, px, sh, nd, q2, p26, lq, aff, tp, st in comps:
    ws.cell(r, 1, nm).font = BOLD if tk == "POU" else BLACK
    ws.cell(r, 2, tk).font = BOLD if tk == "POU" else BLACK
    for col, val, fmt in [(3, px, CUR1), (4, sh, NUM1), (6, nd, CUR),
                          (8, q2, NUM), (9, p26, NUM), (10, lq, PCT),
                          (11, aff, CUR), (12, tp, NUM)]:
        c = ws.cell(r, col, val if val is not None else "n/a")
        c.font = BLUE if val is not None else SUB
        if val is not None: c.number_format = fmt
        else: c.alignment = Alignment(horizontal="right")
    ws.cell(r, 5, f'=IF(N(C{r})*N(D{r})=0,"n/a",C{r}*D{r})').font = BLACK
    ws.cell(r, 5).number_format = CUR
    ws.cell(r, 7, f'=IF(N(E{r})=0,"n/a",E{r}+N(F{r}))').font = BLACK
    ws.cell(r, 7).number_format = CUR
    for col, f_, fmt in [
        (13, f'=IFERROR(G{r}/K{r},"n/a")', MULT),
        (14, f'=IFERROR(G{r}*1000000/H{r},"n/a")', CUR),
        (15, f'=IFERROR(G{r}*1000000/I{r},"n/a")', CUR),
        (16, f'=IFERROR(G{r}*1000/L{r},"n/a")', CUR1),
    ]:
        c = ws.cell(r, col, f_); c.font = BLACK; c.number_format = fmt
        if tk == "POU": c.fill = YEL
    s = ws.cell(r, 17, st)
    s.font = RED if "VERIFY" in st or "NOT" in st else SUB
    if tk == "POU":
        for cc in list(range(1, 13)) + [17]: ws.cell(r, cc).fill = GRN
    r += 1
last = r - 1
for lab, fn in [("Peer median (excl. POU)", "MEDIAN"), ("Peer mean (excl. POU)", "AVERAGE")]:
    ws.cell(r, 1, lab).font = BOLD
    for col in [13, 14, 15, 16]:
        L = get_column_letter(col)
        c = ws.cell(r, col, f'=IFERROR({fn}({L}{first+1}:{L}{last}),"n/a")')
        c.font = BOLD; c.border = TOP
        c.number_format = MULT if col == 13 else (CUR1 if col == 16 else CUR)
    r += 1
med_row = r - 2
r += 1
ws.cell(r, 1, "MEDIAN and AVERAGE ignore text entries, so unpopulated peers drop out automatically rather than counting as zero.").font = SUB
r += 1
ws.cell(r, 1, "2P reserves are only populated for Paramount. Add peer reserve figures from each company's Annual Information Form to make the EV/2P column meaningful.").font = RED

# =====================================================================
# TRANSACTION COMPS
# =====================================================================
ws = wb.create_sheet("Transaction_Comps")
title(ws, "Transaction Comparables — What Acquirers Actually Paid",
      "Control transactions in the Canadian Montney and Duvernay. These carry a control premium that trading comps do not.")
W(ws, {"A": 20, "B": 20, "C": 12, "D": 13, "E": 13, "F": 12, "G": 12, "H": 13, "I": 12, "J": 44})
thdrs = ["Acquirer", "Target", "Announced", "Total value (C$mm)", "Production (Boe/d)",
         "2P (MBoe)", "Premium", "EV/flowing", "EV/2P", "Note"]
r = 4
for i, h in enumerate(thdrs):
    c = ws.cell(r, 1+i, h); c.font = BOLDW; c.fill = HDR
    c.alignment = Alignment(horizontal="right" if i > 2 else "left", wrap_text=True, vertical="center")
ws.row_dimensions[r].height = 30
r += 1
tfirst = r
deals = [
    ("Shell plc", "ARC Resources", "Apr 2026", 22000, 370000, 2000000, 0.27,
     "C$32.80 per share, 75% Shell stock and 25% cash. 1.5mm net Montney acres. Shell disclosed that ~40% liquids generated ~70% of revenue."),
    ("Ovintiv", "NuVista Energy", "Feb 2026", 3800, None, None, None,
     "Closed 3-Feb-2026. Added ~930 net well locations and ~140,000 net acres in the Alberta Montney. VERIFY production to compute multiples."),
]
for aq, tg, dt, val, prod, tp, prem, note in deals:
    ws.cell(r, 1, aq).font = BLACK
    ws.cell(r, 2, tg).font = BOLD
    ws.cell(r, 3, dt).font = BLACK
    for col, v, fmt in [(4, val, CUR), (5, prod, NUM), (6, tp, NUM), (7, prem, PCT)]:
        c = ws.cell(r, col, v if v is not None else "n/a")
        c.font = BLUE if v is not None else RED
        if v is not None: c.number_format = fmt
        else: c.alignment = Alignment(horizontal="right")
    for col, f_, fmt in [(8, f'=IFERROR(D{r}*1000000/E{r},"n/a")', CUR),
                         (9, f'=IFERROR(D{r}*1000/F{r},"n/a")', CUR1)]:
        c = ws.cell(r, col, f_); c.font = BLACK; c.number_format = fmt; c.fill = ORG
    n = ws.cell(r, 10, note); n.font = SUB
    n.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 46
    r += 1
r += 1
r = sec(ws, r, "PARAMOUNT AGAINST THE SHELL / ARC BENCHMARK", 9)
ws.cell(r, 1, "Metric").font = BOLDW
for i, h in enumerate(["Paramount", "Shell / ARC", "Paramount vs deal"]):
    c = ws.cell(r, 2+i, h); c.font = BOLDW; c.alignment = Alignment(horizontal="right")
for c in range(1, 10): ws.cell(r, c).fill = HDR
r += 1
bench_start = r
for lab, pou_f, deal_f, fmt, note in [
    ("EV per flowing Boe/d — Q2 2026 actual", "=Trading_Comps!N5", f"=H{tfirst}", CUR,
     "Paramount looks expensive on today's production."),
    ("EV per flowing Boe/d — 2026E average", "=Trading_Comps!O5", f"=H{tfirst}", CUR,
     "The gap narrows as guided volumes arrive."),
    ("EV per 2P Boe", "=Trading_Comps!P5", f"=I{tfirst}", CUR1,
     "And inverts entirely on reserves — Paramount looks cheap."),
]:
    ws.cell(r, 1, lab).font = BOLD
    c = ws.cell(r, 2, pou_f); c.font = GREEN; c.number_format = fmt
    c = ws.cell(r, 3, deal_f); c.font = GREEN; c.number_format = fmt
    c = ws.cell(r, 4, f'=IFERROR(B{r}/C{r}-1,"n/a")'); c.font = BOLD; c.number_format = '+0%;-0%'
    c.fill = YEL
    n = ws.cell(r, 6, note); n.font = SUB
    r += 1
r += 1
c = ws.cell(r, 1, "THE TENSION: Paramount trades at a large premium to the Shell/ARC deal on current production, yet at a discount on 2P reserves. "
                  "Both are true, and the reconciliation is the investment thesis. Production is about to double while the reserve book already contains "
                  "Sinclair and the Willesden Green inventory. Buying Paramount means paying up for barrels that exist on paper but are not yet flowing. "
                  "Whether that is cheap depends entirely on execution.")
c.font = BLACK
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
c.alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[r].height = 62

# =====================================================================
# VALUATION BRIDGE
# =====================================================================
ws = wb.create_sheet("Valuation_Bridge")
title(ws, "Implied Value per Share — Multiple Methods",
      "Applies peer and transaction multiples to Paramount's own metrics, then compares against Module 4's NAV.")
W(ws, {"A": 44, "B": 14, "C": 14, "D": 14, "E": 3, "F": 50})
r = 4
r = sec(ws, r, "PARAMOUNT INPUTS", 6)
inp = r
for lab, f_, fmt, note in [
    ("Shares outstanding (mm)", "=Trading_Comps!D5", NUM1, None),
    ("Net cash and investments (C$mm)", "=-Trading_Comps!F5", CUR, "Added back to bridge from enterprise value to equity."),
    ("Q2 2026 production (Boe/d)", "=Trading_Comps!H5", NUM, None),
    ("2026E production (Boe/d)", "=Trading_Comps!I5", NUM, None),
    ("2027E production (Boe/d)", 63400, NUM, "MODULE 2 output — the year Sinclair starts contributing."),
    ("Annualised Q2 2026 AFF (C$mm)", "=Trading_Comps!K5", CUR, None),
    ("2027E AFF (C$mm)", 950, CUR, "MODULE 2 Base case."),
    ("2P reserves (MBoe)", "=Trading_Comps!L5", NUM, None),
    ("Current share price (C$)", "=Trading_Comps!C5", CUR1, None),
]:
    ws.cell(r, 1, lab).font = BOLD
    c = ws.cell(r, 2, f_)
    c.font = GREEN if isinstance(f_, str) else BLUE
    c.number_format = fmt
    if note: n = ws.cell(r, 6, note); n.font = SUB
    r += 1
SH, CASH, Q2P, P26, P27, AFF2, AFF27, TWOP, PX = (f"$B${inp+i}" for i in range(9))
r += 1

r = sec(ws, r, "IMPLIED SHARE PRICE BY METHOD", 6)
ws.cell(r, 1, "Method").font = BOLDW
for i, h in enumerate(["Multiple", "Implied EV (C$mm)", "Implied C$/share"]):
    c = ws.cell(r, 2+i, h); c.font = BOLDW; c.alignment = Alignment(horizontal="right")
for c in range(1, 7): ws.cell(r, c).fill = HDR
r += 1
mstart = r
methods = [
    ("Shell/ARC EV per flowing Boe/d, on 2027E volumes",
     "=Transaction_Comps!H5", f"=B{{r}}*{P27}/1000000", CUR,
     "The most defensible single comparison: a control transaction in the same basin, in the same year."),
    ("Shell/ARC EV per 2P Boe",
     "=Transaction_Comps!I5", f"=B{{r}}*{TWOP}/1000", CUR,
     "Rewards Paramount's large undeveloped reserve book."),
    ("Peer median EV/AFF, on 2027E adjusted funds flow",
     f"=Trading_Comps!M{med_row}", f"=B{{r}}*{AFF27}", MULT,
     "Cash-flow based. Check the peer inputs before trusting this."),
    ("Peer median EV per flowing Boe/d, on 2027E volumes",
     f"=Trading_Comps!N{med_row}", f"=B{{r}}*{P27}/1000000", CUR,
     "Minority-stake valuation, so no control premium."),
]
for lab, mult_f, ev_f, mfmt, note in methods:
    ws.cell(r, 1, lab).font = BLACK
    c = ws.cell(r, 2, mult_f); c.font = GREEN; c.number_format = mfmt
    c = ws.cell(r, 3, ev_f.format(r=r)); c.font = BLACK; c.number_format = CUR
    c = ws.cell(r, 4, f"=(C{r}+{CASH})/{SH}"); c.font = BOLD; c.number_format = CUR1; c.fill = YEL
    n = ws.cell(r, 6, note); n.font = SUB
    n.alignment = Alignment(wrap_text=True, vertical="top")
    if len(note) > 70: ws.row_dimensions[r].height = 28
    r += 1
mend = r - 1
for lab, fn in [("Low", "MIN"), ("Median", "MEDIAN"), ("High", "MAX")]:
    ws.cell(r, 1, lab).font = BOLD
    c = ws.cell(r, 4, f'=IFERROR({fn}(D{mstart}:D{mend}),"n/a")')
    c.font = BOLD; c.number_format = CUR1; c.border = TOP
    r += 1
r += 1

r = sec(ws, r, "TRIANGULATION — the three independent routes to a value", 6)
ws.cell(r, 1, "Method").font = BOLDW
for i, h in enumerate(["C$/share", "vs market"]):
    c = ws.cell(r, 2+i, h); c.font = BOLDW; c.alignment = Alignment(horizontal="right")
for c in range(1, 7): ws.cell(r, c).fill = HDR
r += 1
for lab, f_, note in [
    ("Module 4 — sum-of-the-parts NAV", 31.65,
     "Bottom-up asset valuation. Hardcoded here; refresh if the NAV assumptions change."),
    ("Module 5 — comparables median", f"=D{mend+2}",
     "What the market and acquirers are paying for similar barrels."),
    ("Market price today", f"={PX}", "What you can actually buy it for."),
]:
    ws.cell(r, 1, lab).font = BOLD
    c = ws.cell(r, 2, f_)
    c.font = BLUE if isinstance(f_, float) else GREEN
    c.number_format = CUR1; c.fill = YEL
    c = ws.cell(r, 3, f'=IFERROR(B{r}/{PX}-1,"n/a")'); c.font = BOLD; c.number_format = '+0%;-0%'
    n = ws.cell(r, 6, note); n.font = SUB
    r += 1
r += 2
for t in [
    "HOW TO USE THIS: if NAV and comparables both land near the market price, the stock is fairly valued and there is no thesis. "
    "If they diverge, the divergence is the thesis — and you must be able to explain why the market disagrees with you.",
    "The most common error at this stage is quietly discarding whichever method gives the inconvenient answer. Show all of them, "
    "weight them explicitly, and say which you trust most and why.",
]:
    c = ws.cell(r, 1, t); c.font = BLACK
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 40; r += 1

for s in wb.worksheets:
    s.sheet_view.showGridLines = False
wb.save("/home/claude/pou/POU_Comps_v1.xlsx")
print("saved")
