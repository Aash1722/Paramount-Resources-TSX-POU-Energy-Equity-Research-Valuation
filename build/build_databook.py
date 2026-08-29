"""
Paramount Resources (TSX:POU) - PE Take-Private Case Study
Module 1: Data Book (source of truth for all downstream modules)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ---------- Style constants ----------
FONT = "Arial"
BLUE = Font(name=FONT, size=10, color="0000FF")            # hardcoded input
BLACK = Font(name=FONT, size=10)                            # formula
GREEN = Font(name=FONT, size=10, color="008000")            # cross-sheet link
BOLD = Font(name=FONT, size=10, bold=True)
BOLD_W = Font(name=FONT, size=10, bold=True, color="FFFFFF")
TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
SUB = Font(name=FONT, size=10, italic=True, color="595959")
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SEC_FILL = PatternFill("solid", fgColor="D9E2F3")
YEL = PatternFill("solid", fgColor="FFFF00")
GREY = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
TOPLINE = Border(top=Side(style="thin", color="404040"))

CUR = '$#,##0;($#,##0);-'
CUR1 = '$#,##0.00;($#,##0.00);-'
NUM = '#,##0;(#,##0);-'
NUM1 = '#,##0.0;(#,##0.0);-'
PCT = '0.0%'
MULT = '0.00"x"'

def sec(ws, row, text, width=8):
    """Section banner."""
    ws.cell(row, 1, text).font = Font(name=FONT, size=10, bold=True, color="1F3864")
    for c in range(1, width + 1):
        ws.cell(row, c).fill = SEC_FILL
    return row + 1

def title(ws, name, subtitle):
    ws["A1"] = name
    ws["A1"].font = TITLE
    ws["A2"] = subtitle
    ws["A2"].font = SUB
    ws.freeze_panes = "A4"

def put(ws, r, c, v, font=BLACK, fmt=None, align=None, fill=None):
    cell = ws.cell(r, c, v)
    cell.font = font
    if fmt: cell.number_format = fmt
    if align: cell.alignment = Alignment(horizontal=align)
    if fill: cell.fill = fill
    return cell

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


# =====================================================================
# 1. README
# =====================================================================
ws = wb.active
ws.title = "README"
title(ws, "Paramount Resources Ltd. (TSX: POU) — PE Take-Private Case Study",
      "Module 1: Data Book. Every figure below is sourced from public disclosure. All downstream models reference this file.")
widths(ws, {"A": 34, "B": 96})

r = 4
r = sec(ws, r, "PURPOSE", 2)
for label, text in [
    ("This workbook", "Single source of truth. Historical financials, operating data, reserves, guidance, capital structure and the price deck."),
    ("What it is not", "It contains no forecast of the business. The operating model, NAV, LBO and Monte Carlo modules are built separately and link here."),
    ("Rule", "No downstream module should hardcode a POU figure. If a number is needed, add it here first with its source, then reference it."),
]:
    put(ws, r, 1, label, BOLD); put(ws, r, 2, text); ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30
    r += 1

r += 1
r = sec(ws, r, "COLOUR LEGEND (financial modelling convention)", 2)
for colr, meaning in [
    ("Blue text", "Hardcoded input taken from a source document, or a scenario lever you may change."),
    ("Black text", "Formula calculated within the sheet."),
    ("Green text", "Link to another sheet in this workbook."),
    ("Yellow fill", "Key assumption. These are the cells that drive the case — review before every run."),
]:
    f = BLUE if colr == "Blue text" else (GREEN if colr == "Green text" else BOLD)
    put(ws, r, 1, colr, f)
    if colr == "Yellow fill": ws.cell(r, 1).fill = YEL
    put(ws, r, 2, meaning)
    r += 1

r += 1
r = sec(ws, r, "TAB GUIDE", 2)
for tab, desc in [
    ("Snapshot", "Market data, capitalisation, enterprise value, ownership. The two blue cells at top are the only market inputs to refresh."),
    ("Financials_Hist", "Quarterly and annual income, cash flow and balance sheet items, FY2024 through Q2 2026."),
    ("Production_Ops", "Sales volumes by asset and product type; realised prices; netback build; unit operating costs."),
    ("Reserves", "McDaniel report effective 31-Dec-2025: PDP / TP / P+P volumes, before-tax NPV, reserve life, F&D and recycle ratios."),
    ("Guidance", "Company guidance for 2026 and 2027, including the capital programme by asset and the 2027 exit-rate target."),
    ("Assets", "Asset-by-asset descriptors: land, processing capacity, egress, well results, plateau targets. Inputs to the type-curve work."),
    ("Hedges_Credit", "Hedge book and credit facilities as disclosed at Q2 2026."),
    ("PriceDeck", "Bear / Base / Bull commodity assumptions. THE central scenario lever — the Monte Carlo module samples around these."),
    ("Derived_Metrics", "Live trading multiples calculated off Snapshot, Reserves and Production_Ops."),
]:
    put(ws, r, 1, tab, BOLD); put(ws, r, 2, desc)
    r += 1

r += 1
r = sec(ws, r, "PRIMARY SOURCES", 2)
for s, u in [
    ("Q2 2026 results (6-Aug-2026)", "paramount.mediaroom.com — Q2 2026 press release. Volumes, netbacks, capex by asset, hedges, guidance."),
    ("FY2025 results (3-Mar-2026)", "paramount.mediaroom.com — Q4/annual 2025 press release. Reserves tables, FY financials, 2026-27 capital outlook."),
    ("2025 Annual Information Form", "paramountres.com — full McDaniel reserves disclosure, land, risk factors."),
    ("SEDAR+", "sedarplus.ca — financial statements and MD&A."),
    ("Market data", "TSX close and consensus targets. Refresh the two blue cells on Snapshot before each run."),
]:
    put(ws, r, 1, s, BOLD); put(ws, r, 2, u)
    r += 1

r += 2
put(ws, r, 1, "Hypothetical case study for analytical purposes. Not investment advice, and not a recommendation to buy or sell any security.", SUB)


# =====================================================================
# 2. SNAPSHOT
# =====================================================================
ws = wb.create_sheet("Snapshot")
title(ws, "Company Snapshot & Capitalisation",
      "Balance sheet items as at 30-Jun-2026 (Q2 2026 press release). Market data as at the date shown — refresh the blue cells.")
widths(ws, {"A": 46, "B": 16, "C": 12, "D": 60})

r = 4
r = sec(ws, r, "MARKET INPUTS — refresh these two cells", 4)
put(ws, r, 1, "Market data as at", BOLD); put(ws, r, 2, "21-Aug-2026", BLUE, align="right", fill=YEL)
put(ws, r, 4, "Update whenever you re-run the case.", SUB); r += 1
put(ws, r, 1, "Share price (C$/share)", BOLD); put(ws, r, 2, 32.65, BLUE, CUR1, fill=YEL)
put(ws, r, 4, "TSX close. Implied by consensus target of ~C$37.10 quoted at +13.7%.", SUB); r += 1
put(ws, r, 1, "Consensus analyst target (C$/share)", BOLD); put(ws, r, 2, 37.10, BLUE, CUR1)
put(ws, r, 4, "10 analysts. Recent targets: RBC C$37, Scotiabank C$36, CIBC C$34.", SUB); r += 2

r = sec(ws, r, "SHARE COUNT & MARKET CAPITALISATION", 4)
put(ws, r, 1, "Common shares outstanding (mm)", BOLD); put(ws, r, 2, 145.8, BLUE, NUM1)
put(ws, r, 4, "Q2 2026, net of 0.1mm shares held in trust under the RSU plan.", SUB); r += 1
put(ws, r, 1, "Market capitalisation (C$mm)", BOLD); put(ws, r, 2, "=B6*B10", BLACK, CUR); r += 2

r = sec(ws, r, "NET DEBT / (NET CASH) BRIDGE — as at 30-Jun-2026", 4)
put(ws, r, 1, "Cash and cash equivalents (C$mm)", BOLD); put(ws, r, 2, 449.0, BLUE, CUR)
put(ws, r, 4, "Q2 2026 press release.", SUB); r += 1
put(ws, r, 1, "Long-term debt (C$mm)", BOLD); put(ws, r, 2, 0.0, BLUE, CUR)
put(ws, r, 4, "No long-term debt outstanding.", SUB); r += 1
put(ws, r, 1, "Net (cash) debt as reported (C$mm)", BOLD); put(ws, r, 2, -329.1, BLUE, CUR)
put(ws, r, 4, "As reported; includes working capital items, so differs from cash less debt.", SUB); r += 1
put(ws, r, 1, "Investments in securities (C$mm)", BOLD); put(ws, r, 2, 208.0, BLUE, CUR)
put(ws, r, 4, "At 30-Jun-26. AKITA shares received on the Fox Drilling sale were distributed to shareholders 16-Jul-26 — confirm the residual before relying on this.", SUB)
ws.cell(r, 4).alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[r].height = 28; r += 1
put(ws, r, 1, "Undrawn credit facilities (C$mm)", BOLD); put(ws, r, 2, 750.0, BLUE, CUR)
put(ws, r, 4, "C$500mm revolver (matures 15-Dec-2029) plus C$250mm EDC delayed-draw term loan.", SUB); r += 2

r = sec(ws, r, "ENTERPRISE VALUE", 4)
ev_r = r
put(ws, r, 1, "Market capitalisation", BOLD); put(ws, r, 2, "=B11", BLACK, CUR); r += 1
put(ws, r, 1, "Plus: net debt / (less: net cash)", BOLD); put(ws, r, 2, "=B16", BLACK, CUR); r += 1
put(ws, r, 1, "Enterprise value — before securities (C$mm)", BOLD); put(ws, r, 2, f"=SUM(B{ev_r}:B{ev_r+1})", BOLD, CUR)
ws.cell(r, 2).border = TOPLINE; r += 1
put(ws, r, 1, "Less: investments in securities", BOLD); put(ws, r, 2, "=-B17", BLACK, CUR); r += 1
put(ws, r, 1, "Enterprise value — fully adjusted (C$mm)", BOLD); put(ws, r, 2, f"=B{ev_r+2}+B{ev_r+3}", BOLD, CUR)
ws.cell(r, 2).border = TOPLINE
SNAP_EV = r
put(ws, r, 4, "Used as the EV basis in Derived_Metrics.", SUB); r += 2

r = sec(ws, r, "OWNERSHIP — the binding constraint on any take-private", 4)
put(ws, r, 1, "Holder", BOLD_W); put(ws, r, 2, "% held", BOLD_W, align="right"); put(ws, r, 4, "Note", BOLD_W)
for c in [1, 2, 3, 4]: ws.cell(r, c).fill = HDR_FILL
r += 1
own = [
    ("James H.T. Riddell (Chairman, President & CEO)", 0.3283, "~47.1mm shares."),
    ("Susan Riddell Rose (Director)", 0.1044, "~15.0mm shares."),
    ("Brenda Riddell (Director)", 0.1036, "~14.9mm shares."),
]
first_own = r
for n, p, note in own:
    put(ws, r, 1, n); put(ws, r, 2, p, BLUE, PCT); put(ws, r, 4, note, SUB); r += 1
put(ws, r, 1, "Riddell family — combined", BOLD); put(ws, r, 2, f"=SUM(B{first_own}:B{r-1})", BOLD, PCT)
ws.cell(r, 2).border = TOPLINE
put(ws, r, 4, "A controlling block. No take-private can proceed without it.", Font(name=FONT, size=10, italic=True, color="C00000")); r += 1
fam_row = r - 1
put(ws, r, 1, "Institutions", BLACK); put(ws, r, 2, 0.0802, BLUE, PCT)
put(ws, r, 4, "Unusually low for a C$4bn+ issuer.", SUB)
inst_row = r; r += 1
put(ws, r, 1, "Public companies and individual investors", BLACK)
put(ws, r, 2, f"=1-B{fam_row}-B{inst_row}", BLACK, PCT); r += 2

put(ws, r, 1, "IMPLICATION", Font(name=FONT, size=10, bold=True, color="C00000"))
put(ws, r, 2, "A hostile take-private is not available. Any transaction is negotiated with the founding family, and under Canadian MI 61-101 a related-party rollover triggers a formal valuation and majority-of-the-minority approval. Test this before modelling anything else.", BLACK)
ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
ws.row_dimensions[r].height = 42


# =====================================================================
# 3. FINANCIALS_HIST
# =====================================================================
ws = wb.create_sheet("Financials_Hist")
title(ws, "Historical Financial Results",
      "C$mm unless noted. Source: Q2 2026 and FY2025 press releases. Note the Grande Prairie disposition closed 31-Jan-2025 — 2024 is not comparable.")
widths(ws, {"A": 42, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 52})

cols = ["FY2024", "FY2025", "Q2 2025", "Q4 2025", "Q1 2026", "Q2 2026"]
r = 4
put(ws, r, 1, "C$ millions", BOLD_W)
for i, c in enumerate(cols):
    put(ws, r, 2 + i, c, BOLD_W, align="right")
put(ws, r, 8, "Note", BOLD_W)
for c in range(1, 9): ws.cell(r, c).fill = HDR_FILL
r += 1

data = [
    ("SEC", "EARNINGS & CASH FLOW", None, None),
    ("N", "Net income / (loss)", [335.9, 1288.7, 4.2, -1.9, 53.2, 68.3], "FY2025 inflated by the C$3.24bn Grande Prairie disposition gain."),
    ("N", "Cash from operating activities", [815.3, 417.3, 39.8, 185.4, 116.2, 144.4], None),
    ("N", "Adjusted funds flow (AFF)", [930.3, 467.2, 81.5, 140.1, 143.4, 175.6], "Company capital management measure. Closest public proxy for cash EBITDA."),
    ("N", "Free cash flow", [37.3, -385.5, -85.5, -84.6, -146.9, -128.5], "Negative for six straight quarters — the company is funding a build-out."),
    ("SEC", "CAPITAL EXPENDITURE BY ASSET", None, None),
    ("N", "Willesden Green", [233.5, 569.6, 106.2, 158.3, 161.7, 186.8], None),
    ("N", "Sinclair", [14.5, 64.9, 5.0, 35.0, 59.6, 99.7], "Ramping hard ahead of Q4 2027 plant start-up."),
    ("N", "Kaybob", [172.6, 121.3, 40.2, 20.8, 31.8, 3.5], "Now in maintenance mode."),
    ("N", "Fox Drilling", [8.4, 8.5, 1.6, 2.2, 1.5, 4.3], "Subsidiary sold to AKITA 30-Jun-2026."),
    ("N", "Corporate, other and sold assets", [413.2, 24.2, 4.6, -7.7, 2.3, 2.4], "FY2024 includes C$414.6mm on assets since sold."),
    ("T", "Total capital expenditures", None, None),
    ("N", "Asset retirement obligations settled", [38.1, 39.0, 3.0, 9.4, 26.5, 1.5], None),
    ("SEC", "BALANCE SHEET", None, None),
    ("N", "Total assets", [4757.5, 3587.2, 3517.9, None, 3746.7, 3832.1], None),
    ("N", "Investments in securities", [563.9, 137.3, 575.3, None, 141.4, 208.0], None),
    ("N", "Long-term debt", [173.0, 0.0, 0.0, None, 0.0, 0.0], "Debt-free since the Grande Prairie disposition."),
    ("N", "Net (cash) debt", [188.4, -672.8, -500.9, None, -515.8, -329.1], "Net cash is being consumed by the capital programme."),
    ("N", "Common shares outstanding (mm)", [146.9, 144.2, 143.3, 144.2, 144.9, 145.8], None),
]

cap_rows = []
total_row = None
for kind, label, vals, note in data:
    if kind == "SEC":
        r = sec(ws, r, label, 8); continue
    if kind == "T":
        put(ws, r, 1, label, BOLD)
        for i in range(6):
            col = get_column_letter(2 + i)
            put(ws, r, 2 + i, f"=SUM({col}{cap_rows[0]}:{col}{cap_rows[-1]})", BOLD, CUR)
            ws.cell(r, 2 + i).border = TOPLINE
        put(ws, r, 8, "Ties to reported totals: FY24 842.2, FY25 788.5, Q2'26 296.7.", SUB)
        total_row = r; r += 1; continue
    put(ws, r, 1, label, BOLD if label.startswith("Total") else BLACK)
    for i, v in enumerate(vals):
        if v is None:
            put(ws, r, 2 + i, "n/a", SUB, align="right")
        else:
            fmt = NUM1 if "shares" in label else CUR
            put(ws, r, 2 + i, v, BLUE, fmt)
    if note: put(ws, r, 8, note, SUB)
    if label in ("Willesden Green", "Sinclair", "Kaybob", "Fox Drilling", "Corporate, other and sold assets"):
        cap_rows.append(r)
    if label == "Cash from operating activities": CFO_ROW = r
    if label == "Adjusted funds flow (AFF)": AFF_ROW = r
    r += 1

r += 1
r = sec(ws, r, "DERIVED", 8)
put(ws, r, 1, "AFF / cash from operating activities", BOLD)
for i in range(6):
    col = get_column_letter(2 + i)
    put(ws, r, 2 + i, f"=IFERROR({col}{AFF_ROW}/{col}{CFO_ROW},\"\")", BLACK, PCT)
r += 1
put(ws, r, 1, "Capital intensity (capex / AFF)", BOLD)
for i in range(6):
    col = get_column_letter(2 + i)
    put(ws, r, 2 + i, f"=IFERROR({col}{total_row}/{col}{AFF_ROW},\"\")", BLACK, MULT)
put(ws, r, 8, "Above 1.0x means outspend. This is the single most important fact about POU today.",
    Font(name=FONT, size=10, italic=True, color="C00000"))


# =====================================================================
# 4. PRODUCTION_OPS
# =====================================================================
ws = wb.create_sheet("Production_Ops")
title(ws, "Production, Realised Pricing & Netbacks",
      "Sales volumes and unit economics. Source: Q2 2026 and FY2025 press releases.")
widths(ws, {"A": 42, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 52})

r = 4
put(ws, r, 1, "", BOLD_W)
for i, c in enumerate(cols):
    put(ws, r, 2 + i, c, BOLD_W, align="right")
put(ws, r, 8, "Note", BOLD_W)
for c in range(1, 9): ws.cell(r, c).fill = HDR_FILL
r += 1

prod = [
    ("SEC", "SALES VOLUMES BY PRODUCT", None, None, None),
    ("N", "Natural gas (MMcf/d)", [306.8, 131.9, 103.3, 133.1, 144.5, 144.4], NUM1, None),
    ("N", "Condensate and oil (Bbl/d)", [40432, 16402, 11636, 19472, 18137, 17803], NUM, None),
    ("N", "Other NGLs (Bbl/d)", [6920, 3853, 2786, 5318, 6037, 5415], NUM, "Ethane, propane, butane."),
    ("C", "Total (Boe/d)", None, NUM, "6 Mcf : 1 Bbl. Company notes the 2026 value ratio is closer to 60:1."),
    ("F", "% liquids", None, PCT, None),
    ("SEC", "SALES VOLUMES BY ASSET (Boe/d)", None, None, None),
    ("N", "Willesden Green", [7537, 14161, 9223, 25752, 28750, 27466], NUM, "From ~7 kboe/d in Jan-25 to >29 kboe/d in Dec-25. The growth engine."),
    ("N", "Kaybob", [22404, 21216, 21962, 20387, 19088, 19413], NUM, "Maintenance: company guides 19-20 kboe/d through 2028."),
    ("N", "Other properties", [1226, 770, 446, 834, 417, 400], NUM, None),
    ("N", "Sold assets (Karr, Wapiti, Zama)", [67323, 6091, 0, 0, 0, 0], NUM, "Disposed 31-Jan-2025 for C$3.243bn."),
    ("T2", "Total (Boe/d)", None, NUM, None),
    ("SEC", "REALISED PRICING", None, None, None),
    ("N", "Natural gas (C$/Mcf)", [1.99, 3.02, 3.07, 3.58, 3.52, 2.50], CUR1, "~48% of remaining-2026 gas priced ex-AECO at Dawn, Malin, Emerson."),
    ("N", "Condensate and oil (C$/Bbl)", [96.96, 85.40, 82.84, 76.66, 96.27, 126.96], CUR1, "Condensate is the value driver, not gas."),
    ("N", "Other NGLs (C$/Bbl)", [35.37, 30.46, 27.02, 27.15, 30.61, 38.50], CUR1, None),
    ("SEC", "NETBACK BUILD (C$/Boe)", None, None, None),
    ("N", "Petroleum and natural gas sales", [48.83, 47.77, 44.20, 45.92, 52.65, 60.59], CUR1, None),
    ("N", "Royalties", [-6.18, -3.31, -2.00, -2.61, -2.90, -3.25], CUR1, "Low: new-well royalty rates on Duvernay volumes."),
    ("N", "Operating expense", [-13.15, -11.66, -12.39, -9.84, -9.81, -8.66], CUR1, "Five consecutive quarters of decline as owned facilities absorb volume."),
    ("N", "Transportation and NGLs processing", [-3.76, -4.47, -4.57, -4.81, -4.83, -4.75], CUR1, None),
    ("N", "Net commodities purchased / (sold)", [0.15, 0.33, 0.11, 0.26, 0.11, 0.24], CUR1, "Corporate marketing item, roughly nets to zero."),
    ("T3", "Netback", None, CUR1, "Ties to reported: FY24 25.89, FY25 28.66, Q2'26 44.17."),
    ("N", "Risk management settlements", [1.01, 3.30, 5.16, 4.73, -0.46, -2.48], CUR1, None),
    ("T4", "Netback incl. risk management", None, CUR1, None),
    ("SEC", "ASSET-LEVEL UNIT COSTS", None, None, None),
    ("N", "Willesden Green operating expense (C$/Boe)", [None, 6.26, None, 4.71, None, None], CUR1, "The structural cost advantage. Company-owned Alhambra plant."),
]

vol_rows, nb_rows = [], []
tot_vol_row = nb_row = nb_incl_row = None
for item in prod:
    kind, label, vals, fmt, note = item
    if kind == "SEC":
        r = sec(ws, r, label, 8); continue
    if kind == "C":
        put(ws, r, 1, label, BOLD)
        for i in range(6):
            col = get_column_letter(2 + i)
            put(ws, r, 2 + i, f"={col}{r-3}*1000/6+{col}{r-2}+{col}{r-1}", BOLD, NUM)
            ws.cell(r, 2 + i).border = TOPLINE
        tot_vol_row = r; put(ws, r, 8, note, SUB); r += 1; continue
    if kind == "F":
        put(ws, r, 1, label, BLACK)
        for i in range(6):
            col = get_column_letter(2 + i)
            put(ws, r, 2 + i, f"=({col}6+{col}7)/{col}{tot_vol_row}", BLACK, PCT)
        r += 1; continue
    if kind == "T2":
        put(ws, r, 1, label, BOLD)
        for i in range(6):
            col = get_column_letter(2 + i)
            put(ws, r, 2 + i, f"=SUM({col}{vol_rows[0]}:{col}{vol_rows[-1]})", BOLD, NUM)
            ws.cell(r, 2 + i).border = TOPLINE
        put(ws, r, 8, "Cross-check: must equal the by-product total above.", SUB); r += 1; continue
    if kind == "T3":
        put(ws, r, 1, label, BOLD)
        for i in range(6):
            col = get_column_letter(2 + i)
            put(ws, r, 2 + i, f"=SUM({col}{nb_rows[0]}:{col}{nb_rows[-1]})", BOLD, CUR1)
            ws.cell(r, 2 + i).border = TOPLINE
        nb_row = r; put(ws, r, 8, note, SUB); r += 1; continue
    if kind == "T4":
        put(ws, r, 1, label, BOLD)
        for i in range(6):
            col = get_column_letter(2 + i)
            put(ws, r, 2 + i, f"={col}{nb_row}+{col}{r-1}", BOLD, CUR1)
            ws.cell(r, 2 + i).border = TOPLINE
        nb_incl_row = r; r += 1; continue

    put(ws, r, 1, label, BLACK)
    for i, v in enumerate(vals):
        if v is None:
            put(ws, r, 2 + i, "n/a", SUB, align="right")
        else:
            put(ws, r, 2 + i, v, BLUE, fmt)
    if note: put(ws, r, 8, note, SUB)
    if label in ("Willesden Green", "Kaybob", "Other properties", "Sold assets (Karr, Wapiti, Zama)"):
        vol_rows.append(r)
    if label in ("Petroleum and natural gas sales", "Royalties", "Operating expense",
                 "Transportation and NGLs processing", "Net commodities purchased / (sold)"):
        nb_rows.append(r)
    r += 1

r += 1
r = sec(ws, r, "DERIVED — ANNUALISED CASH NETBACK", 8)
put(ws, r, 1, "Netback annualised (C$mm)", BOLD)
for i in range(6):
    col = get_column_letter(2 + i)
    put(ws, r, 2 + i, f"={col}{nb_incl_row}*{col}{tot_vol_row}*365/1000000", BLACK, CUR)
put(ws, r, 8, "Field-level cash generation at the period's volumes and netback. Before G&A, interest and tax.", SUB)


# =====================================================================
# 5. RESERVES
# =====================================================================
ws = wb.create_sheet("Reserves")
title(ws, "Reserves — McDaniel Report, effective 31-Dec-2025",
      "Gross reserves. NPV is before tax and does not represent fair market value. Price forecast is the average of Sproule, GLJ and McDaniel decks.")
widths(ws, {"A": 44, "B": 16, "C": 16, "D": 16, "E": 4, "F": 54})

r = 4
r = sec(ws, r, "RESERVE VOLUMES", 4)
put(ws, r, 1, "", BOLD_W)
for i, c in enumerate(["PDP", "Total Proved (1P)", "Proved + Probable (2P)"]):
    put(ws, r, 2 + i, c, BOLD_W, align="right")
for c in range(1, 5): ws.cell(r, c).fill = HDR_FILL
r += 1
hdr_res = r - 1
for label, vals, fmt in [
    ("Natural gas (Bcf)", [202, 611, 2094], NUM),
    ("NGLs (MBbl)", [23492, 95890, 168336], NUM),
    ("Crude oil (MBbl)", [1988, 2290, 4243], NUM),
]:
    put(ws, r, 1, label, BLACK)
    for i, v in enumerate(vals): put(ws, r, 2 + i, v, BLUE, fmt)
    r += 1
put(ws, r, 1, "Total (MBoe)", BOLD)
for i in range(3):
    col = get_column_letter(2 + i)
    put(ws, r, 2 + i, f"={col}{r-3}*1000/6+{col}{r-2}+{col}{r-1}", BOLD, NUM)
    ws.cell(r, 2 + i).border = TOPLINE
tot_res_row = r
RES_VOL_ROW = r
put(ws, r, 6, "Reported: 59,151 / 199,989 / 521,518 MBoe.", SUB); r += 1
put(ws, r, 1, "% liquids", BLACK)
for i in range(3):
    col = get_column_letter(2 + i)
    put(ws, r, 2 + i, f"=({col}{r-3}+{col}{r-2})/{col}{tot_res_row}", BLACK, PCT)
put(ws, r, 6, "2P is gassier than 1P — Sinclair dry gas dominates the probable book.", SUB); r += 2

r = sec(ws, r, "BEFORE-TAX NET PRESENT VALUE OF FUTURE NET REVENUE (C$mm)", 4)
put(ws, r, 1, "", BOLD_W)
for i, c in enumerate(["Volumes (MBoe)", "Undiscounted", "Discounted at 10%"]):
    put(ws, r, 2 + i, c, BOLD_W, align="right")
for c in range(1, 5): ws.cell(r, c).fill = HDR_FILL
r += 1
put(ws, r, 1, "Total Proved (1P)", BOLD); r += 1
npv_start = r
for label, vals in [("Developed", [70266, 357, 729]), ("Undeveloped", [129723, 2485, 950])]:
    put(ws, r, 1, "   " + label, BLACK)
    for i, v in enumerate(vals): put(ws, r, 2 + i, v, BLUE, NUM if i == 0 else CUR)
    r += 1
put(ws, r, 1, "   Total Proved", BOLD)
for i in range(3):
    col = get_column_letter(2 + i)
    put(ws, r, 2 + i, f"=SUM({col}{npv_start}:{col}{r-1})", BOLD, NUM if i == 0 else CUR)
    ws.cell(r, 2 + i).border = TOPLINE
tp_npv10 = f"D{r}"
r += 1
put(ws, r, 1, "Proved plus Probable (2P)", BOLD); r += 1
npv2_start = r
for label, vals in [("Developed", [96568, 929, 972]), ("Undeveloped", [424951, 7401, 2278])]:
    put(ws, r, 1, "   " + label, BLACK)
    for i, v in enumerate(vals): put(ws, r, 2 + i, v, BLUE, NUM if i == 0 else CUR)
    r += 1
put(ws, r, 1, "   Total Proved plus Probable", BOLD)
for i in range(3):
    col = get_column_letter(2 + i)
    put(ws, r, 2 + i, f"=SUM({col}{npv2_start}:{col}{r-1})", BOLD, NUM if i == 0 else CUR)
    ws.cell(r, 2 + i).border = TOPLINE
pp_npv10 = f"D{r}"
PP_NPV10_ROW = r
pp_npv0 = f"C{r}"
put(ws, r, 6, "2P NPV10 of C$3,250mm is the anchor for the P/NAV test in Derived_Metrics.", SUB)
r += 2

r = sec(ws, r, "NPV10 PER SHARE — the central valuation tension", 4)
put(ws, r, 1, "2P NPV10, before tax (C$mm)", BOLD); put(ws, r, 2, f"={pp_npv10}", GREEN, CUR); r += 1
put(ws, r, 1, "Plus: net cash at 30-Jun-2026 (C$mm)", BOLD); put(ws, r, 2, "=-Snapshot!B16", GREEN, CUR); r += 1
put(ws, r, 1, "Adjusted 2P NAV (C$mm)", BOLD); put(ws, r, 2, f"=SUM(B{r-2}:B{r-1})", BOLD, CUR)
ws.cell(r, 2).border = TOPLINE; r += 1
put(ws, r, 1, "Shares outstanding (mm)", BOLD); put(ws, r, 2, "=Snapshot!B10", GREEN, NUM1); r += 1
put(ws, r, 1, "2P NAV per share (C$)", BOLD); put(ws, r, 2, f"=B{r-2}/B{r-1}", BOLD, CUR1)
ws.cell(r, 2).border = TOPLINE
put(ws, r, 6, "Excludes undeveloped land, infrastructure above reserve value and any resource beyond 2P.", SUB); r += 1
put(ws, r, 1, "Current share price (C$)", BOLD); put(ws, r, 2, "=Snapshot!B6", GREEN, CUR1); r += 1
put(ws, r, 1, "Price / 2P NAV", BOLD); put(ws, r, 2, f"=B{r-1}/B{r-2}", BOLD, MULT)
ws.cell(r, 2).fill = YEL
put(ws, r, 6, "Above 1.0x means the market already pays for value the engineered reserve report does not book.",
    Font(name=FONT, size=10, italic=True, color="C00000")); r += 2

r = sec(ws, r, "RESERVE LIFE, F&D AND RECYCLE RATIOS (2025)", 4)
put(ws, r, 1, "", BOLD_W)
for i, c in enumerate(["PDP", "Total Proved", "Proved + Probable"]):
    put(ws, r, 2 + i, c, BOLD_W, align="right")
for c in range(1, 5): ws.cell(r, c).fill = HDR_FILL
r += 1
for label, vals, fmt, note in [
    ("Reserve life index (years)", [4.5, 15.2, 39.5], NUM1, "PDP RLI of 4.5 years is short — this is a company that must keep drilling."),
    ("F&D cost (C$/Boe)", [24.42, 24.15, 11.67], CUR1, "Includes future development capital for Alhambra and Sinclair."),
    ("Recycle ratio (x)", [1.2, 1.2, 2.5], MULT, "1.2x on a proved basis is thin. The 2.5x on 2P assumes Sinclair is delivered."),
    ("2025 reserve replacement (x)", [2.4, 5.2, 21.7], MULT, None),
]:
    put(ws, r, 1, label, BLACK)
    for i, v in enumerate(vals): put(ws, r, 2 + i, v, BLUE, fmt)
    if note: put(ws, r, 6, note, SUB)
    r += 1


# =====================================================================
# 6. GUIDANCE
# =====================================================================
ws = wb.create_sheet("Guidance")
title(ws, "Company Guidance & Capital Plan",
      "Source: Q2 2026 press release (6-Aug-2026), updated from the March 2026 outlook.")
widths(ws, {"A": 44, "B": 15, "C": 15, "D": 15, "E": 4, "F": 58})

r = 4
r = sec(ws, r, "2026 GUIDANCE (revised upward at Q2)", 4)
put(ws, r, 1, "", BOLD_W)
for i, c in enumerate(["Low", "High", "Midpoint"]): put(ws, r, 2 + i, c, BOLD_W, align="right")
for c in range(1, 5): ws.cell(r, c).fill = HDR_FILL
r += 1
for label, lo, hi, fmt, note in [
    ("Q3 2026 sales volumes (Boe/d)", 50000, 53000, NUM, "47% liquids."),
    ("Q4 2026 sales volumes (Boe/d)", 60000, 63000, NUM, "49% liquids. Alhambra phase 2 at full rate."),
    ("2026 annual sales volumes (Boe/d)", 51000, 53000, NUM, "49% liquids. Raised 2,000 Boe/d at midpoint at Q2."),
    ("2026 capital expenditures (C$mm)", 1000, 1100, CUR, "Unchanged."),
    ("2026 abandonment & reclamation (C$mm)", 35, 35, CUR, None),
]:
    put(ws, r, 1, label, BLACK)
    put(ws, r, 2, lo, BLUE, fmt); put(ws, r, 3, hi, BLUE, fmt)
    put(ws, r, 4, f"=AVERAGE(B{r}:C{r})", BLACK, fmt)
    put(ws, r, 6, note or "", SUB)
    r += 1
r += 1

r = sec(ws, r, "2027 OUTLOOK", 4)
for label, lo, hi, fmt, note in [
    ("2027 annual sales volumes (Boe/d)", 60000, 65000, NUM, "50% liquids."),
    ("2027 exit rate (Boe/d)", 100000, 100000, NUM, "Stated as 'over 100,000', 35% liquids. Roughly double the 2026 average."),
    ("2027 capital expenditures (C$mm)", 1000, 1000, CUR, "Midpoint outlook."),
]:
    put(ws, r, 1, label, BLACK)
    put(ws, r, 2, lo, BLUE, fmt); put(ws, r, 3, hi, BLUE, fmt)
    put(ws, r, 4, f"=AVERAGE(B{r}:C{r})", BLACK, fmt)
    put(ws, r, 6, note, SUB)
    r += 1
r += 1

r = sec(ws, r, "CAPITAL ALLOCATION BY ASSET (C$mm, March 2026 outlook)", 4)
put(ws, r, 1, "", BOLD_W)
for i, c in enumerate(["2026", "2027"]): put(ws, r, 2 + i, c, BOLD_W, align="right")
for c in range(1, 5): ws.cell(r, c).fill = HDR_FILL
r += 1
alloc_start = r
for label, v26, v27, note in [
    ("Willesden Green", 630, 440, "Drilling to fill Alhambra phase 2, plus the Leafland interconnect."),
    ("Sinclair", 360, 440, "Plant construction and 15 wells in 2026, completions and tie-ins in 2027."),
    ("Kaybob and corporate", 110, 120, "Balancing figure to the stated midpoint."),
]:
    put(ws, r, 1, label, BLACK)
    put(ws, r, 2, v26, BLUE, CUR); put(ws, r, 3, v27, BLUE, CUR)
    put(ws, r, 6, note, SUB); r += 1
put(ws, r, 1, "Total", BOLD)
for i in range(2):
    col = get_column_letter(2 + i)
    put(ws, r, 2 + i, f"=SUM({col}{alloc_start}:{col}{r-1})", BOLD, CUR)
    ws.cell(r, 2 + i).border = TOPLINE
ALLOC_TOTAL = r
GUID_Q4_ROW = 7
r += 2

r = sec(ws, r, "THE FUNDING GAP — why this is not a conventional LBO candidate", 4)
put(ws, r, 1, "", BOLD_W)
for i, c in enumerate(["2026E", "2027E"]): put(ws, r, 2 + i, c, BOLD_W, align="right")
for c in range(1, 5): ws.cell(r, c).fill = HDR_FILL
r += 1
gap_start = r
put(ws, r, 1, "Adjusted funds flow — analyst placeholder (C$mm)", BOLD)
put(ws, r, 2, 700, BLUE, CUR, fill=YEL); put(ws, r, 3, 950, BLUE, CUR, fill=YEL)
put(ws, r, 6, "PLACEHOLDER. Replace with output of the operating model in Module 2. 2026E is roughly H1 annualised.", SUB)
ws.cell(r, 6).alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[r].height = 28
r += 1
put(ws, r, 1, "Less: capital expenditures", BLACK)
put(ws, r, 2, f"=-B{ALLOC_TOTAL}", GREEN, CUR); put(ws, r, 3, f"=-C{ALLOC_TOTAL}", GREEN, CUR)
r += 1
put(ws, r, 1, "Less: abandonment & reclamation", BLACK)
put(ws, r, 2, -35, BLUE, CUR); put(ws, r, 3, -35, BLUE, CUR)
r += 1
put(ws, r, 1, "Free cash flow before dividends (C$mm)", BOLD)
for i in range(2):
    col = get_column_letter(2 + i)
    put(ws, r, 2 + i, f"=SUM({col}{gap_start}:{col}{r-1})", BOLD, CUR)
    ws.cell(r, 2 + i).border = TOPLINE
r += 1
put(ws, r, 1, "Less: dividends at C$0.60/share annualised", BLACK)
put(ws, r, 2, "=-0.6*Snapshot!B10", GREEN, CUR); put(ws, r, 3, "=-0.6*Snapshot!B10", GREEN, CUR)
r += 1
put(ws, r, 1, "Cumulative funding requirement (C$mm)", BOLD)
put(ws, r, 2, f"=B{r-2}+B{r-1}", BOLD, CUR)
put(ws, r, 3, f"=B{r}+C{r-2}+C{r-1}", BOLD, CUR)
for i in range(2): ws.cell(r, 2 + i).border = TOPLINE; ws.cell(r, 2 + i).fill = YEL
put(ws, r, 6, "Compare against C$449mm cash plus C$750mm undrawn facilities. A sponsor's leverage competes with this programme for the same balance sheet.",
    Font(name=FONT, size=10, italic=True, color="C00000"))
ws.cell(r, 6).alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[r].height = 30


# =====================================================================
# 7. ASSETS
# =====================================================================
ws = wb.create_sheet("Assets")
title(ws, "Asset Descriptors",
      "Inputs to asset-level type curves and the NAV build. Source: Q2 2026 and FY2025 press releases.")
widths(ws, {"A": 40, "B": 30, "C": 30, "D": 30, "E": 30})

r = 4
put(ws, r, 1, "", BOLD_W)
for i, c in enumerate(["Willesden Green Duvernay", "Sinclair Montney", "Kaybob", "Other / legacy"]):
    put(ws, r, 2 + i, c, BOLD_W, align="left")
for c in range(1, 6): ws.cell(r, c).fill = HDR_FILL
r += 1

asset_rows = [
    ("Status", "Producing, ramping", "Under construction, first gas Q4 2027", "Producing, maintenance mode", "Undeveloped land bank"),
    ("Q2 2026 volumes (Boe/d)", "27,466 (55% liquids)", "Nil", "19,413 (40% liquids)", "400"),
    ("Net acres", "Over 350,000", "Over 170,000", "110,000 Duvernay + 180,000 Montney", "1.3mm heavy oil; 110k Muskwa; 195k Besa River; 170k NWT"),
    ("Processing capacity", "Alhambra 20,000 Bbl/d + 100 MMcf/d raw after phase 2; Leafland 6,000 Bbl/d + 22 MMcf/d",
     "Sinclair Plant designed for 400 MMcf/d raw", "Company-owned gathering and processing", "None"),
    ("Expansion optionality", "Phase 3 adds 50 MMcf/d + 10,000 Bbl/d; FID possible Q4 2026, start-up as early as H2 2028",
     "Land position supports a higher plateau than currently targeted", "None planned", "Requires capital and a commodity view"),
    ("Contracted egress", "In place", "335 MMcf/d firm from Q4 2027", "In place", "n/a"),
    ("Well results", "210-day peak ~1,240 Boe/d (56% liquids) across 16 wells; 13 of 16 paid out in ~8 months",
     "Test rates 34 MMcf/d upper bench, 25 MMcf/d lower bench — highest publicly recorded in Alberta Montney",
     "Duvernay wells offsetting decline", "n/a"),
    ("Plateau target", "~50,000 Boe/d sustainable over 20+ years", "Over 300 MMcf/d of sales gas", "19,000-20,000 Boe/d through 2028", "n/a"),
    ("Unit operating cost", "C$4.71/Boe in Q4 2025", "Designed as high-rate, low-cost dry gas", "Higher — legacy conventional mix", "n/a"),
    ("Key risk for the model", "Type-curve durability past 300 days; wells are choked, so declines are managed not natural",
     "Single-asset execution risk: plant on time and on budget, and gas price at start-up",
     "Decline once Duvernay drilling stops", "Zero value in most NAV cases; option value only"),
]
for label, *vals in asset_rows:
    put(ws, r, 1, label, BOLD)
    for i, v in enumerate(vals):
        c = put(ws, r, 2 + i, v, BLACK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = BOX
    ws.row_dimensions[r].height = 44
    r += 1

r += 1
r = sec(ws, r, "TOTAL LAND POSITION AT 31-DEC-2025 (thousand acres)", 5)
put(ws, r, 1, "Acreage assigned reserves", BOLD); put(ws, r, 2, 852, BLUE, NUM); put(ws, r, 3, 704, BLUE, NUM); r += 1
put(ws, r, 1, "Acreage not assigned reserves", BOLD); put(ws, r, 2, 3363, BLUE, NUM); put(ws, r, 3, 2349, BLUE, NUM); r += 1
put(ws, r, 1, "Total", BOLD)
put(ws, r, 2, f"=SUM(B{r-2}:B{r-1})", BOLD, NUM); put(ws, r, 3, f"=SUM(C{r-2}:C{r-1})", BOLD, NUM)
ws.cell(r, 2).border = TOPLINE; ws.cell(r, 3).border = TOPLINE
put(ws, r - 3, 2, "Gross", BOLD); put(ws, r - 3, 3, "Net", BOLD)


# =====================================================================
# 8. HEDGES_CREDIT
# =====================================================================
ws = wb.create_sheet("Hedges_Credit")
title(ws, "Hedge Book & Credit Facilities",
      "As disclosed in the Q2 2026 press release (6-Aug-2026).")
widths(ws, {"A": 34, "B": 22, "C": 26, "D": 24, "E": 46})

r = 4
r = sec(ws, r, "COMMODITY AND FX CONTRACTS", 5)
put(ws, r, 1, "Instrument", BOLD_W); put(ws, r, 2, "Notional", BOLD_W)
put(ws, r, 3, "Price / rate", BOLD_W); put(ws, r, 4, "Remaining term", BOLD_W); put(ws, r, 5, "Note", BOLD_W)
for c in range(1, 6): ws.cell(r, c).fill = HDR_FILL
r += 1
for row in [
    ("NYMEX WTI swap (sale)", "7,000 Bbl/d", "C$109.19/Bbl", "Jul 2026 - Dec 2026", "Roughly 39% of Q2 liquids volumes."),
    ("NYMEX WTI swap (sale)", "6,000 Bbl/d", "C$100.34/Bbl", "Jan 2027 - Dec 2027", "Coverage falls as volumes double — 2027 is largely unhedged."),
    ("Citygate / Malin basis swap", "10,000 MMBtu/d", "Citygate less US$0.97/MMBtu", "Jul 2026 - Oct 2028", "Basis only, no flat-price protection."),
    ("FX average rate forward (sale)", "US$10mm / month", "1.3810 CAD/USD", "Jul 2026 - Dec 2026", None),
    ("FX average rate forward (sale)", "US$10mm / month", "1.3680 CAD/USD", "Jan 2027 - Dec 2027", None),
]:
    for i, v in enumerate(row):
        if v is not None:
            c = put(ws, r, 1 + i, v, BLUE if i in (1, 2) else BLACK)
            if i == 4: c.font = SUB
    r += 1
r += 1
put(ws, r, 1, "Observation", Font(name=FONT, size=10, bold=True, color="C00000"))
put(ws, r, 2, "There is no flat-price natural gas hedge. Sinclair is a dry gas project reaching first production in Q4 2027 into an unhedged market. A sponsor with leverage would need a materially larger hedge book, which costs upside — model this explicitly in the LBO.", BLACK)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[r].height = 44
r += 2

r = sec(ws, r, "CREDIT FACILITIES", 5)
for label, v, note in [
    ("Senior secured revolving bank credit facility (C$mm)", 500, "Covenant-based. Maturity extended to 15-Dec-2029."),
    ("EDC non-revolving delayed-draw term loan (C$mm)", 250, "Five-year, non-amortising. Secured December 2025."),
    ("Total facilities (C$mm)", None, "Fully undrawn at 30-Jun-2026."),
    ("Cash and cash equivalents (C$mm)", None, "Links to Snapshot."),
    ("Total liquidity (C$mm)", None, None),
]:
    put(ws, r, 1, label, BOLD)
    if label.startswith("Total facilities"):
        put(ws, r, 2, f"=SUM(B{r-2}:B{r-1})", BOLD, CUR); ws.cell(r, 2).border = TOPLINE
    elif label.startswith("Cash"):
        put(ws, r, 2, "=Snapshot!B14", GREEN, CUR)
    elif label.startswith("Total liquidity"):
        put(ws, r, 2, f"=B{r-2}+B{r-1}", BOLD, CUR); ws.cell(r, 2).border = TOPLINE; ws.cell(r, 2).fill = YEL
    else:
        put(ws, r, 2, v, BLUE, CUR)
    if note: put(ws, r, 5, note, SUB)
    r += 1


# =====================================================================
# 9. PRICEDECK
# =====================================================================
ws = wb.create_sheet("PriceDeck")
title(ws, "Commodity Price Deck — Bear / Base / Bull",
      "THE central scenario lever. Every yellow cell is an assumption you own and must defend. The Monte Carlo module samples around the Base column.")
widths(ws, {"A": 34, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 4, "H": 54})

r = 4
put(ws, r, 1, "Active scenario (1 = Bear, 2 = Base, 3 = Bull)", BOLD)
put(ws, r, 2, 2, BLUE, "0", align="right", fill=YEL)
put(ws, r, 8, "Downstream models read the 'Active' block at the bottom of this sheet.", SUB)
scen_cell = "$B$4"
r += 2

years = ["2026E", "2027E", "2028E", "2029E", "2030E"]
decks = {
    "Bear": {
        "WTI (US$/Bbl)": [76, 62, 58, 58, 58],
        "AECO (C$/GJ)": [1.90, 1.80, 1.90, 2.00, 2.10],
        "Dawn / diversified premium (C$/GJ)": [0.70, 0.70, 0.70, 0.70, 0.70],
        "CAD/USD exchange rate": [1.38, 1.40, 1.42, 1.42, 1.42],
        "Condensate differential to WTI (C$/Bbl)": [-2.00, -2.00, -2.00, -2.00, -2.00],
    },
    "Base": {
        "WTI (US$/Bbl)": [80, 75, 72, 70, 70],
        "AECO (C$/GJ)": [2.40, 2.90, 3.20, 3.30, 3.40],
        "Dawn / diversified premium (C$/GJ)": [0.90, 0.90, 0.90, 0.90, 0.90],
        "CAD/USD exchange rate": [1.37, 1.36, 1.35, 1.35, 1.35],
        "Condensate differential to WTI (C$/Bbl)": [0.00, 0.00, 0.00, 0.00, 0.00],
    },
    "Bull": {
        "WTI (US$/Bbl)": [84, 95, 92, 90, 88],
        "AECO (C$/GJ)": [3.00, 4.00, 4.50, 4.75, 5.00],
        "Dawn / diversified premium (C$/GJ)": [1.10, 1.20, 1.20, 1.20, 1.20],
        "CAD/USD exchange rate": [1.35, 1.33, 1.32, 1.32, 1.32],
        "Condensate differential to WTI (C$/Bbl)": [2.00, 2.00, 2.00, 2.00, 2.00],
    },
}
line_items = list(decks["Base"].keys())
fmts = {"WTI (US$/Bbl)": CUR1, "AECO (C$/GJ)": CUR1, "Dawn / diversified premium (C$/GJ)": CUR1,
        "CAD/USD exchange rate": '0.000', "Condensate differential to WTI (C$/Bbl)": CUR1}

block_rows = {}
for name in ["Bear", "Base", "Bull"]:
    r = sec(ws, r, f"{name.upper()} CASE", 6)
    put(ws, r, 1, "", BOLD_W)
    for i, y in enumerate(years): put(ws, r, 2 + i, y, BOLD_W, align="right")
    for c in range(1, 7): ws.cell(r, c).fill = HDR_FILL
    r += 1
    block_rows[name] = r
    for li in line_items:
        put(ws, r, 1, li, BLACK)
        for i, v in enumerate(decks[name][li]):
            put(ws, r, 2 + i, v, BLUE, fmts[li], fill=YEL)
        r += 1
    r += 1

put(ws, r, 1, "Sinclair start-up delay (years) — Bull / Base / Bear", BOLD)
put(ws, r, 2, 0, BLUE, NUM1, fill=YEL); put(ws, r, 3, 0.5, BLUE, NUM1, fill=YEL); put(ws, r, 4, 1.0, BLUE, NUM1, fill=YEL)
put(ws, r, 8, "Non-price lever. Slippage at Sinclair moves the case more than a dollar of WTI.", SUB)
r += 2

r = sec(ws, r, "ACTIVE DECK — read by all downstream models", 6)
put(ws, r, 1, "", BOLD_W)
for i, y in enumerate(years): put(ws, r, 2 + i, y, BOLD_W, align="right")
for c in range(1, 7): ws.cell(r, c).fill = HDR_FILL
r += 1
active_start = r
for j, li in enumerate(line_items):
    put(ws, r, 1, li, BOLD)
    for i in range(5):
        col = get_column_letter(2 + i)
        f = (f'=IF({scen_cell}=1,{col}{block_rows["Bear"]+j},'
             f'IF({scen_cell}=3,{col}{block_rows["Bull"]+j},{col}{block_rows["Base"]+j}))')
        put(ws, r, 2 + i, f, GREEN, fmts[li])
        ws.cell(r, 2 + i).fill = GREY
    r += 1
r += 1
put(ws, r, 1, "Realised gas price check (C$/Mcf)", BOLD)
for i in range(5):
    col = get_column_letter(2 + i)
    put(ws, r, 2 + i, f"=({col}{active_start+1}+{col}{active_start+2})*1.055", BLACK, CUR1)
put(ws, r, 8, "Approximate GJ to Mcf conversion of 1.055. Sanity check: FY2025 realised was C$3.02/Mcf.", SUB)
r += 1
put(ws, r, 1, "Condensate price (C$/Bbl)", BOLD)
for i in range(5):
    col = get_column_letter(2 + i)
    put(ws, r, 2 + i, f"={col}{active_start}*{col}{active_start+3}+{col}{active_start+4}", BLACK, CUR1)
put(ws, r, 8, "Sanity check: Q2 2026 realised condensate and oil was C$126.96/Bbl — a wide premium worth investigating.", SUB)
r += 2
put(ws, r, 1, "HEALTH WARNING", Font(name=FONT, size=10, bold=True, color="C00000"))
put(ws, r, 2, "These decks are illustrative starting points, not forecasts. Before the investment committee module, replace them with the forward strip on your valuation date and cite it. A case study is judged on whether its assumptions are defensible, not on whether the IRR is high.", BLACK)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[r].height = 42


# =====================================================================
# 10. DERIVED_METRICS
# =====================================================================
ws = wb.create_sheet("Derived_Metrics")
title(ws, "Trading Multiples — Where POU Trades Today",
      "Live off Snapshot, Reserves and Production_Ops. This is the 'entry price' evidence for the investment committee.")
widths(ws, {"A": 46, "B": 16, "C": 4, "D": 66})

r = 4
r = sec(ws, r, "INPUTS", 4)
inp = r
for label, f, fmt, note in [
    ("Share price (C$)", "=Snapshot!B6", CUR1, None),
    ("Shares outstanding (mm)", "=Snapshot!B10", NUM1, None),
    ("Market capitalisation (C$mm)", "=Snapshot!B11", CUR, None),
    ("Enterprise value, fully adjusted (C$mm)", f"=Snapshot!B{SNAP_EV}", CUR, "Market cap less net cash and securities."),
    ("Q2 2026 production (Boe/d)", f"=Production_Ops!G{tot_vol_row}", NUM, None),
    ("Q4 2026E production, midpoint (Boe/d)", f"=Guidance!D{GUID_Q4_ROW}", NUM, "Guided exit-quarter rate."),
    ("Q2 2026 AFF annualised (C$mm)", f"=Financials_Hist!G{AFF_ROW}*4", CUR, "Crude annualisation. Replace with the operating model output in Module 2."),
    ("H1 2026 AFF annualised (C$mm)", f"=(Financials_Hist!F{AFF_ROW}+Financials_Hist!G{AFF_ROW})*2", CUR, None),
    ("2P reserves (MBoe)", f"=Reserves!D{RES_VOL_ROW}", NUM, None),
    ("2P NPV10 before tax (C$mm)", f"=Reserves!D{PP_NPV10_ROW}", CUR, None),
]:
    put(ws, r, 1, label, BOLD); put(ws, r, 2, f, GREEN, fmt)
    if note: put(ws, r, 4, note, SUB)
    r += 1
r += 1

price_r, sh_r, mc_r, ev_r2, q2v_r, q4v_r, aff_q2_r, aff_h1_r, res_r, npv_r = range(inp, inp + 10)

r = sec(ws, r, "MULTIPLES", 4)
for label, f, fmt, note in [
    ("EV / AFF — annualised Q2 2026", f"=B{ev_r2}/B{aff_q2_r}", MULT,
     "Canadian intermediate producers have generally traded in a 3-5x band on cash flow. Verify against a live comp set."),
    ("EV / AFF — annualised H1 2026", f"=B{ev_r2}/B{aff_h1_r}", MULT, None),
    ("EV / flowing barrel — Q2 2026 (C$/Boe/d)", f"=B{ev_r2}*1000000/B{q2v_r}", CUR,
     "Recent Canadian liquids-rich deals have cleared roughly C$50,000-90,000/Boe/d. Confirm with current transactions."),
    ("EV / flowing barrel — Q4 2026E (C$/Boe/d)", f"=B{ev_r2}*1000000/B{q4v_r}", CUR,
     "The forward number. The gap between these two rows is what you are being asked to pay for growth."),
    ("EV / 2P reserves (C$/Boe)", f"=B{ev_r2}*1000/B{res_r}", CUR1, None),
    ("Market cap / 2P NPV10 (P/NAV)", f"=B{mc_r}/B{npv_r}", MULT,
     "Above 1.0x means the market pays more than the engineered before-tax reserve value."),
    ("2P NPV10 per share (C$)", f"=B{npv_r}/B{sh_r}", CUR1, None),
    ("Implied premium in the share price to 2P NPV10 per share", f"=B{price_r}/(B{npv_r}/B{sh_r})-1", PCT, None),
]:
    put(ws, r, 1, label, BOLD); put(ws, r, 2, f, BLACK, fmt)
    ws.cell(r, 2).fill = YEL
    if note:
        put(ws, r, 4, note, SUB)
        ws.cell(r, 4).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 26
    r += 1
r += 1

r = sec(ws, r, "TAKE-PRIVATE PRICE LADDER — what a bid would have to clear", 4)
put(ws, r, 1, "Premium to market", BOLD_W); put(ws, r, 2, "Offer (C$/sh)", BOLD_W, align="right")
put(ws, r, 4, "Implied equity cheque and P/NAV", BOLD_W)
for c in range(1, 5): ws.cell(r, c).fill = HDR_FILL
r += 1
ladder_start = r
for prem in [0.00, 0.15, 0.25, 0.35]:
    put(ws, r, 1, prem, BLUE, PCT)
    put(ws, r, 2, f"=B{price_r}*(1+A{r})", BLACK, CUR1)
    put(ws, r, 4, f"=\"Equity value C$\"&TEXT(B{r}*B{sh_r},\"#,##0\")&\"mm  |  P/NAV \"&TEXT(B{r}*B{sh_r}/B{npv_r},\"0.00\")&\"x\"", BLACK)
    r += 1
r += 1
put(ws, r, 1, "READ THIS BEFORE MODULE 2", Font(name=FONT, size=10, bold=True, color="C00000"))
put(ws, r, 2, "If a 25-35% control premium puts the purchase price at a large multiple of 2P NPV10, the whole return depends on value the reserve report does not yet book: Sinclair, the Alhambra phase 3 expansion, and the undeveloped land. Your NAV module must value those explicitly and defensibly, because that is where the entire thesis lives.", BLACK)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[r].height = 56

# ---------- Global font pass ----------
for sheet in wb.worksheets:
    sheet.sheet_view.showGridLines = False
    for row in sheet.iter_rows():
        for cell in row:
            if cell.font is None or cell.font.name is None:
                cell.font = BLACK

wb.save("/home/claude/pou/POU_DataBook_v1.xlsx")
print("saved")
