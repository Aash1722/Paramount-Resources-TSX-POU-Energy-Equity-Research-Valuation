"""
Paramount Resources (TSX:POU) - Module 2: Operating Model
Five-year asset-level forecast: production -> revenue -> netback -> AFF -> FCF.
Self-contained and scenario-switchable.
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
TITLE = Font(name=F, size=14, bold=True, color="1F3864")
SUB   = Font(name=F, size=10, italic=True, color="595959")
HDR   = PatternFill("solid", fgColor="1F3864")
SECF  = PatternFill("solid", fgColor="D9E2F3")
YEL   = PatternFill("solid", fgColor="FFFF00")
GRY   = PatternFill("solid", fgColor="F2F2F2")
TOP   = Border(top=Side(style="thin", color="404040"))

CUR, CUR1, NUM, NUM1, PCT, PCT0, MULT = (
    '$#,##0;($#,##0);-', '$#,##0.00;($#,##0.00);-', '#,##0;(#,##0);-',
    '#,##0.0;(#,##0.0);-', '0.0%', '0%', '0.00"x"')

YEARS = ["2026E", "2027E", "2028E", "2029E", "2030E"]
NY = len(YEARS)
COLS = [get_column_letter(2 + i) for i in range(NY)]   # B..F
NOTE_COL = 8                                            # H

R = {}   # row registry: R[(sheet,key)] = row

def title(ws, name, sub):
    ws["A1"] = name; ws["A1"].font = TITLE
    ws["A2"] = sub;  ws["A2"].font = SUB
    ws.freeze_panes = "B5"

def sec(ws, r, text, w=8):
    ws.cell(r, 1, text).font = Font(name=F, size=10, bold=True, color="1F3864")
    for c in range(1, w + 1): ws.cell(r, c).fill = SECF
    return r + 1

def hdrow(ws, r, label=""):
    ws.cell(r, 1, label).font = BOLDW
    for i, y in enumerate(YEARS):
        c = ws.cell(r, 2 + i, y); c.font = BOLDW; c.alignment = Alignment(horizontal="right")
    ws.cell(r, NOTE_COL, "Note").font = BOLDW
    for c in range(1, NOTE_COL + 1): ws.cell(r, c).fill = HDR
    return r + 1

def line(ws, r, sheet, key, label, values=None, formula=None, fmt=NUM,
         font=None, note=None, total=False, fill=None, indent=False):
    """values: list of hardcodes. formula: callable(col_letter, col_index)->str."""
    lab = ("   " if indent else "") + label
    c = ws.cell(r, 1, lab); c.font = font or (BOLD if total else BLACK)
    for i, col in enumerate(COLS):
        if formula is not None:
            cell = ws.cell(r, 2 + i, formula(col, i)); cell.font = BLACK if not total else BOLD
        else:
            v = values[i]
            cell = ws.cell(r, 2 + i, v); cell.font = BLUE
        cell.number_format = fmt
        if total: cell.border = TOP
        if fill: cell.fill = fill
    if note:
        n = ws.cell(r, NOTE_COL, note); n.font = SUB
        n.alignment = Alignment(wrap_text=True, vertical="top")
        if len(note) > 70: ws.row_dimensions[r].height = 28
    R[(sheet, key)] = r
    return r + 1

def W(ws, spec):
    for k, v in spec.items(): ws.column_dimensions[k].width = v


# ======================================================================
# README
# ======================================================================
ws = wb.active; ws.title = "README"
title(ws, "Paramount Resources — Module 2: Five-Year Operating Model",
      "Asset-level production forecast driving revenue, netback, adjusted funds flow and free cash flow, 2026E–2030E.")
W(ws, {"A": 32, "B": 100})
r = 4
r = sec(ws, r, "HOW THIS MODEL WORKS", 2)
for a, b in [
    ("Structure", "Assumptions drive Production, which drives Revenue_Netback, which drives Cash_Flow. Summary reports outputs and runs sanity checks against company guidance."),
    ("Scenario switch", "Assumptions!B4. Enter 1 for Bear, 2 for Base, 3 for Bull. Every price line re-reads automatically."),
    ("Self-contained", "The price deck is duplicated here rather than linked to the Data Book, so the file runs standalone. If you change the deck in one file, change it in both."),
    ("Units", "Volumes in Boe/d unless stated. Gas converted at 6 Mcf : 1 Boe. All dollars are Canadian millions unless stated."),
    ("Colour", "Blue = hardcoded assumption you own. Black = formula. Green = cross-sheet link. Yellow = the levers that move the answer."),
]:
    ws.cell(r, 1, a).font = BOLD
    c = ws.cell(r, 2, b); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30; r += 1

r += 1
r = sec(ws, r, "THE THREE JUDGEMENTS THAT DRIVE THE ANSWER", 2)
for a, b in [
    ("1. Willesden Green ramp", "Grows from ~27.5 kBoe/d in Q2 2026 toward the company's stated ~50 kBoe/d plateau. Liquids-rich (~55%), so it carries the revenue."),
    ("2. Sinclair timing", "First gas Q4 2027, ramping through 2028. Firm egress of 335 MMcf/d caps it near 56 kBoe/d. Dry gas, so it adds volume far faster than it adds netback."),
    ("3. Kaybob decline", "Held at 19–20 kBoe/d to 2028 per company guidance, then declining as capital stays with the growth assets."),
]:
    ws.cell(r, 1, a).font = BOLD
    c = ws.cell(r, 2, b); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30; r += 1

r += 1
r = sec(ws, r, "RECONCILIATION TO GUIDANCE — read this before defending the model", 2)
for a, b in [
    ("2026E and 2027E", "Tuned to land inside company guidance. Summary tab checks this automatically."),
    ("The 2027 exit rate", "The company guides to over 100,000 Boe/d exiting 2027 while guiding 2027 average volumes of only 60,000–65,000. Those reconcile only if the exit rate is a December spot rate rather than a Q4 average. This model assumes exactly that, and it is the single most important interpretive call in the file. State it out loud in any interview."),
    ("The oil price deck", "WTI traded near US$86 in late August 2026, driven by US-Iran tension and Strait of Hormuz disruption, with futures in backwardation. The Base case therefore starts near spot and declines as the geopolitical risk premium unwinds toward mid-cycle. Bear assumes rapid de-escalation; Bull assumes sustained disruption. Do not model spot flat for five years — the forward curve does not."),
    ("2028E–2030E", "No company guidance exists. These are my assumptions and should be defended on Sinclair egress capacity and the Willesden Green plateau, not on management statements."),
]:
    ws.cell(r, 1, a).font = BOLD
    c = ws.cell(r, 2, b); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 42 if len(b) > 150 else 30; r += 1

r += 1
ws.cell(r, 1, "Hypothetical case study. Not investment advice.").font = SUB


# ======================================================================
# ASSUMPTIONS
# ======================================================================
ws = wb.create_sheet("Assumptions"); S = "Assumptions"
title(ws, "Assumptions & Scenario Switch",
      "Yellow cells are the levers. Everything downstream is formula-driven off this sheet.")
W(ws, {"A": 44, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 3, "H": 52})

r = 4
ws.cell(r, 1, "SCENARIO  (1 = Bear, 2 = Base, 3 = Bull)").font = BOLD
c = ws.cell(r, 2, 2); c.font = BLUE; c.fill = YEL; c.number_format = "0"
c.alignment = Alignment(horizontal="center")
ws.cell(r, NOTE_COL, "Change this one cell to re-run the whole model.").font = SUB
SW = "$B$4"
r += 2

# ---- price decks ----
DECKS = {
    "Bear": {"wti": [76, 62, 58, 58, 58], "aeco": [1.90, 1.80, 1.90, 2.00, 2.10],
             "prem": [0.70]*5, "fx": [1.38, 1.40, 1.42, 1.42, 1.42], "cdiff": [-2.0]*5},
    "Base": {"wti": [80, 75, 72, 70, 70], "aeco": [2.40, 2.90, 3.20, 3.30, 3.40],
             "prem": [0.90]*5, "fx": [1.37, 1.36, 1.35, 1.35, 1.35], "cdiff": [0.0]*5},
    "Bull": {"wti": [84, 95, 92, 90, 88], "aeco": [3.00, 4.00, 4.50, 4.75, 5.00],
             "prem": [1.10, 1.20, 1.20, 1.20, 1.20], "fx": [1.35, 1.33, 1.32, 1.32, 1.32], "cdiff": [2.0]*5},
}
PLINES = [("wti", "WTI (US$/Bbl)", CUR1), ("aeco", "AECO (C$/GJ)", CUR1),
          ("prem", "Diversification premium (C$/GJ)", CUR1),
          ("fx", "CAD/USD", '0.000'), ("cdiff", "Condensate diff. to WTI (C$/Bbl)", CUR1)]

deck_rows = {}
for name in ["Bear", "Base", "Bull"]:
    r = sec(ws, r, f"{name.upper()} PRICE DECK", 6)
    r = hdrow(ws, r)
    deck_rows[name] = {}
    for key, lab, fmt in PLINES:
        deck_rows[name][key] = r
        r = line(ws, r, S, f"{name}_{key}", lab, values=DECKS[name][key], fmt=fmt, fill=YEL)
    r += 1

r = sec(ws, r, "ACTIVE PRICE DECK — read by Revenue_Netback", 6)
r = hdrow(ws, r)
for key, lab, fmt in PLINES:
    br, ba, bu = deck_rows["Bear"][key], deck_rows["Base"][key], deck_rows["Bull"][key]
    r = line(ws, r, S, f"act_{key}", lab,
             formula=lambda col, i, br=br, ba=ba, bu=bu:
                 f"=IF({SW}=1,{col}{br},IF({SW}=3,{col}{bu},{col}{ba}))",
             fmt=fmt, font=GREEN, fill=GRY, total=False)
r += 1

# ---- realised price bridge ----
r = sec(ws, r, "REALISED PRICE BUILD", 6)
r = hdrow(ws, r)
ar = R[(S, "act_aeco")]; pr = R[(S, "act_prem")]
r = line(ws, r, S, "gas_price", "Realised natural gas (C$/Mcf)",
         formula=lambda col, i: f"=({col}{ar}+{col}{pr})*1.055", fmt=CUR1,
         note="GJ to Mcf at 1.055. Roughly 48% of gas is priced at Dawn, Malin and Emerson rather than AECO, which is what the premium represents.")
wr = R[(S, "act_wti")]; fr = R[(S, "act_fx")]; cd = R[(S, "act_cdiff")]
r = line(ws, r, S, "cond_price", "Realised condensate and oil (C$/Bbl)",
         formula=lambda col, i: f"={col}{wr}*{col}{fr}+{col}{cd}", fmt=CUR1,
         note="Validated on Q1 2026: actual C$96.27 vs WTI ~US$70 x 1.37 = C$95.9. Q2 2026 printed C$126.96 because WTI spiked on Hormuz disruption.")
r = line(ws, r, S, "ngl_price", "Realised other NGLs (C$/Bbl)",
         formula=lambda col, i: f"={col}{R[(S,'cond_price')]}*{col}{r+0}", fmt=CUR1)
# replace with a clean ratio row instead
ws.cell(R[(S, "ngl_price")], 1, "Realised other NGLs (C$/Bbl)")
ngl_ratio_row = r
r = line(ws, r, S, "ngl_ratio", "   Other NGLs as % of condensate price", values=[0.32]*5, fmt=PCT, fill=YEL,
         note="Historical range roughly 30–36%. Ethane, propane and butane.")
for i, col in enumerate(COLS):
    cell = ws.cell(R[(S, "ngl_price")], 2 + i)
    cell.value = f"={col}{R[(S,'cond_price')]}*{col}{R[(S,'ngl_ratio')]}"
    cell.font = BLACK; cell.number_format = CUR1
r += 1

# ---- production assumptions ----
r = sec(ws, r, "PRODUCTION — ANNUAL AVERAGE BY ASSET (Boe/d)", 6)
r = hdrow(ws, r)
r = line(ws, r, S, "wg_vol", "Willesden Green", values=[32300, 41000, 47000, 50000, 50000], fmt=NUM, fill=YEL,
         note="Q4 2026 implied ~42 kBoe/d as Alhambra phase 2 fills. Plateaus at the company's stated ~50 kBoe/d.")
r = line(ws, r, S, "kb_vol", "Kaybob", values=[19300, 19000, 19000, 17000, 15500], fmt=NUM, fill=YEL,
         note="Company guides 19–20 kBoe/d through 2028. Declines thereafter as capital stays with growth assets.")
r = line(ws, r, S, "sc_vol", "Sinclair", values=[0, 3000, 45000, 55000, 58000], fmt=NUM, fill=YEL,
         note="First gas Q4 2027, so 2027 is a partial-quarter average only. Capped near 56 kBoe/d by 335 MMcf/d firm egress.")
r = line(ws, r, S, "ot_vol", "Other properties", values=[400, 400, 350, 300, 300], fmt=NUM, fill=YEL)
r = line(ws, r, S, "tot_vol", "Total sales volumes (Boe/d)",
         formula=lambda col, i: f"=SUM({col}{R[(S,'wg_vol')]}:{col}{R[(S,'ot_vol')]})",
         fmt=NUM, total=True, fill=YEL)
r += 1

r = sec(ws, r, "PRODUCT MIX BY ASSET (% of Boe that is liquids, and condensate share of those liquids)", 6)
r = hdrow(ws, r)
for k, lab, vals, note in [
    ("wg_liq", "Willesden Green — liquids %", [0.55]*5, "Q2 2026 actual was ~55%. Liquids-rich Duvernay."),
    ("wg_cnd", "Willesden Green — condensate share of liquids", [0.80]*5, None),
    ("kb_liq", "Kaybob — liquids %", [0.40]*5, None),
    ("kb_cnd", "Kaybob — condensate share of liquids", [0.70]*5, None),
    ("sc_liq", "Sinclair — liquids %", [0.10]*5, "Dry gas Montney. Adds volume much faster than it adds netback."),
    ("sc_cnd", "Sinclair — condensate share of liquids", [0.75]*5, None),
    ("ot_liq", "Other — liquids %", [0.30]*5, None),
    ("ot_cnd", "Other — condensate share of liquids", [0.70]*5, None),
]:
    r = line(ws, r, S, k, lab, values=vals, fmt=PCT0, fill=YEL, note=note)
r += 1

r = sec(ws, r, "UNIT COSTS AND CORPORATE ITEMS", 6)
r = hdrow(ws, r)
r = line(ws, r, S, "roy_pct", "Royalty rate (% of petroleum and natural gas sales)", values=[0.055, 0.060, 0.070, 0.080, 0.085],
         fmt=PCT, fill=YEL, note="Q2 2026 was 5.4%. Rises as new-well royalty rate holidays are consumed. A commonly missed driver.")
r = line(ws, r, S, "opex_boe", "Operating expense (C$/Boe)", values=[8.50, 7.75, 6.75, 6.25, 6.10], fmt=CUR1, fill=YEL,
         note="Q2 2026 was C$8.66 and falling. Sinclair volumes dilute the corporate average further.")
r = line(ws, r, S, "tran_boe", "Transportation and NGLs processing (C$/Boe)", values=[4.75, 4.60, 4.25, 4.10, 4.05], fmt=CUR1, fill=YEL)
r = line(ws, r, S, "ga_boe", "General and administrative (C$/Boe)", values=[1.60, 1.40, 1.00, 0.90, 0.90], fmt=CUR1, fill=YEL)
r = line(ws, r, S, "hedge_boe", "Risk management settlements (C$/Boe)", values=[-1.50, -0.75, 0.00, 0.00, 0.00], fmt=CUR1, fill=YEL,
         note="Q2 2026 was negative C$2.48. Hedges roll off after 2027, leaving the model unhedged.")
r = line(ws, r, S, "tax_rate", "Cash tax rate (% of pre-tax cash flow)", values=[0.00, 0.00, 0.05, 0.12, 0.18], fmt=PCT, fill=YEL,
         note="Nil near-term: large tax pools remain from the Grande Prairie disposition. Verify pool balances in the AIF before defending.")
r = line(ws, r, S, "int_inc", "Net interest income / (expense) (C$mm)", values=[12, 4, -8, -6, 0], fmt=CUR, fill=YEL,
         note="Interest earned on cash early, modest drawings later. Revisit if the funding gap is financed with debt.")
r += 1

r = sec(ws, r, "CAPITAL AND RETURNS", 6)
r = hdrow(ws, r)
r = line(ws, r, S, "capex", "Capital expenditures (C$mm)", values=[1050, 1000, 800, 700, 700], fmt=CUR, fill=YEL,
         note="2026 and 2027 are company guidance. 2028 onward assumes Sinclair construction is complete and spending normalises.")
r = line(ws, r, S, "aro", "Abandonment and reclamation (C$mm)", values=[35, 35, 40, 40, 45], fmt=CUR, fill=YEL)
r = line(ws, r, S, "dps", "Dividend per share (C$)", values=[0.60]*5, fmt=CUR1, fill=YEL,
         note="ASSUMPTION — confirm the current declared rate before use.")
r = line(ws, r, S, "shares", "Shares outstanding (mm)", values=[145.8]*5, fmt=NUM1, fill=YEL)
r = line(ws, r, S, "open_cash", "Opening net cash at 1-Jan-2026 (C$mm)", values=[672.8, 0, 0, 0, 0], fmt=CUR,
         note="Actual at 31-Dec-2025. Only the 2026 cell is used; later years roll forward.")


# ======================================================================
# PRODUCTION
# ======================================================================
ws = wb.create_sheet("Production"); P = "Production"
title(ws, "Production Forecast",
      "Asset volumes converted into gas, condensate and NGL streams. All links green to Assumptions.")
W(ws, {"A": 44, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 3, "H": 52})
r = 4
r = sec(ws, r, "VOLUMES BY ASSET (Boe/d)", 6)
r = hdrow(ws, r)
for k, lab in [("wg_vol", "Willesden Green"), ("kb_vol", "Kaybob"), ("sc_vol", "Sinclair"), ("ot_vol", "Other properties")]:
    r = line(ws, r, P, k, lab, formula=lambda col, i, k=k: f"=Assumptions!{col}{R[(S,k)]}", fmt=NUM, font=GREEN)
r = line(ws, r, P, "tot", "Total (Boe/d)",
         formula=lambda col, i: f"=SUM({col}{R[(P,'wg_vol')]}:{col}{R[(P,'ot_vol')]})", fmt=NUM, total=True)
r = line(ws, r, P, "yoy", "Year-over-year growth",
         formula=lambda col, i: (f"={col}{R[(P,'tot')]}/52133-1" if i == 0
                                 else f"={col}{R[(P,'tot')]}/{COLS[i-1]}{R[(P,'tot')]}-1"),
         fmt=PCT, note="2026 growth is measured against the 2026 guidance midpoint reconstruction of 52,133 Boe/d.")
r += 1

r = sec(ws, r, "LIQUIDS AND GAS SPLIT (Boe/d)", 6)
r = hdrow(ws, r)
pairs = [("wg", "wg_vol", "wg_liq", "wg_cnd", "Willesden Green"),
         ("kb", "kb_vol", "kb_liq", "kb_cnd", "Kaybob"),
         ("sc", "sc_vol", "sc_liq", "sc_cnd", "Sinclair"),
         ("ot", "ot_vol", "ot_liq", "ot_cnd", "Other")]
for tag, vk, lk, ck, lab in pairs:
    r = line(ws, r, P, f"{tag}_liq", f"{lab} — liquids",
             formula=lambda col, i, vk=vk, lk=lk: f"={col}{R[(P,vk)]}*Assumptions!{col}{R[(S,lk)]}", fmt=NUM, indent=True)
r = line(ws, r, P, "tot_liq", "Total liquids (Bbl/d)",
         formula=lambda col, i: f"=SUM({col}{R[(P,'wg_liq')]}:{col}{R[(P,'ot_liq')]})", fmt=NUM, total=True)
r = line(ws, r, P, "tot_gas_boe", "Total natural gas (Boe/d)",
         formula=lambda col, i: f"={col}{R[(P,'tot')]}-{col}{R[(P,'tot_liq')]}", fmt=NUM, total=True)
r = line(ws, r, P, "liq_pct", "Liquids as % of total",
         formula=lambda col, i: f"={col}{R[(P,'tot_liq')]}/{col}{R[(P,'tot')]}", fmt=PCT, fill=YEL,
         note="Watch this fall as Sinclair dry gas arrives. It is the reason revenue per Boe declines even as volumes double.")
r += 1

r = sec(ws, r, "SALES STREAMS", 6)
r = hdrow(ws, r)
for tag, _, _, ck, lab in pairs:
    r = line(ws, r, P, f"{tag}_cnd", f"{lab} — condensate",
             formula=lambda col, i, tag=tag, ck=ck: f"={col}{R[(P,tag+'_liq')]}*Assumptions!{col}{R[(S,ck)]}", fmt=NUM, indent=True)
r = line(ws, r, P, "cond", "Condensate and oil (Bbl/d)",
         formula=lambda col, i: f"=SUM({col}{R[(P,'wg_cnd')]}:{col}{R[(P,'ot_cnd')]})", fmt=NUM, total=True)
r = line(ws, r, P, "ngl", "Other NGLs (Bbl/d)",
         formula=lambda col, i: f"={col}{R[(P,'tot_liq')]}-{col}{R[(P,'cond')]}", fmt=NUM, total=True)
r = line(ws, r, P, "gas", "Natural gas (MMcf/d)",
         formula=lambda col, i: f"={col}{R[(P,'tot_gas_boe')]}*6/1000", fmt=NUM1, total=True,
         note="Sanity check: FY2025 actual was 131.9 MMcf/d; Q2 2026 was 144.4 MMcf/d.")
r += 1
r = line(ws, r, P, "check", "Check — streams reconvert to total Boe/d",
         formula=lambda col, i: f"={col}{R[(P,'gas')]}*1000/6+{col}{R[(P,'cond')]}+{col}{R[(P,'ngl')]}-{col}{R[(P,'tot')]}",
         fmt=NUM1, note="Must be zero in every column.", font=RED)


# ======================================================================
# REVENUE_NETBACK
# ======================================================================
ws = wb.create_sheet("Revenue_Netback"); V = "Revenue_Netback"
title(ws, "Revenue and Netback",
      "Volumes times realised prices, less royalties, operating and transportation costs.")
W(ws, {"A": 44, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 3, "H": 52})
r = 4
r = sec(ws, r, "REVENUE (C$mm)", 6)
r = hdrow(ws, r)
r = line(ws, r, V, "rev_gas", "Natural gas",
         formula=lambda col, i: f"=Production!{col}{R[(P,'gas')]}*1000*Assumptions!{col}{R[(S,'gas_price')]}*365/1000000", fmt=CUR)
r = line(ws, r, V, "rev_cnd", "Condensate and oil",
         formula=lambda col, i: f"=Production!{col}{R[(P,'cond')]}*Assumptions!{col}{R[(S,'cond_price')]}*365/1000000", fmt=CUR)
r = line(ws, r, V, "rev_ngl", "Other NGLs",
         formula=lambda col, i: f"=Production!{col}{R[(P,'ngl')]}*Assumptions!{col}{R[(S,'ngl_price')]}*365/1000000", fmt=CUR)
r = line(ws, r, V, "rev", "Petroleum and natural gas sales",
         formula=lambda col, i: f"=SUM({col}{R[(V,'rev_gas')]}:{col}{R[(V,'rev_ngl')]})", fmt=CUR, total=True)
r = line(ws, r, V, "rev_boe", "   Revenue per Boe (C$)",
         formula=lambda col, i: f"={col}{R[(V,'rev')]}*1000000/Production!{col}{R[(P,'tot')]}/365", fmt=CUR1,
         note="Q2 2026 actual was C$60.59/Boe. Expect this to fall as Sinclair gas dilutes the mix.")
r += 1

r = sec(ws, r, "NETBACK (C$mm)", 6)
r = hdrow(ws, r)
r = line(ws, r, V, "roy", "Royalties",
         formula=lambda col, i: f"=-{col}{R[(V,'rev')]}*Assumptions!{col}{R[(S,'roy_pct')]}", fmt=CUR)
for k, ak, lab in [("opex", "opex_boe", "Operating expense"), ("tran", "tran_boe", "Transportation and NGLs processing")]:
    r = line(ws, r, V, k, lab,
             formula=lambda col, i, ak=ak: f"=-Production!{col}{R[(P,'tot')]}*Assumptions!{col}{R[(S,ak)]}*365/1000000", fmt=CUR)
r = line(ws, r, V, "nb", "Operating netback",
         formula=lambda col, i: f"={col}{R[(V,'rev')]}+SUM({col}{R[(V,'roy')]}:{col}{R[(V,'tran')]})", fmt=CUR, total=True)
r = line(ws, r, V, "hedge", "Risk management settlements",
         formula=lambda col, i: f"=Production!{col}{R[(P,'tot')]}*Assumptions!{col}{R[(S,'hedge_boe')]}*365/1000000", fmt=CUR)
r = line(ws, r, V, "nb_all", "Netback including risk management",
         formula=lambda col, i: f"={col}{R[(V,'nb')]}+{col}{R[(V,'hedge')]}", fmt=CUR, total=True)
r = line(ws, r, V, "nb_boe", "   Netback per Boe (C$)",
         formula=lambda col, i: f"={col}{R[(V,'nb_all')]}*1000000/Production!{col}{R[(P,'tot')]}/365", fmt=CUR1, fill=YEL,
         note="FY2025 actual was C$31.96 including hedges; Q2 2026 was C$41.69. This is the headline unit economic.")


# ======================================================================
# CASH_FLOW
# ======================================================================
ws = wb.create_sheet("Cash_Flow"); C = "Cash_Flow"
title(ws, "Adjusted Funds Flow and Free Cash Flow",
      "Netback less corporate costs, then capital. The net cash rollforward is the solvency test.")
W(ws, {"A": 44, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 3, "H": 52})
r = 4
r = sec(ws, r, "ADJUSTED FUNDS FLOW (C$mm)", 6)
r = hdrow(ws, r)
r = line(ws, r, C, "nb", "Netback including risk management",
         formula=lambda col, i: f"=Revenue_Netback!{col}{R[(V,'nb_all')]}", fmt=CUR, font=GREEN)
r = line(ws, r, C, "ga", "General and administrative",
         formula=lambda col, i: f"=-Production!{col}{R[(P,'tot')]}*Assumptions!{col}{R[(S,'ga_boe')]}*365/1000000", fmt=CUR)
r = line(ws, r, C, "int", "Net interest income / (expense)",
         formula=lambda col, i: f"=Assumptions!{col}{R[(S,'int_inc')]}", fmt=CUR, font=GREEN)
r = line(ws, r, C, "pretax", "Pre-tax cash flow",
         formula=lambda col, i: f"=SUM({col}{R[(C,'nb')]}:{col}{R[(C,'int')]})", fmt=CUR, total=True)
r = line(ws, r, C, "tax", "Cash taxes",
         formula=lambda col, i: f"=-MAX(0,{col}{R[(C,'pretax')]})*Assumptions!{col}{R[(S,'tax_rate')]}", fmt=CUR)
r = line(ws, r, C, "aff", "Adjusted funds flow",
         formula=lambda col, i: f"={col}{R[(C,'pretax')]}+{col}{R[(C,'tax')]}", fmt=CUR, total=True, fill=YEL)
r = line(ws, r, C, "aff_sh", "   Adjusted funds flow per share (C$)",
         formula=lambda col, i: f"={col}{R[(C,'aff')]}/Assumptions!{col}{R[(S,'shares')]}", fmt=CUR1)
r += 1

r = sec(ws, r, "FREE CASH FLOW (C$mm)", 6)
r = hdrow(ws, r)
r = line(ws, r, C, "capex", "Capital expenditures",
         formula=lambda col, i: f"=-Assumptions!{col}{R[(S,'capex')]}", fmt=CUR, font=GREEN)
r = line(ws, r, C, "aro", "Abandonment and reclamation",
         formula=lambda col, i: f"=-Assumptions!{col}{R[(S,'aro')]}", fmt=CUR, font=GREEN)
r = line(ws, r, C, "fcf", "Free cash flow before dividends",
         formula=lambda col, i: f"={col}{R[(C,'aff')]}+{col}{R[(C,'capex')]}+{col}{R[(C,'aro')]}", fmt=CUR, total=True, fill=YEL)
r = line(ws, r, C, "div", "Dividends",
         formula=lambda col, i: f"=-Assumptions!{col}{R[(S,'dps')]}*Assumptions!{col}{R[(S,'shares')]}", fmt=CUR)
r = line(ws, r, C, "fcf_ad", "Free cash flow after dividends",
         formula=lambda col, i: f"={col}{R[(C,'fcf')]}+{col}{R[(C,'div')]}", fmt=CUR, total=True)
r = line(ws, r, C, "capint", "   Capital intensity (capex / AFF)",
         formula=lambda col, i: f"=IFERROR(-{col}{R[(C,'capex')]}/{col}{R[(C,'aff')]},\"n/m\")", fmt=MULT,
         note="Above 1.0x is outspend. The year this crosses below 1.0x is when the equity story changes.")
r += 1

r = sec(ws, r, "NET CASH / (NET DEBT) ROLLFORWARD (C$mm)", 6)
r = hdrow(ws, r)
r = line(ws, r, C, "open", "Opening net cash / (net debt)",
         formula=lambda col, i: (f"=Assumptions!B{R[(S,'open_cash')]}" if i == 0
                                  else f"={COLS[i-1]}{r+2}"), fmt=CUR)
r = line(ws, r, C, "chg", "Free cash flow after dividends",
         formula=lambda col, i: f"={col}{R[(C,'fcf_ad')]}", fmt=CUR, font=GREEN)
r = line(ws, r, C, "close", "Closing net cash / (net debt)",
         formula=lambda col, i: f"={col}{R[(C,'open')]}+{col}{R[(C,'chg')]}", fmt=CUR, total=True, fill=YEL,
         note="Negative means net debt. Compare against C$750mm of undrawn facilities — that is the true liquidity runway.")
r += 1
r = line(ws, r, C, "lev", "Net debt / AFF",
         formula=lambda col, i: f"=IFERROR(MAX(0,-{col}{R[(C,'close')]})/{col}{R[(C,'aff')]},\"n/m\")", fmt=MULT,
         note="Canadian producers are generally uncomfortable above 1.5x. Watch the trough year.")


# ======================================================================
# SUMMARY
# ======================================================================
ws = wb.create_sheet("Summary"); U = "Summary"
title(ws, "Output Summary & Guidance Checks",
      "The numbers you quote in a pitch, plus automatic tests that the model respects company guidance.")
W(ws, {"A": 46, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 3, "H": 50})
r = 4
r = sec(ws, r, "KEY OUTPUTS", 6)
r = hdrow(ws, r)
for k, src, lab, fmt in [
    ("vol", f"Production!{{c}}{R[(P,'tot')]}", "Production (Boe/d)", NUM),
    ("liq", f"Production!{{c}}{R[(P,'liq_pct')]}", "Liquids %", PCT),
    ("rev", f"Revenue_Netback!{{c}}{R[(V,'rev')]}", "Revenue (C$mm)", CUR),
    ("nbb", f"Revenue_Netback!{{c}}{R[(V,'nb_boe')]}", "Netback (C$/Boe)", CUR1),
    ("aff", f"Cash_Flow!{{c}}{R[(C,'aff')]}", "Adjusted funds flow (C$mm)", CUR),
    ("cap", f"Assumptions!{{c}}{R[(S,'capex')]}", "Capital expenditures (C$mm)", CUR),
    ("fcf", f"Cash_Flow!{{c}}{R[(C,'fcf')]}", "Free cash flow before dividends (C$mm)", CUR),
    ("cash", f"Cash_Flow!{{c}}{R[(C,'close')]}", "Closing net cash / (net debt) (C$mm)", CUR),
]:
    r = line(ws, r, U, k, lab, formula=lambda col, i, s=src: "=" + s.format(c=col), fmt=fmt, font=GREEN,
             fill=YEL if k in ("aff", "fcf", "cash") else None)
r += 1

r = sec(ws, r, "SANITY CHECKS AGAINST COMPANY GUIDANCE", 6)
hr = r
ws.cell(r, 1, "Test").font = BOLDW
for j, h in enumerate(["Model", "Guidance low", "Guidance high", "Result"]):
    c = ws.cell(r, 2 + j, h); c.font = BOLDW; c.alignment = Alignment(horizontal="right")
for c in range(1, 7): ws.cell(r, c).fill = HDR
r += 1
checks = [
    ("2026E production (Boe/d)", f"=B{R[(U,'vol')]}", 51000, 53000, NUM),
    ("2027E production (Boe/d)", f"=C{R[(U,'vol')]}", 60000, 65000, NUM),
    ("2026E capital expenditures (C$mm)", f"=B{R[(U,'cap')]}", 1000, 1100, CUR),
    ("2027E capital expenditures (C$mm)", f"=C{R[(U,'cap')]}", 950, 1050, CUR),
]
for lab, f_, lo, hi, fmt in checks:
    ws.cell(r, 1, lab).font = BOLD
    a = ws.cell(r, 2, f_); a.font = BLACK; a.number_format = fmt
    b = ws.cell(r, 3, lo); b.font = BLUE; b.number_format = fmt
    c2 = ws.cell(r, 4, hi); c2.font = BLUE; c2.number_format = fmt
    d = ws.cell(r, 5, f'=IF(AND(B{r}>=C{r},B{r}<=D{r}),"IN RANGE","OUTSIDE — REVIEW")')
    d.font = BOLD; d.alignment = Alignment(horizontal="right")
    r += 1
r += 1
ws.cell(r, 1, "Stream reconciliation (must be zero)").font = BOLD
e = ws.cell(r, 2, f"=SUM(Production!B{R[(P,'check')]}:F{R[(P,'check')]})")
e.font = BLACK; e.number_format = NUM1
ws.cell(r, 5, '=IF(ABS(B' + str(r) + ')<0.01,"OK","ERROR")').font = BOLD
ws.cell(r, 5).alignment = Alignment(horizontal="right")
r += 2

r = sec(ws, r, "WHAT THE MODEL SAYS — fill in after running all three scenarios", 6)
for t in [
    "1. Production roughly doubles by 2029 on the Sinclair build-out, but liquids weighting falls, so netback per Boe compresses.",
    "2. The outspend persists through 2027. Check the trough year on the net cash rollforward against C$750mm of undrawn facilities.",
    "3. Run the Bear case. If the trough net debt exceeds available liquidity, the equity story depends on capital markets access, not operations.",
]:
    c = ws.cell(r, 1, t); c.font = BLACK
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 26
    r += 1

for sheet in wb.worksheets:
    sheet.sheet_view.showGridLines = False

wb.save("/home/claude/pou/POU_OperatingModel_v1.xlsx")
print("saved")
