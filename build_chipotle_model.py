"""
Chipotle Mexican Grill (NYSE: CMG) — Three-Statement Model + DCF
=================================================================
Generates a fully linked, formula-driven Excel model:

  Assumptions        all drivers in one place, with Bear / Base / Bull scenario switch
  Income Statement   restaurant unit build -> revenue -> restaurant-level margin -> net income
  Balance Sheet      condensed from the 10-K, projected from operating drivers, with balance check
  Cash Flow          indirect method, reconciles to ending cash on the balance sheet
  DCF                unlevered FCF, CAPM cost of equity, Gordon-growth terminal value,
                     implied share price vs. market, WACC x growth sensitivity

Historical data (FY2024A, FY2025A) comes from Chipotle's Q4/FY2025 earnings release,
filed with the SEC as Exhibit 99.1 to Form 8-K on Feb 3, 2026.
All dollar figures are in $ millions unless noted.

Usage:   python build_chipotle_model.py   ->  chipotle_3_statement_dcf.xlsx
Requires: openpyxl
"""

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

COMPANY = "Chipotle Mexican Grill (NYSE: CMG)"
OUTFILE = "chipotle_3_statement_dcf.xlsx"
SOURCE = ("Source: Chipotle Q4/FY2025 earnings release, SEC Form 8-K Ex. 99.1 (Feb 3, 2026), "
          "https://www.sec.gov/Archives/edgar/data/1058090/000105809026000007/cmg-20260203xex991.htm")

YEARS = ["2024A", "2025A", "2026E", "2027E", "2028E", "2029E", "2030E"]
COLS = [get_column_letter(i) for i in range(2, 2 + len(YEARS))]   # B..H
H1, H2 = COLS[0], COLS[1]                                           # B (2024A), C (2025A)
PCOLS = COLS[2:]                                                    # D..H projections
LAST = PCOLS[-1]                                                    # H

# ------------------------------------------------------------------ styles
FONT = "Arial"
BLUE = Font(name=FONT, color="0000FF")      # hardcoded inputs
BLACK = Font(name=FONT, color="000000")     # formulas on the same sheet
GREEN = Font(name=FONT, color="008000")     # links to other sheets
GREY = Font(name=FONT, italic=True, color="808080")
HEADER_FILL = PatternFill("solid", fgColor="451400")   # Chipotle-ish dark brown
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF")
KEY_FILL = PatternFill("solid", fgColor="FFFF00")
LIVE_FILL = PatternFill("solid", fgColor="F2F2F2")
TOP = Border(top=Side(style="thin"))

MONEY = '$#,##0.0;($#,##0.0);"-"'
PCT = '0.0%;(0.0%);"-"'
NUM0 = '#,##0;(#,##0);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
DAYS = '0.0" days"'
MMPER = '$#,##0.000;($#,##0.000);"-"'
PRICE = '$#,##0.00'
MULT = '0.0x'


def prev_col(col):
    return get_column_letter(_idx(col) - 1)


def _idx(col):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(col)


def put(ws, ref, value, fmt=None, font=None, bold=False, fill=None, border=False):
    cell = ws[ref]
    cell.value = value
    if font is None:
        if isinstance(value, str) and value.startswith("="):
            font = GREEN if "!" in value else BLACK
        else:
            font = BLUE
    if bold:
        font = Font(name=FONT, bold=True, color=font.color.rgb if font.color else "000000")
    cell.font = font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if border:
        cell.border = TOP
    return cell


def label(ws, row, text, bold=False, indent=0, border=False):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, bold=bold)
    c.alignment = Alignment(indent=indent)
    if border:
        c.border = TOP


def header(ws, row, title, cols=None, start_col=2):
    cols = YEARS if cols is None else cols
    ws.cell(row=row, column=1, value=title)
    for i, y in enumerate(cols):
        ws.cell(row=row, column=start_col + i, value=y)
    for c in range(1, start_col + len(cols)):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="left" if c == 1 else "center")


def row(ws, r, text, fn, fmt=MONEY, cols=PCOLS, bold=False, border=False, indent=0, fill=None):
    """Write fn(col, prev_col) across cols."""
    label(ws, r, text, bold=bold, indent=indent, border=border)
    for col in cols:
        put(ws, f"{col}{r}", fn(col, prev_col(col)), fmt, bold=bold, border=border, fill=fill)


def hist(ws, r, values, fmt=MONEY):
    """Hardcode historical values (2024A, 2025A); None skips a cell."""
    for col, v in zip([H1, H2], values):
        if v is not None:
            put(ws, f"{col}{r}", v, fmt)


def inputs(ws, r, text, values, fmt, indent=0):
    label(ws, r, text, indent=indent)
    for col, v in zip(PCOLS, values):
        put(ws, f"{col}{r}", v, fmt)


def setup(ws, title, width=46):
    ws["A1"] = f"{COMPANY} — {title}"
    ws["A1"].font = Font(name=FONT, bold=True, size=14)
    ws["A2"] = "($ in millions, except per-restaurant and per-share data; fiscal years end Dec 31)"
    ws["A2"].font = GREY
    ws.column_dimensions["A"].width = width
    for col in COLS:
        ws.column_dimensions[col].width = 12.5
    ws.sheet_view.showGridLines = False


wb = Workbook()

# ======================================================================
# 1. ASSUMPTIONS
# ======================================================================
A = wb.active
A.title = "Assumptions"
setup(A, "Assumptions", width=52)

header(A, 4, "Scenario control", cols=["Input"])
label(A, 5, "Active scenario (1 = Bear, 2 = Base, 3 = Bull)")
put(A, "B5", 2, "0", fill=KEY_FILL)
label(A, 6, "Active scenario name")
put(A, "B6", '=CHOOSE(B5,"Bear","Base","Bull")')

header(A, 8, "Unit growth & revenue")
inputs(A, 9, "New company-owned restaurant openings", [345, 360, 375, 385, 395], NUM0)
A["D9"].comment = Comment("FY2026 guidance: 350-370 openings incl. 10-15 international "
                          "partner-operated, i.e. ~340-355 company-owned.", "Model")
inputs(A, 10, "Permanent closures & relocations", [20, 20, 20, 20, 20], NUM0)
A["D10"].comment = Comment("FY2025: 13 permanent closures + 5 relocations.", "Model")


def scenario_block(r_live, text, bear, base, bull, fmt):
    """Live row picks Bear/Base/Bull from the three rows beneath it via CHOOSE."""
    label(A, r_live, f"{text} (live)", bold=True)
    for col in PCOLS:
        put(A, f"{col}{r_live}",
            f"=CHOOSE($B$5,{col}{r_live + 1},{col}{r_live + 2},{col}{r_live + 3})",
            fmt, fill=LIVE_FILL, bold=True)
    for off, (name, vals) in enumerate([("Bear", bear), ("Base", base), ("Bull", bull)], start=1):
        inputs(A, r_live + off, name, vals, fmt, indent=2)


scenario_block(11, "Revenue per average restaurant growth %",
               [-0.02, 0.00, 0.01, 0.02, 0.02],
               [0.00, 0.02, 0.03, 0.03, 0.03],
               [0.02, 0.04, 0.05, 0.05, 0.05], PCT)
A["D13"].comment = Comment("Base 2026 = 0%, consistent with management's 'about flat' "
                           "comparable sales guidance.", "Model")

header(A, 16, "Restaurant operating costs (% of revenue)")
put(A, "C16", "2025A")
A["C16"].fill, A["C16"].font = HEADER_FILL, HEADER_FONT
scenario_block(17, "Food, beverage & packaging %",
               [0.304, 0.304, 0.303, 0.302, 0.302],
               [0.300, 0.298, 0.296, 0.296, 0.296],
               [0.295, 0.293, 0.290, 0.288, 0.286], PCT)
scenario_block(21, "Labor %",
               [0.258, 0.258, 0.257, 0.256, 0.256],
               [0.254, 0.252, 0.250, 0.250, 0.250],
               [0.250, 0.246, 0.242, 0.239, 0.236], PCT)
inputs(A, 25, "Occupancy %", [0.053] * 5, PCT)
scenario_block(26, "Other operating costs %",
               [0.150, 0.150, 0.150, 0.149, 0.149],
               [0.148, 0.147, 0.146, 0.146, 0.146],
               [0.145, 0.142, 0.140, 0.138, 0.136], PCT)
# 2025A reference ratios (formulas, so the source of each assumption is visible)
put(A, "C17", "='Income Statement'!C14/'Income Statement'!C13", PCT)
put(A, "C21", "='Income Statement'!C15/'Income Statement'!C13", PCT)
put(A, "C25", "='Income Statement'!C16/'Income Statement'!C13", PCT)
put(A, "C26", "='Income Statement'!C17/'Income Statement'!C13", PCT)

header(A, 31, "Other operating & capital drivers")
other = [
    (32, "G&A % of revenue", [0.054, 0.053, 0.053, 0.052, 0.052], PCT,
     "='Income Statement'!C19/'Income Statement'!C13"),
    (33, "D&A % of revenue", [0.031] * 5, PCT,
     "='Income Statement'!C20/'Income Statement'!C13"),
    (34, "Pre-opening cost per new restaurant ($mm)", [0.150] * 5, MMPER,
     "='Income Statement'!C21/'Income Statement'!C7"),
    (35, "Impairment & closure costs % of revenue", [0.002] * 5, PCT,
     "='Income Statement'!C22/'Income Statement'!C13"),
    (36, "Capex % of revenue", [0.056] * 5, PCT,
     "=666.336/'Income Statement'!C13"),
    (37, "Stock-based compensation % of revenue", [0.010] * 5, PCT,
     "=119.543/'Income Statement'!C13"),
    (38, "Yield on beginning cash & investments", [0.040] * 5, PCT,
     "='Income Statement'!C24/AVERAGE('Balance Sheet'!B5:C5)"),
    (39, "Effective tax rate", [0.250] * 5, PCT,
     "='Income Statement'!C26/'Income Statement'!C25"),
    (40, "Share repurchases % of free cash flow", [0.90] * 5, PCT, None),
]
for r, text, vals, fmt, ref in other:
    inputs(A, r, text, vals, fmt)
    if ref:
        put(A, f"C{r}", ref, fmt)
A["C34"].comment = Comment("FY2025 pre-opening costs / 334 company-owned openings.", "Model")

A["C36"].comment = Comment("FY2025 purchases of leasehold improvements, property & equipment "
                           "= $666.3mm (cash flow statement).", "Model")
A["C37"].comment = Comment("FY2025 stock-based compensation = $119.5mm (cash flow statement).", "Model")
A["D39"].comment = Comment("FY2026 guidance: underlying effective tax rate of 24-26%.", "Model")
A["D40"].comment = Comment("Chipotle pays no dividend and returns cash via buybacks "
                           "($2.4bn repurchased in FY2025).", "Model")

header(A, 42, "Working capital & lease drivers")
wc = [
    (43, "Accounts receivable days (on revenue)", 4.8, DAYS,
     "='Balance Sheet'!C6/'Income Statement'!C13*365"),
    (44, "Inventory days (on food cost)", 5.1, DAYS,
     "='Balance Sheet'!C7/'Income Statement'!C14*365"),
    (45, "Prepaid & other current assets % of revenue", 0.018, PCT,
     "='Balance Sheet'!C8/'Income Statement'!C13"),
    (46, "Accounts payable days (on food cost)", 22.0, DAYS,
     "='Balance Sheet'!C14/'Income Statement'!C14*365"),
    (47, "Accrued payroll & liabilities % of revenue", 0.036, PCT,
     "='Balance Sheet'!C15/'Income Statement'!C13"),
    (48, "Unearned revenue % of revenue", 0.020, PCT,
     "='Balance Sheet'!C16/'Income Statement'!C13"),
    (49, "Operating lease asset per restaurant ($mm)", 1.104, MMPER,
     "='Balance Sheet'!C10/'Income Statement'!C9"),
    (50, "Operating lease liability per restaurant ($mm)", 1.256, MMPER,
     "='Balance Sheet'!C17/'Income Statement'!C9"),
]
for r, text, v, fmt, ref in wc:
    inputs(A, r, text, [v] * 5, fmt)
    put(A, f"C{r}", ref, fmt)

header(A, 52, "Valuation inputs", cols=["Input"])
val = [
    (53, "Risk-free rate (10-year US Treasury)", 0.049, PCT,
     "10-yr UST traded roughly 4.7%-5.1% in Oct-Nov 2026. Update to the current yield."),
    (54, "Equity risk premium", 0.050, PCT, "Common practitioner assumption (~4.5%-5.5%)."),
    (55, "Levered beta", 1.00, "0.00", "CMG beta ~1.0 per Google Finance (Nov 2026)."),
    (56, "Terminal growth rate", 0.030, PCT,
     "Long-run nominal GDP-type growth; Chipotle still has unit growth runway in 2030."),
    (57, "Shares outstanding at 12/31/2025 (mm)", 1304.360, NUM1,
     "1,304,360 thousand shares issued per the 12/31/2025 balance sheet."),
    (58, "Market share price ($, reference)", 34.43, PRICE,
     "NYSE close on Nov 6, 2026. Update before using."),
    (59, "Exit multiple: EV / 2030E EBITDA (cross-check)", 14.0, MULT,
     "Cross-check only. Set this from a peer comps table (e.g., large-cap restaurant operators)."),
]
for r, text, v, fmt, note in val:
    label(A, r, text)
    c = put(A, f"B{r}", v, fmt)
    c.comment = Comment(note, "Model")
    if r in (55, 56):
        c.fill = KEY_FILL

notes = [
    "Legend: blue = hardcoded input · black = formula · green = link to another sheet · "
    "yellow = key lever · grey rows = live scenario values picked by the switch in B5.",
    "Column C on this sheet shows the FY2025 actual ratio behind each assumption.",
    "Simplifications: balance sheet condensed to 12 lines; cash includes short- and long-term "
    "investments and restricted cash; no funded debt, so WACC = cost of equity.",
    "Operating leases are treated as operating costs (rent sits in occupancy expense), so lease "
    "liabilities are not added to debt and lease balance changes are excluded from free cash flow.",
    "Stock-based compensation is treated as a real cost in the DCF (not added back to free cash flow).",
    SOURCE,
]
for i, n in enumerate(notes):
    A[f"A{61 + i}"] = n
    A[f"A{61 + i}"].font = GREY
A.freeze_panes = "B5"

# ======================================================================
# 2. INCOME STATEMENT
# ======================================================================
IS = wb.create_sheet("Income Statement")
setup(IS, "Income Statement")
header(IS, 4, "Income Statement")

label(IS, 5, "Restaurant unit build (company-owned)", bold=True)
row(IS, 6, "Beginning restaurants", lambda c, p: f"={p}9", NUM0, indent=1)
row(IS, 7, "New openings", lambda c, p: f"=Assumptions!{c}9", NUM0, indent=1)
row(IS, 8, "Closures & relocations", lambda c, p: f"=-Assumptions!{c}10", NUM0, indent=1)
hist(IS, 9, [3726, 4042], NUM0)
hist(IS, 7, [None, 334], NUM0)
hist(IS, 8, [None, -18], NUM0)
put(IS, f"{H2}6", f"={H1}9", NUM0)
row(IS, 9, "Ending restaurants", lambda c, p: f"=SUM({c}6:{c}8)", NUM0, bold=True, border=True)
row(IS, 10, "Average restaurants", lambda c, p: f"=AVERAGE({p}9,{c}9)", NUM1, cols=[H2] + PCOLS, indent=1)
put(IS, f"{H2}11", f"={H2}13/{H2}10", MMPER)
row(IS, 11, "Revenue per average restaurant ($mm)",
    lambda c, p: f"={p}11*(1+Assumptions!{c}11)", MMPER, indent=1)

hist_is = {
    13: [11313.853, 11925.601],
    14: [3374.516, 3526.992],
    15: [2789.789, 2991.680],
    16: [563.374, 624.898],
    17: [1568.482, 1755.824],
    19: [697.483, 652.017],
    20: [335.030, 361.382],
    21: [41.897, 49.507],
    22: [26.949, 27.503],
    24: [93.897, 73.721],
    26: [476.120, 473.758],
}
for r, v in hist_is.items():
    hist(IS, r, v)

row(IS, 13, "Total revenue", lambda c, p: f"={c}10*{c}11", bold=True)
row(IS, 14, "Food, beverage & packaging", lambda c, p: f"={c}13*Assumptions!{c}17", indent=1)
row(IS, 15, "Labor", lambda c, p: f"={c}13*Assumptions!{c}21", indent=1)
row(IS, 16, "Occupancy", lambda c, p: f"={c}13*Assumptions!{c}25", indent=1)
row(IS, 17, "Other operating costs", lambda c, p: f"={c}13*Assumptions!{c}26", indent=1)
row(IS, 18, "Restaurant-level operating profit", lambda c, p: f"={c}13-SUM({c}14:{c}17)",
    cols=COLS, bold=True, border=True)
row(IS, 19, "General & administrative", lambda c, p: f"={c}13*Assumptions!{c}32", indent=1)
row(IS, 20, "Depreciation & amortization", lambda c, p: f"={c}13*Assumptions!{c}33", indent=1)
row(IS, 21, "Pre-opening costs", lambda c, p: f"={c}7*Assumptions!{c}34", indent=1)
row(IS, 22, "Impairment, closure costs & asset disposals",
    lambda c, p: f"={c}13*Assumptions!{c}35", indent=1)
row(IS, 23, "Income from operations (EBIT)", lambda c, p: f"={c}18-SUM({c}19:{c}22)",
    cols=COLS, bold=True, border=True)
# interest on BEGINNING cash balance avoids a circular reference
row(IS, 24, "Interest & other income, net",
    lambda c, p: f"='Balance Sheet'!{p}5*Assumptions!{c}38", indent=1)
row(IS, 25, "Income before income taxes", lambda c, p: f"={c}23+{c}24", cols=COLS, border=True)
row(IS, 26, "Provision for income taxes", lambda c, p: f"={c}25*Assumptions!{c}39", indent=1)
row(IS, 27, "Net income", lambda c, p: f"={c}25-{c}26", cols=COLS, bold=True, border=True)

label(IS, 29, "Key metrics", bold=True)
row(IS, 30, "Revenue growth", lambda c, p: f"=IF({p}13=0,0,{c}13/{p}13-1)", PCT,
    cols=COLS[1:], indent=1)
row(IS, 31, "Restaurant-level margin", lambda c, p: f"=IF({c}13=0,0,{c}18/{c}13)", PCT,
    cols=COLS, indent=1)
row(IS, 32, "EBITDA (EBIT + D&A)", lambda c, p: f"={c}23+{c}20", cols=COLS, indent=1)
row(IS, 33, "Operating margin", lambda c, p: f"=IF({c}13=0,0,{c}23/{c}13)", PCT, cols=COLS, indent=1)
row(IS, 34, "Net margin", lambda c, p: f"=IF({c}13=0,0,{c}27/{c}13)", PCT, cols=COLS, indent=1)
IS["A36"] = ("Check: FY2025A EBIT ties to reported $1,935.8mm, net income to $1,535.8mm, "
             "restaurant-level margin to 25.4%.")
IS["A36"].font = GREY
IS["A37"] = SOURCE
IS["A37"].font = GREY
IS.freeze_panes = "B5"

# ======================================================================
# 3. BALANCE SHEET (condensed)
# ======================================================================
BS = wb.create_sheet("Balance Sheet")
setup(BS, "Balance Sheet (condensed)")
header(BS, 4, "Balance Sheet")

hist_bs = {
    5: [2320.782, 1281.623],   # cash + ST inv + LT inv + restricted cash
    6: [143.963, 156.466],
    7: [48.942, 49.508],
    8: [164.767, 211.843],     # prepaid & other + income tax receivable
    9: [2390.126, 2679.361],
    10: [4000.127, 4463.010],
    11: [135.667, 152.720],    # other assets + goodwill
    14: [210.695, 212.813],
    15: [441.660, 432.574],    # accrued payroll & benefits + accrued liabilities
    16: [238.577, 240.375],
    17: [4540.618, 5075.814],  # current + long-term operating lease liabilities
    18: [117.278, 202.348],    # deferred income taxes + other liabilities
    21: [3655.546, 2830.607],
}
for r, v in hist_bs.items():
    hist(BS, r, v)

row(BS, 5, "Cash, investments & restricted cash", lambda c, p: f"='Cash Flow'!{c}26", indent=1)
BS["C5"].comment = Comment("Cash $350.5 + ST investments $698.6 + LT investments $197.1 "
                           "+ restricted cash $35.4", "Model")
row(BS, 6, "Accounts receivable", lambda c, p: f"='Income Statement'!{c}13*Assumptions!{c}43/365", indent=1)
row(BS, 7, "Inventory", lambda c, p: f"='Income Statement'!{c}14*Assumptions!{c}44/365", indent=1)
row(BS, 8, "Prepaid expenses & other current assets",
    lambda c, p: f"='Income Statement'!{c}13*Assumptions!{c}45", indent=1)
row(BS, 9, "Leasehold improvements, property & equipment, net",
    lambda c, p: f"={p}9-'Cash Flow'!{c}19-'Income Statement'!{c}20-'Income Statement'!{c}22", indent=1)
row(BS, 10, "Operating lease assets",
    lambda c, p: f"='Income Statement'!{c}9*Assumptions!{c}49", indent=1)
row(BS, 11, "Other assets & goodwill", lambda c, p: f"={p}11", indent=1)
row(BS, 12, "Total assets", lambda c, p: f"=SUM({c}5:{c}11)", cols=COLS, bold=True, border=True)

row(BS, 14, "Accounts payable", lambda c, p: f"='Income Statement'!{c}14*Assumptions!{c}46/365", indent=1)
row(BS, 15, "Accrued payroll & accrued liabilities",
    lambda c, p: f"='Income Statement'!{c}13*Assumptions!{c}47", indent=1)
row(BS, 16, "Unearned revenue (gift cards, rewards)",
    lambda c, p: f"='Income Statement'!{c}13*Assumptions!{c}48", indent=1)
row(BS, 17, "Operating lease liabilities (current + long-term)",
    lambda c, p: f"='Income Statement'!{c}9*Assumptions!{c}50", indent=1)
row(BS, 18, "Deferred income taxes & other liabilities", lambda c, p: f"={p}18", indent=1)
row(BS, 19, "Total liabilities", lambda c, p: f"=SUM({c}14:{c}18)", cols=COLS, bold=True, border=True)

row(BS, 21, "Total shareholders' equity",
    lambda c, p: f"={p}21+'Income Statement'!{c}27+'Cash Flow'!{c}9+'Cash Flow'!{c}22", indent=1)
row(BS, 22, "Total liabilities & equity", lambda c, p: f"={c}19+{c}21", cols=COLS, bold=True, border=True)

row(BS, 24, "Balance check (must be 0)", lambda c, p: f"=ROUND({c}12-{c}22,3)", cols=COLS,
    bold=True, fill=KEY_FILL)
BS["A26"] = ("Historical totals tie to the reported balance sheet: total assets $9,204.4mm (2024) "
             "and $8,994.5mm (2025).")
BS["A26"].font = GREY
BS["A27"] = SOURCE
BS["A27"].font = GREY
BS.freeze_panes = "B5"

# ======================================================================
# 4. CASH FLOW STATEMENT (projections)
# ======================================================================
CF = wb.create_sheet("Cash Flow")
setup(CF, "Cash Flow Statement")
header(CF, 4, "Cash Flow Statement")
CF[f"{H1}4"].value = ""
CF[f"{H2}4"].value = ""

label(CF, 5, "Operating activities", bold=True)
row(CF, 6, "Net income", lambda c, p: f"='Income Statement'!{c}27", indent=1)
row(CF, 7, "Depreciation & amortization", lambda c, p: f"='Income Statement'!{c}20", indent=1)
row(CF, 8, "Impairment & closure costs (non-cash)", lambda c, p: f"='Income Statement'!{c}22", indent=1)
row(CF, 9, "Stock-based compensation",
    lambda c, p: f"='Income Statement'!{c}13*Assumptions!{c}37", indent=1)
row(CF, 10, "(Increase) / decrease in accounts receivable",
    lambda c, p: f"='Balance Sheet'!{p}6-'Balance Sheet'!{c}6", indent=1)
row(CF, 11, "(Increase) / decrease in inventory",
    lambda c, p: f"='Balance Sheet'!{p}7-'Balance Sheet'!{c}7", indent=1)
row(CF, 12, "(Increase) / decrease in prepaid & other",
    lambda c, p: f"='Balance Sheet'!{p}8-'Balance Sheet'!{c}8", indent=1)
row(CF, 13, "Increase / (decrease) in accounts payable",
    lambda c, p: f"='Balance Sheet'!{c}14-'Balance Sheet'!{p}14", indent=1)
row(CF, 14, "Increase / (decrease) in accrued liabilities",
    lambda c, p: f"='Balance Sheet'!{c}15-'Balance Sheet'!{p}15", indent=1)
row(CF, 15, "Increase / (decrease) in unearned revenue",
    lambda c, p: f"='Balance Sheet'!{c}16-'Balance Sheet'!{p}16", indent=1)
row(CF, 16, "Net change in operating lease balances",
    lambda c, p: (f"=('Balance Sheet'!{c}17-'Balance Sheet'!{p}17)"
                  f"-('Balance Sheet'!{c}10-'Balance Sheet'!{p}10)"), indent=1)
row(CF, 17, "Cash flow from operations", lambda c, p: f"=SUM({c}6:{c}16)", bold=True, border=True)

label(CF, 18, "Investing activities", bold=True)
row(CF, 19, "Capital expenditures", lambda c, p: f"=-'Income Statement'!{c}13*Assumptions!{c}36", indent=1)
row(CF, 20, "Cash flow from investing", lambda c, p: f"={c}19", bold=True, border=True)

label(CF, 21, "Financing activities", bold=True)
row(CF, 22, "Share repurchases",
    lambda c, p: f"=-MAX(0,{c}17+{c}19)*Assumptions!{c}40", indent=1)
row(CF, 23, "Cash flow from financing", lambda c, p: f"={c}22", bold=True, border=True)

row(CF, 24, "Net change in cash", lambda c, p: f"={c}17+{c}20+{c}23", bold=True)
row(CF, 25, "Beginning cash, investments & restricted cash",
    lambda c, p: f"='Balance Sheet'!{p}5", indent=1)
row(CF, 26, "Ending cash, investments & restricted cash",
    lambda c, p: f"={c}25+{c}24", bold=True, border=True)

label(CF, 28, "Memo", bold=True)
row(CF, 29, "Free cash flow (CFO − capex)", lambda c, p: f"={c}17+{c}19", indent=1)
CF["A31"] = ("Reference: FY2025A reported CFO $2,113.9mm, capex $666.3mm, "
             "share repurchases $2,425.5mm.")
CF["A31"].font = GREY
CF.freeze_panes = "B5"

# ======================================================================
# 5. DCF
# ======================================================================
D = wb.create_sheet("DCF")
setup(D, "Discounted Cash Flow Valuation", width=50)
header(D, 4, "Unlevered free cash flow")
D[f"{H1}4"].value = ""
D[f"{H2}4"].value = ""

row(D, 5, "EBIT", lambda c, p: f"='Income Statement'!{c}23")
row(D, 6, "Less: taxes on EBIT", lambda c, p: f"=-{c}5*Assumptions!{c}39", indent=1)
row(D, 7, "NOPAT", lambda c, p: f"={c}5+{c}6", bold=True, border=True)
row(D, 8, "Plus: D&A", lambda c, p: f"='Income Statement'!{c}20", indent=1)
row(D, 9, "Plus: impairment (non-cash)", lambda c, p: f"='Income Statement'!{c}22", indent=1)
row(D, 10, "Less: capex", lambda c, p: f"='Cash Flow'!{c}19", indent=1)
row(D, 11, "Less: increase in net working capital",
    lambda c, p: f"=SUM('Cash Flow'!{c}10:'Cash Flow'!{c}15)", indent=1)
row(D, 12, "Unlevered free cash flow", lambda c, p: f"=SUM({c}7:{c}11)", bold=True, border=True)
label(D, 13, "Discount period (years from 12/31/2025)", indent=1)
put(D, f"{PCOLS[0]}13", "=1", "0.0")
for col in PCOLS[1:]:
    put(D, f"{col}13", f"={prev_col(col)}13+1", "0.0")
row(D, 14, "Discount factor", lambda c, p: f"=1/(1+$D$22)^{c}13", "0.000", indent=1)
row(D, 15, "PV of free cash flow", lambda c, p: f"={c}12*{c}14", bold=True)

header(D, 17, "Cost of capital", cols=["", "", "Value"])
coc = [
    (18, "Risk-free rate", "=Assumptions!B53", PCT),
    (19, "Equity risk premium", "=Assumptions!B54", PCT),
    (20, "Levered beta", "=Assumptions!B55", "0.00"),
    (21, "Cost of equity (CAPM: Rf + β × ERP)", "=D18+D20*D19", PCT),
    (22, "WACC (= cost of equity; CMG has no funded debt)", "=D21", PCT),
]
for r, text, f, fmt in coc:
    label(D, r, text, bold=(r == 22))
    put(D, f"D{r}", f, fmt, bold=(r == 22), fill=KEY_FILL if r == 22 else None)

header(D, 24, "Valuation (as of 12/31/2025 balance sheet)", cols=["", "", "Value"])
valuation = [
    (25, "Sum of PV of free cash flows (2026E-2030E)", "=SUM(D15:H15)", MONEY),
    (26, "Terminal growth rate", "=Assumptions!B56", PCT),
    (27, "Terminal value (Gordon growth, end of 2030E)", f"={LAST}12*(1+D26)/(D22-D26)", MONEY),
    (28, "PV of terminal value", f"=D27*{LAST}14", MONEY),
    (29, "Enterprise value", "=D25+D28", MONEY),
    (30, "Plus: cash, investments & restricted cash", "='Balance Sheet'!C5", MONEY),
    (31, "Equity value", "=D29+D30", MONEY),
    (32, "Shares outstanding (mm)", "=Assumptions!B57", NUM1),
    (33, "Implied value per share", "=IF(D32=0,0,D31/D32)", PRICE),
    (34, "Market share price (reference)", "=Assumptions!B58", PRICE),
    (35, "Implied upside / (downside)", "=IF(D34=0,0,D33/D34-1)", PCT),
    (37, "Terminal value as % of enterprise value", "=IF(D29=0,0,D28/D29)", PCT),
    (38, "Implied EV / 2026E EBITDA", "=IF('Income Statement'!D32=0,0,D29/'Income Statement'!D32)", MULT),
    (39, "Market-implied EV / 2026E EBITDA",
     "=IF('Income Statement'!D32=0,0,(D34*D32-D30)/'Income Statement'!D32)", MULT),
    (40, "Exit multiple: EV / 2030E EBITDA", "=Assumptions!B59", MULT),
    (41, "Implied value per share (exit multiple method)",
     f"=IF(D32=0,0,(D25+'Income Statement'!{LAST}32*D40*{LAST}14+D30)/D32)", PRICE),
    (42, "Perpetual growth implied by exit multiple",
     f"=('Income Statement'!{LAST}32*D40*D22-{LAST}12)/('Income Statement'!{LAST}32*D40+{LAST}12)", PCT),
    (43, "Perpetual growth implied by market share price",
     f"=(((D34*D32-D30)-D25)/{LAST}14*D22-{LAST}12)/(((D34*D32-D30)-D25)/{LAST}14+{LAST}12)", PCT),
]
for r, text, f, fmt in valuation:
    bold = r in (29, 31, 33, 35, 43)
    label(D, r, text, bold=bold)
    put(D, f"D{r}", f, fmt, bold=bold, fill=KEY_FILL if r in (33, 43) else None)
label(D, 36, "Cross-checks", bold=True)
D["A44"] = '=CONCATENATE("Scenario shown: ",Assumptions!B6)'
D["A44"].font = Font(name=FONT, bold=True, color="008000")

# ---- Sensitivity: implied share price, WACC (rows) x terminal growth (columns)
D["A46"] = "Sensitivity: implied value per share"
D["A46"].font = Font(name=FONT, bold=True)
D["A47"] = "Rows = WACC, columns = terminal growth rate"
D["A47"].font = GREY
steps = [-0.01, -0.005, 0, 0.005, 0.01]
for j, dg in enumerate(steps):
    col = PCOLS[j]
    put(D, f"{col}48", f"=$D$26+({dg})", PCT, font=Font(name=FONT, bold=True))
for i, dw in enumerate(steps):
    r = 49 + i
    put(D, f"C{r}", f"=$D$22+({dw})", PCT, font=Font(name=FONT, bold=True))
    for j in range(len(steps)):
        col = PCOLS[j]
        w, g = f"$C{r}", f"{col}$48"
        formula = (f"=(SUMPRODUCT($D$12:$H$12/(1+{w})^$D$13:$H$13)"
                   f"+$H$12*(1+{g})/({w}-{g})/(1+{w})^$H$13+$D$30)/$D$32")
        put(D, f"{col}{r}", formula, PRICE)
D["F51"].fill = KEY_FILL  # base case
D["A55"] = ("Notes: free cash flows discounted at year-end from the 12/31/2025 balance sheet date. "
            "SBC is treated as a cash-equivalent cost (not added back).")
D["A55"].font = GREY
D.freeze_panes = "B5"

wb.save(OUTFILE)
print(f"Saved {OUTFILE}")
