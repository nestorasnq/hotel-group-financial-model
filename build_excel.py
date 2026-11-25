"""Builds output/hotel_group_model.xlsx from the files in data/.
Every result in the workbook is an Excel formula; the inputs are typed in blue,
links to another sheet in green.
Step 1: KPIs. Step 2: Data, Mapping, Contract, P&L, Checks.
Step 3: Fixed assets (register, net book value, journal entry), Returns (working capital, days, ROCE)."""
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BLUE = Font(name="Arial", size=10, color="0000FF")
BLACK = Font(name="Arial", size=10)
GREEN = Font(name="Arial", size=10, color="008000")
RED_BOLD = Font(name="Arial", size=10, bold=True, color="C00000")
BOLD = Font(name="Arial", size=10, bold=True)
TITLE = Font(name="Arial", size=13, bold=True)
NOTE = Font(name="Arial", size=9, italic=True, color="555555")
HEAD_FILL = PatternFill("solid", start_color="DCEBF1")
TOP = Border(top=Side(style="thin"))

EUR = '€#,##0;(€#,##0);"-"'
EUR2 = '€#,##0.00;(€#,##0.00);"-"'
NUM = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'


def read_hotels(path="data/hotels.csv"):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def build_kpis(ws, hotels):
    ws["A1"] = "Step 1 · Hotel KPIs, 2025 (made-up hotels)"
    ws["A1"].font = TITLE
    ws["A2"] = "Blue = input you can change. Black = formula. Group KPIs: add up first, then divide."
    ws["A2"].font = NOTE

    heads = ["Hotel", "Location", "Rooms", "Days open", "Occupancy", "ADR (€)",
             "Available room-nights", "Sold room-nights", "Rooms revenue (€)", "RevPAR (€)"]
    for c, h in enumerate(heads, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = BOLD
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    first = 5
    for i, h in enumerate(hotels):
        r = first + i
        ws.cell(r, 1, h["hotel"]).font = BLACK
        ws.cell(r, 2, h["location"]).font = BLACK
        ws.cell(r, 3, int(h["rooms"])).font = BLUE
        ws.cell(r, 4, int(h["days_open"])).font = BLUE
        ws.cell(r, 5, float(h["occupancy"])).font = BLUE
        ws.cell(r, 6, float(h["adr_eur"])).font = BLUE
        ws.cell(r, 7, f"=C{r}*D{r}")            # available room-nights
        ws.cell(r, 8, f"=G{r}*E{r}")            # sold room-nights
        ws.cell(r, 9, f"=H{r}*F{r}")            # rooms revenue
        ws.cell(r, 10, f"=IF(G{r}=0,0,I{r}/G{r})")  # RevPAR
    last = first + len(hotels) - 1
    g = last + 1  # group row

    ws.cell(g, 1, "Group").font = BOLD
    ws.cell(g, 2, "add up, then divide").font = NOTE
    ws.cell(g, 3, f"=SUM(C{first}:C{last})")
    ws.cell(g, 5, f"=IF(G{g}=0,0,H{g}/G{g})")   # group occupancy
    ws.cell(g, 6, f"=IF(H{g}=0,0,I{g}/H{g})")   # group ADR
    ws.cell(g, 7, f"=SUM(G{first}:G{last})")
    ws.cell(g, 8, f"=SUM(H{first}:H{last})")
    ws.cell(g, 9, f"=SUM(I{first}:I{last})")
    ws.cell(g, 10, f"=IF(G{g}=0,0,I{g}/G{g})")  # group RevPAR

    fmts = {3: NUM, 4: NUM, 5: PCT, 6: EUR, 7: NUM, 8: NUM, 9: EUR, 10: EUR2}
    for r in range(first, g + 1):
        for c, f in fmts.items():
            ws.cell(r, c).number_format = f
            if r == g:
                ws.cell(r, c).font = BOLD
                ws.cell(r, c).border = TOP
    ws.cell(g, 1).border = TOP
    ws.cell(g, 2).border = TOP

    # Comparison and checks
    c0 = g + 2
    ws.cell(c0, 1, "For comparison only (the wrong way)").font = BOLD
    ws.cell(c0 + 1, 1, "Simple average occupancy")
    ws.cell(c0 + 1, 5, f"=AVERAGE(E{first}:E{last})").number_format = PCT
    ws.cell(c0 + 2, 1, "Simple average ADR")
    ws.cell(c0 + 2, 6, f"=AVERAGE(F{first}:F{last})").number_format = EUR2

    k0 = c0 + 4
    ws.cell(k0, 1, "Checks (0 = pass)").font = BOLD
    ws.cell(k0 + 1, 1, "Group RevPAR − group occupancy × group ADR")
    ws.cell(k0 + 1, 10, f"=ROUND(J{g}-E{g}*F{g},6)").number_format = '0.000000'
    ws.cell(k0 + 2, 1, "Sum of hotel revenues − group revenue")
    ws.cell(k0 + 2, 10, f"=ROUND(SUMPRODUCT(C{first}:C{last},D{first}:D{last},E{first}:E{last},F{first}:F{last})-I{g},2)").number_format = '0.00'

    n0 = k0 + 4
    ws.cell(n0, 1, "Source: hotel numbers are made up (round numbers in a realistic range for Greece). "
                   "Formulas: available = rooms × days open; sold = available × occupancy; "
                   "revenue = sold × ADR; RevPAR = revenue ÷ available.").font = NOTE

    widths = {"A": 34, "B": 28, "C": 8, "D": 9, "E": 11, "F": 9, "G": 12, "H": 12, "I": 15, "J": 11}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[4].height = 30
    ws.freeze_panes = "C5"
    return {"first": first, "last": last, "group": g, "checks": (k0 + 1, k0 + 2)}


# ---------------------------------------------------------------- Step 2
PNL_LINES = [  # (label, kind) kind: "map" = SUMIFS from Data; others are formulas
    ("Rooms revenue", "map"), ("F&B and other revenue", "map"), ("Total revenue", "sum_rev"),
    ("Staff costs", "map"), ("Food and beverage cost", "map"), ("Utilities and energy", "map"),
    ("Sales and marketing", "map"), ("Repairs and maintenance", "map"), ("Other operating costs", "map"),
    ("Total operating costs", "sum_cost"), ("GOP (gross operating profit)", "gop"), ("GOP margin", "gop_m"),
    ("Management fee (as invoiced)", "map_fee"), ("EBITDA", "ebitda"), ("EBITDA margin", "ebitda_m"),
    ("Depreciation", "map"), ("EBIT", "ebit"), ("Share of group GOP", "gop_share"),
]
HOTELS = ["City", "Resort", "Island"]


def read_csv(path):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def head_row(ws, row, heads):
    for c, h in enumerate(heads, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = BOLD
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def build_data(ws, ledger):
    ws["A1"] = "Step 2 · The accounting export, 2025 (made up; one line per hotel and account)"
    ws["A1"].font = TITLE
    ws["A2"] = ("Columns A–E are the export as received (blue). F–I are formulas: F counts this hotel+account so far, "
                "G flags a repeat, H looks up the P&L line, I is the amount used (0 for a duplicate).")
    ws["A2"].font = NOTE
    head_row(ws, 4, ["Line", "Hotel", "Account", "Account name", "Amount (€)",
                     "Count so far", "Duplicate?", "P&L line (from Mapping)", "Amount used (€)"])
    first = 5
    for i, r in enumerate(ledger):
        row = first + i
        ws.cell(row, 1, int(r["line"])).font = BLUE
        ws.cell(row, 2, r["hotel"]).font = BLUE
        ws.cell(row, 3, int(r["account"])).font = BLUE
        ws.cell(row, 4, r["account_name"]).font = BLUE
        ws.cell(row, 5, float(r["amount_eur"])).font = BLUE
        ws.cell(row, 6, f"=COUNTIFS($B${first}:B{row},B{row},$C${first}:C{row},C{row})")
        ws.cell(row, 7, f'=IF(F{row}>1,"DUPLICATE","")')
        ws.cell(row, 8, f'=IFERROR(INDEX(Mapping!$C$5:$C$19,MATCH(C{row},Mapping!$A$5:$A$19,0)),"UNMAPPED")').font = GREEN
        ws.cell(row, 9, f'=IF(G{row}="DUPLICATE",0,E{row})')
        ws.cell(row, 5).number_format = NUM
        ws.cell(row, 9).number_format = NUM
    last = first + len(ledger) - 1
    widths = {"A": 6, "B": 9, "C": 9, "D": 30, "E": 13, "F": 8, "G": 12, "H": 28, "I": 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[4].height = 30
    ws.freeze_panes = "A5"
    return first, last


def build_mapping(ws, mapping):
    ws["A1"] = "Step 2 · Mapping: which account feeds which P&L line"
    ws["A1"].font = TITLE
    ws["A2"] = "Every account in the export must appear here once. Blue = input you can change."
    ws["A2"].font = NOTE
    head_row(ws, 4, ["Account", "Account name", "P&L line"])
    for i, m in enumerate(mapping):
        r = 5 + i
        ws.cell(r, 1, int(m["account"])).font = BLUE
        ws.cell(r, 2, m["account_name"]).font = BLUE
        ws.cell(r, 3, m["pnl_line"]).font = BLUE
    for col, w in {"A": 9, "B": 32, "C": 28}.items():
        ws.column_dimensions[col].width = w


def build_contract(ws, terms):
    ws["A1"] = "Step 2 · Management agreement (made-up contract terms)"
    ws["A1"].font = TITLE
    head_row(ws, 4, ["Term", "Value", "Source / note"])
    t = {r["term"]: r for r in terms}
    rows = [("Base fee, % of total revenue", "base_fee_pct_of_revenue"),
            ("Incentive fee, % of GOP", "incentive_fee_pct_of_gop")]
    for i, (label, key) in enumerate(rows):
        r = 5 + i
        ws.cell(r, 1, label)
        c = ws.cell(r, 2, float(t[key]["value"]))
        c.font = BLUE
        c.number_format = PCT
        c.fill = PatternFill("solid", start_color="FFFF00")
        ws.cell(r, 3, t[key]["note"]).font = NOTE
        ws.cell(r, 3).alignment = Alignment(wrap_text=False)
    ws["A8"] = "Fee per contract = base % × total revenue + incentive % × GOP. Real contracts often use GOP after the base fee (adjusted GOP); this model uses plain GOP."
    ws["A8"].font = NOTE
    for col, w in {"A": 30, "B": 14, "C": 14, "D": 14, "E": 14, "F": 16}.items():
        ws.column_dimensions[col].width = w


def build_fee_parts(ws, prow):
    """Contract sheet, rows 10-13: the fee each hotel should pay under the contract, in its two parts."""
    ws["A10"] = "Fee per contract, in parts"
    ws["A10"].font = BOLD
    head_row(ws, 11, ["Hotel", "Total revenue (€)", "GOP (€)", "Base fee (€)", "Incentive fee (€)", "Fee per contract (€)"])
    for i, h in enumerate(HOTELS):
        r = 12 + i
        ws.cell(r, 1, h)
        for col, label in ((2, "Total revenue"), (3, "GOP (gross operating profit)")):
            c = ws.cell(r, col, f"=INDEX('P&L'!$B${prow[label]}:$D${prow[label]},MATCH($A{r},'P&L'!$B$3:$D$3,0))")
            c.font = GREEN
            c.number_format = EUR
        ws.cell(r, 4, f"=$B$5*B{r}").number_format = EUR
        ws.cell(r, 5, f"=$B$6*C{r}").number_format = EUR
        ws.cell(r, 6, f"=D{r}+E{r}").number_format = EUR


def build_pnl(ws, d_first, d_last):
    ws["A1"] = "Step 2 · Hotel P&L, 2025 (€)"
    ws["A1"].font = TITLE
    ws["A2"] = "Each mapped line is one SUMIFS over Data: amount used, where hotel = column head and P&L line = row label."
    ws["A2"].font = NOTE
    head_row(ws, 3, ["P&L line"] + HOTELS + ["Group"])
    rows = {}
    r = 4
    amt = f"Data!$I${d_first}:$I${d_last}"
    hot = f"Data!$B${d_first}:$B${d_last}"
    lin = f"Data!$H${d_first}:$H${d_last}"
    for label, kind in PNL_LINES:
        rows[label] = r
        ws.cell(r, 1, label)
        for j in range(4):
            col = 2 + j
            L = "BCDE"[j]
            if kind in ("map", "map_fee"):
                key = "Management fee" if kind == "map_fee" else label
                f = f'=SUMIFS({amt},{hot},{L}$3,{lin},"{key}")' if j < 3 else f"=SUM(B{r}:D{r})"
            elif kind == "sum_rev":
                f = f"={L}{rows['Rooms revenue']}+{L}{rows['F&B and other revenue']}"
            elif kind == "sum_cost":
                f = f"=SUM({L}{rows['Staff costs']}:{L}{rows['Other operating costs']})"
            elif kind == "gop":
                f = f"={L}{rows['Total revenue']}-{L}{rows['Total operating costs']}"
            elif kind == "gop_m":
                f = f"=IF({L}{rows['Total revenue']}=0,0,{L}{rows['GOP (gross operating profit)']}/{L}{rows['Total revenue']})"
            elif kind == "ebitda":
                f = f"={L}{rows['GOP (gross operating profit)']}-{L}{rows['Management fee (as invoiced)']}"
            elif kind == "ebitda_m":
                f = f"=IF({L}{rows['Total revenue']}=0,0,{L}{rows['EBITDA']}/{L}{rows['Total revenue']})"
            elif kind == "ebit":
                f = f"={L}{rows['EBITDA']}-{L}{rows['Depreciation']}"
            elif kind == "gop_share":
                f = f"=IF($E{rows['GOP (gross operating profit)']}=0,0,{L}{rows['GOP (gross operating profit)']}/$E{rows['GOP (gross operating profit)']})"
            c = ws.cell(r, col, f)
            c.number_format = PCT if (kind.endswith("_m") or kind == "gop_share") else EUR
            if kind in ("map", "map_fee") and j < 3:
                c.font = GREEN
        if kind in ("sum_rev", "sum_cost", "gop", "ebitda", "ebit"):
            for col in range(1, 6):
                ws.cell(r, col).font = BOLD
                ws.cell(r, col).border = TOP
        if kind.endswith("_m") or kind == "gop_share":
            ws.cell(r, 1).font = NOTE
        r += 1
    for col, w in {"A": 32, "B": 14, "C": 14, "D": 14, "E": 14}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B4"
    return rows


def build_checks(ws, d_first, d_last, prow, fa):
    ws["A1"] = "Step 2 · Checks before anyone sees the numbers"
    ws["A1"].font = TITLE
    ws["A2"] = "OK = passes. FLAG = needs action (a problem found and handled, or money to claim back)."
    ws["A2"].font = NOTE
    head_row(ws, 4, ["Check", "City", "Resort", "Island", "Group / result", "Status"])
    D = lambda c: f"Data!${c}${d_first}:${c}${d_last}"
    P = lambda col, label: f"'P&L'!{col}{prow[label]}"
    # 1. mapping complete + totals tie
    ws.cell(5, 1, "1a. Accounts in the export with no P&L line")
    ws.cell(5, 5, f'=COUNTIF({D("H")},"UNMAPPED")').number_format = NUM
    ws.cell(5, 6, '=IF(E5=0,"OK","FLAG")')
    ws.cell(6, 1, "1b. Export total (amounts used) − P&L total")
    ws.cell(6, 5, f"=ROUND(SUM({D('I')})-({P('E','Total revenue')}+{P('E','Total operating costs')}"
                  f"+{P('E','Management fee (as invoiced)')}+{P('E','Depreciation')}),2)").number_format = EUR
    ws.cell(6, 6, '=IF(ABS(E6)<0.5,"OK","FLAG")')
    # 2. duplicates
    ws.cell(7, 1, "2a. Duplicate lines found (removed from the totals)")
    ws.cell(7, 5, f'=COUNTIF({D("G")},"DUPLICATE")').number_format = NUM
    ws.cell(7, 6, '=IF(E7=0,"OK","FLAG")')
    ws.cell(8, 1, "2b. Amount removed as duplicate (€)")
    ws.cell(8, 5, f'=SUMIFS({D("E")},{D("G")},"DUPLICATE")').number_format = EUR
    ws.cell(9, 1, "2c. First duplicate: export line")
    ws.cell(9, 5, f'=IFERROR(INDEX({D("A")},MATCH("DUPLICATE",{D("G")},0)),"none")')
    # 3. rooms revenue tie-out with Step 1
    ws.cell(10, 1, "3. Rooms revenue: ledger − Step 1 KPIs (€)")
    for j, L in enumerate("BCD"):
        c = ws.cell(10, 2 + j, f"=ROUND({P(L,'Rooms revenue')}-KPIs!I{5 + j},2)")
        c.number_format = EUR
        c.font = GREEN
    ws.cell(10, 5, "=SUMPRODUCT(ABS(B10:D10))").number_format = EUR
    ws.cell(10, 6, '=IF(E10<0.5,"OK","FLAG")')
    # 4. management fee
    ws.cell(11, 1, "4a. Fee per contract (€)")
    ws.cell(12, 1, "4b. Fee as invoiced (€)")
    ws.cell(13, 1, "4c. Invoiced − contract: to claim back (€)")
    for j, L in enumerate("BCD"):
        col = 2 + j
        c = ws.cell(11, col, f"=Contract!$B$5*{P(L,'Total revenue')}+Contract!$B$6*{P(L,'GOP (gross operating profit)')}")
        c.number_format = EUR; c.font = GREEN
        c = ws.cell(12, col, f"={P(L,'Management fee (as invoiced)')}")
        c.number_format = EUR; c.font = GREEN
        ws.cell(13, col, f"=ROUND({L}12-{L}11,2)").number_format = EUR
    ws.cell(11, 5, "=SUM(B11:D11)").number_format = EUR
    ws.cell(12, 5, "=SUM(B12:D12)").number_format = EUR
    ws.cell(13, 5, "=SUM(B13:D13)").number_format = EUR
    ws.cell(13, 6, '=IF(SUMPRODUCT(ABS(B13:D13))<0.5,"OK","FLAG")')
    # 5. register depreciation = ledger depreciation
    ws.cell(14, 1, "5. Depreciation: fixed-asset register − ledger (€)")
    for j, L in enumerate("BCD"):
        c = ws.cell(14, 2 + j, f"=ROUND({FA}!G{fa['summary'][HOTELS[j]]}-{P(L,'Depreciation')},2)")
        c.number_format = EUR
        c.font = GREEN
    ws.cell(14, 5, "=SUMPRODUCT(ABS(B14:D14))").number_format = EUR
    ws.cell(14, 6, '=IF(E14<0.5,"OK","FLAG")')
    # 6. journal entry balances
    ws.cell(15, 1, "6. Journal entry: total debits − total credits (€)")
    c = ws.cell(15, 5, f"=ROUND({FA}!C{fa['je_total']}-{FA}!D{fa['je_total']},2)")
    c.number_format = EUR
    c.font = GREEN
    ws.cell(15, 6, '=IF(ABS(E15)<0.5,"OK","FLAG")')
    ws.cell(16, 1, "4d. EBITDA if the overcharge is recovered (€)")
    for j, L in enumerate("BCD"):
        c = ws.cell(16, 2 + j, f"={P(L,'EBITDA')}+{L}13")
        c.number_format = EUR
        c.font = GREEN
    ws.cell(16, 5, "=SUM(B16:D16)").number_format = EUR
    for r in (5, 6, 7, 10, 13, 14, 15):
        ws.cell(r, 6).font = BOLD
    ws["A18"] = ("Both problems flagged here were planted on purpose in the made-up data, to show that the checks work: "
                 "export line 40 repeats line 37 (Island electricity), and the City's operator invoiced the base fee at 4% instead of 3%.")
    ws["A15"].font = NOTE
    for col, w in {"A": 46, "B": 13, "C": 13, "D": 13, "E": 15, "F": 9}.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------- Step 3
FA = "'Fixed assets'"


def build_fixed_assets(ws, assets):
    ws["A1"] = "Step 3 · Fixed-asset register, net book value and the depreciation journal entry (made-up assets)"
    ws["A1"].font = TITLE
    ws["A2"] = ("Depreciation per year = cost ÷ useful life (land: life 0 = never depreciated). Years in use = year end − first year + 1 "
                "(in use from 1 January). Depreciation so far = per year × years, never more than cost. Net book value = cost − depreciation so far.")
    ws["A2"].font = NOTE
    ws["A3"] = "Year end"
    ws["B3"] = 2025
    ws["B3"].font = BLUE
    head_row(ws, 5, ["Hotel", "Asset", "Cost (€)", "Useful life (years)", "In use from",
                     "Years in use", "Depreciation per year (€)", "Depreciation so far (€)", "Net book value (€)"])
    first = 6
    for i, a in enumerate(assets):
        r = first + i
        ws.cell(r, 1, a["hotel"]).font = BLUE
        ws.cell(r, 2, a["asset"]).font = BLUE
        ws.cell(r, 3, float(a["cost_eur"])).font = BLUE
        ws.cell(r, 4, int(a["useful_life_years"])).font = BLUE
        ws.cell(r, 5, int(a["in_use_from_year"])).font = BLUE
        ws.cell(r, 6, f"=$B$3-E{r}+1")
        ws.cell(r, 7, f"=IF(D{r}=0,0,C{r}/D{r})")
        ws.cell(r, 8, f"=MIN(G{r}*F{r},C{r})")
        ws.cell(r, 9, f"=C{r}-H{r}")
        for c in (3, 7, 8, 9):
            ws.cell(r, c).number_format = NUM
    last = first + len(assets) - 1

    s0 = last + 3  # summary by hotel
    ws.cell(s0 - 1, 1, "By hotel").font = BOLD
    head_row(ws, s0, ["Hotel", "", "Cost (€)", "", "", "", "Depreciation 2025 (€)", "Depreciation so far (€)", "Net book value (€)"])
    rng = lambda col: f"${col}${first}:${col}${last}"
    rows = {}
    for j, h in enumerate(HOTELS + ["Group"]):
        r = s0 + 1 + j
        rows[h] = r
        ws.cell(r, 1, h).font = BOLD if h == "Group" else BLACK
        for col in "CGHI":
            if h == "Group":
                f = f"=SUM({col}{s0 + 1}:{col}{s0 + 3})"
            else:
                f = f"=SUMIFS({rng(col)},{rng('A')},$A{r})"
            c = ws[f"{col}{r}"]
            c.value = f
            c.number_format = NUM
            if h == "Group":
                c.font = BOLD
                c.border = TOP

    j0 = rows["Group"] + 3  # journal entry
    ws.cell(j0 - 1, 1, "Journal entry, 31 Dec 2025: depreciation for the year (debits must equal credits)").font = BOLD
    head_row(ws, j0, ["Hotel", "Account", "Debit (€)", "Credit (€)"])
    r = j0 + 1
    for h in HOTELS:
        ws.cell(r, 1, h)
        ws.cell(r, 2, "7000 Depreciation expense (P&L)")
        ws.cell(r, 3, f"=G{rows[h]}").number_format = NUM
        ws.cell(r + 1, 1, h)
        ws.cell(r + 1, 2, "1990 Accumulated depreciation (balance sheet)")
        ws.cell(r + 1, 4, f"=G{rows[h]}").number_format = NUM
        r += 2
    ws.cell(r, 1, "Total").font = BOLD
    ws.cell(r, 3, f"=SUM(C{j0 + 1}:C{r - 1})").number_format = NUM
    ws.cell(r, 4, f"=SUM(D{j0 + 1}:D{r - 1})").number_format = NUM
    for col in (1, 3, 4):
        ws.cell(r, col).font = BOLD
        ws.cell(r, col).border = TOP
    ws.cell(r + 2, 1, "In larger systems such as SAP, a monthly depreciation run calculates and posts this entry automatically; "
                      "in smaller ones the accountant posts it by hand. Either way the entry is the same. Account codes are made up.").font = NOTE
    for col, w in {"A": 10, "B": 44, "C": 14, "D": 14, "E": 11, "F": 10, "G": 16, "H": 16, "I": 16}.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[5].height = 30
    return {"first": first, "last": last, "summary": rows, "je_total": r}


def build_returns(ws, balances, prow, fa):
    ws["A1"] = "Step 3 · Working capital, working-capital days and return on capital (31 Dec 2025)"
    ws["A1"].font = TITLE
    ws["A2"] = ("Blue = year-end balances (made up). Green = links to P&L and Fixed assets. "
                "Supplier costs = operating costs − staff costs (staff are paid monthly, not on supplier credit).")
    ws["A2"].font = NOTE
    head_row(ws, 4, ["Measure"] + HOTELS + ["Group"])
    b = {r["hotel"]: r for r in balances}
    labels = ["Receivables (€)", "Inventory: food and drink (€)", "Payables to suppliers (€)", "Working capital (€)",
              "Revenue (€)", "Food and beverage cost (€)", "Supplier costs (€)",
              "DSO: days customers take to pay", "DIO: days stock sits", "DPO: days we take to pay", "Cash cycle (days)",
              "EBIT (€)", "Net book value (€)", "Capital employed (€)", "ROCE"]
    for i, lab in enumerate(labels):
        ws.cell(5 + i, 1, lab)
    P = lambda col, label: f"'P&L'!{col}{prow[label]}"
    for j, h in enumerate(HOTELS + ["Group"]):
        c = 2 + j
        L = "BCDE"[j]
        if h != "Group":
            for k, key in enumerate(["receivables_eur", "inventory_eur", "payables_eur"]):
                ws.cell(5 + k, c, float(b[h][key])).font = BLUE
            links = {9: P(L, "Total revenue"), 10: P(L, "Food and beverage cost"),
                     11: f"{P(L, 'Total operating costs')}-{P(L, 'Staff costs')}",
                     16: P(L, "EBIT"),
                     17: f"SUMIFS({FA}!$I${fa['summary']['City']}:$I${fa['summary']['Island']},{FA}!$A${fa['summary']['City']}:$A${fa['summary']['Island']},{L}$4)"}
            for r, f in links.items():
                ws.cell(r, c, "=" + f).font = GREEN
        else:
            for r in (5, 6, 7, 9, 10, 11, 16, 17):
                ws.cell(r, c, f"=SUM(B{r}:D{r})")
        ws.cell(8, c, f"={L}5+{L}6-{L}7")
        ws.cell(12, c, f"=IF({L}9=0,0,{L}5/{L}9*365)")
        ws.cell(13, c, f"=IF({L}10=0,0,{L}6/{L}10*365)")
        ws.cell(14, c, f"=IF({L}11=0,0,{L}7/{L}11*365)")
        ws.cell(15, c, f"={L}12+{L}13-{L}14")
        ws.cell(18, c, f"={L}17+{L}8")
        ws.cell(19, c, f"=IF({L}18=0,0,{L}16/{L}18)")
        for r in range(5, 20):
            ws.cell(r, c).number_format = PCT2 if r == 19 else ('0.0;(0.0);"-"' if 12 <= r <= 15 else EUR)
    for r in (8, 18, 19):
        for c in range(1, 6):
            ws.cell(r, c).font = BOLD
            ws.cell(r, c).border = TOP
    ws["A21"] = ("31 December is the off-season: the Resort is closed, so its receivables and stock are low. "
                 "A year-end snapshot can mislead for a seasonal business. Book values: land stays at what was paid for it.")
    ws["A21"].font = NOTE
    for col, w in {"A": 34, "B": 14, "C": 14, "D": 14, "E": 14}.items():
        ws.column_dimensions[col].width = w


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "KPIs"
    layout = build_kpis(ws, read_hotels())
    d_first, d_last = build_data(wb.create_sheet("Data"), read_csv("data/ledger_extract.csv"))
    build_mapping(wb.create_sheet("Mapping"), read_csv("data/account_mapping.csv"))
    build_contract(wb.create_sheet("Contract"), read_csv("data/management_agreement.csv"))
    prow = build_pnl(wb.create_sheet("P&L"), d_first, d_last)
    build_fee_parts(wb["Contract"], prow)
    fa = build_fixed_assets(wb.create_sheet("Fixed assets"), read_csv("data/fixed_assets.csv"))
    build_returns(wb.create_sheet("Returns"), read_csv("data/balances_2025.csv"), prow, fa)
    build_checks(wb.create_sheet("Checks"), d_first, d_last, prow, fa)
    wb.save("output/hotel_group_model.xlsx")
    print("Saved output/hotel_group_model.xlsx; fixed assets", fa)


if __name__ == "__main__":
    main()
