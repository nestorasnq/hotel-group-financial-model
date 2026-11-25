"""Checks the Excel workbook in five ways and writes output/check_log.txt.

1. Inputs   - every blue input cell equals the CSV file it came from.
2. Formulas - every result is a formula (no typed-in numbers), and each repeated row
              uses the same formula as the first row of its table.
3. Results  - every Excel result equals model.py, which does the same maths in Python.
4. Checks   - the workbook's own checks (sheet Checks) show what model.py expects.
5. Change   - change a few inputs in a copy, let LibreOffice recalculate it, and confirm
              Excel and Python move to the same new numbers. Needs LibreOffice ("soffice").

Run after the workbook has been recalculated (open and save it in Excel, or LibreOffice):
    python check_excel.py
"""
import os
import shutil
import subprocess
import tempfile
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

import model

XLSX = "output/hotel_group_model.xlsx"
TOL = 0.005
log, fails = [], 0


def rec(ok, text):
    global fails
    fails += not ok
    log.append(("OK    " if ok else "FAIL  ") + text)


def close(a, b):
    if isinstance(b, str) or isinstance(a, str):
        return a == b
    return a is not None and abs(a - b) <= TOL


def find_row(ws, label, col=1, start=1, end=200):
    for r in range(start, end):
        if ws.cell(r, col).value == label:
            return r
    raise KeyError(f"{ws.title}: no row '{label}'")


# ---------------------------------------------------------------- 1. inputs
def check_inputs(wbf, inp):
    k = wbf["KPIs"]
    for i, h in enumerate(inp["hotels"]):
        r = 5 + i
        for col, key, conv in ((3, "rooms", int), (4, "days_open", int), (5, "occupancy", float), (6, "adr_eur", float)):
            rec(k.cell(r, col).value == conv(h[key]), f"input KPIs!{k.cell(r, col).coordinate} = hotels.csv {h['hotel']} {key}")
    d = wbf["Data"]
    bad = 0
    for i, row in enumerate(inp["ledger"]):
        r = 5 + i
        vals = [d.cell(r, c).value for c in range(1, 6)]
        want = [int(row["line"]), row["hotel"], int(row["account"]), row["account_name"], float(row["amount_eur"])]
        bad += vals != want
    rec(bad == 0, f"input Data!A5:E{4 + len(inp['ledger'])} = ledger_extract.csv ({len(inp['ledger'])} lines, {bad} different)")
    m = wbf["Mapping"]
    bad = sum([m.cell(5 + i, 1).value, m.cell(5 + i, 3).value] != [int(x["account"]), x["pnl_line"]]
              for i, x in enumerate(inp["mapping"]))
    rec(bad == 0, f"input Mapping = account_mapping.csv ({len(inp['mapping'])} accounts, {bad} different)")
    c = wbf["Contract"]
    terms = {t["term"]: float(t["value"]) for t in inp["contract"]}
    rec(c["B5"].value == terms["base_fee_pct_of_revenue"], "input Contract!B5 = base fee")
    rec(c["B6"].value == terms["incentive_fee_pct_of_gop"], "input Contract!B6 = incentive fee")
    fa = wbf["Fixed assets"]
    rec(fa["B3"].value == inp["year_end"], "input Fixed assets!B3 = year end")
    bad = 0
    for i, a in enumerate(inp["assets"]):
        r = 6 + i
        vals = [fa.cell(r, col).value for col in range(1, 6)]
        bad += vals != [a["hotel"], a["asset"], float(a["cost_eur"]), int(a["useful_life_years"]), int(a["in_use_from_year"])]
    rec(bad == 0, f"input Fixed assets!A6:E{5 + len(inp['assets'])} = fixed_assets.csv ({bad} different)")
    rt = wbf["Returns"]
    bal = {b["hotel"]: b for b in inp["balances"]}
    bad = 0
    for j, h in enumerate(model.HOTELS):
        for k_, key in enumerate(["receivables_eur", "inventory_eur", "payables_eur"]):
            bad += rt.cell(5 + k_, 2 + j).value != float(bal[h][key])
    rec(bad == 0, f"input Returns!B5:D7 = balances_2025.csv ({bad} different)")


# ---------------------------------------------------------------- 2. formulas
def is_input(cell):
    return cell.font is not None and cell.font.color is not None and cell.font.color.rgb in ("000000FF", "FF0000FF")


def check_formulas(wbf):
    typed = []
    for ws in wbf.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, (int, float)) and not isinstance(v, bool) and not is_input(cell):
                    typed.append(f"{ws.title}!{cell.coordinate}")
    rec(not typed, "every number that is not a blue input is a formula" + (f" — typed numbers at {typed[:10]}" if typed else ""))
    # repeated rows: each row's formula must be the first row's formula, moved down
    tables = [("KPIs", 5, 7, "GHIJ"), ("Data", 5, 50, "FGHI"), ("Fixed assets", 6, 14, "FGHI"), ("Contract", 12, 14, "BCDEF")]
    for sheet, r1, r2, cols in tables:
        ws, diff = wbf[sheet], []
        for col in cols:
            first = ws[f"{col}{r1}"].value
            for r in range(r1 + 1, r2 + 1):
                expect = Translator(first, origin=f"{col}{r1}").translate_formula(f"{col}{r}")
                if ws[f"{col}{r}"].value != expect:
                    diff.append(f"{col}{r}")
        rec(not diff, f"{sheet}: columns {cols}, rows {r1}–{r2} all use the same formula" + (f" — differ at {diff}" if diff else ""))
    # hotel columns in P&L and Returns: B, C, D must be the same formula moved across
    for sheet, r1, r2 in (("P&L", 4, 21), ("Returns", 8, 19)):
        ws, diff = wbf[sheet], []
        for r in range(r1, r2 + 1):
            b = ws[f"B{r}"].value
            if not (isinstance(b, str) and b.startswith("=")):
                continue
            for col in "CD":
                if ws[f"{col}{r}"].value != Translator(b, origin=f"B{r}").translate_formula(f"{col}{r}"):
                    diff.append(f"{col}{r}")
        rec(not diff, f"{sheet}: the three hotel columns use the same formulas" + (f" — differ at {diff}" if diff else ""))


# ---------------------------------------------------------------- 3 + 4. results and checks
def compare_results(wbv, res, tag=""):
    n0 = len(log)
    k = wbv["KPIs"]
    for name, vals in res["kpis"].items():
        r = find_row(k, name, start=5, end=9)
        cols = {"Available": 7, "Sold": 8, "Rooms revenue": 9, "RevPAR": 10, "Occupancy": 5, "ADR": 6}
        for key, py in vals.items():
            xl = k.cell(r, cols[key]).value
            rec(close(xl, py), f"{tag}KPIs {name} {key}: excel {xl} python {py:.4f}")
    p = wbv["P&L"]
    for r in range(4, 22):
        label = p.cell(r, 1).value
        for j, h in enumerate(model.HOTELS + ["Group"]):
            xl, py = p.cell(r, 2 + j).value, res["pnl"][h][label]
            rec(close(xl, py), f"{tag}P&L {label} {h}: excel {xl} python {py:.4f}")
    fa = wbv["Fixed assets"]
    for i, a in enumerate(res["register"]):
        r = 6 + i
        for col, key in ((6, "Years in use"), (7, "Depreciation per year"), (8, "Depreciation so far"), (9, "Net book value")):
            rec(close(fa.cell(r, col).value, a[key]), f"{tag}Fixed assets {a['hotel']} {a['asset']} {key}: excel {fa.cell(r, col).value} python {a[key]}")
    for h in model.HOTELS + ["Group"]:
        r = find_row(fa, h, start=16, end=23)
        for col, key in ((3, "Cost"), (7, "Depreciation 2025"), (8, "Depreciation so far"), (9, "Net book value")):
            rec(close(fa.cell(r, col).value, res["by_hotel"][h][key]), f"{tag}Fixed assets total {h} {key}: excel {fa.cell(r, col).value}")
    rj = find_row(fa, "Total", start=24, end=40)
    rec(close(fa.cell(rj, 3).value, res["je"]["debits"]), f"{tag}Journal entry total debits: excel {fa.cell(rj, 3).value}")
    rec(close(fa.cell(rj, 4).value, res["je"]["credits"]), f"{tag}Journal entry total credits: excel {fa.cell(rj, 4).value}")
    rt = wbv["Returns"]
    for r in range(5, 20):
        label = rt.cell(r, 1).value
        for j, h in enumerate(model.HOTELS + ["Group"]):
            xl, py = rt.cell(r, 2 + j).value, res["returns"][h][label]
            rec(close(xl, py), f"{tag}Returns {label} {h}: excel {xl} python {py:.4f}")
    ck, c = wbv["Checks"], res["checks"]
    rec(close(ck["E5"].value, c["1a"]), f"{tag}check 1a unmapped accounts: {ck['E5'].value}")
    rec(close(ck["E6"].value, c["1b"]), f"{tag}check 1b export total − P&L total: {ck['E6'].value}")
    rec(close(ck["E7"].value, c["2a"]), f"{tag}check 2a duplicates: {ck['E7'].value}")
    rec(close(ck["E8"].value, c["2b"]), f"{tag}check 2b amount removed: {ck['E8'].value}")
    rec(close(ck["E9"].value, c["2c"]), f"{tag}check 2c duplicate line: {ck['E9'].value}")
    for j, h in enumerate(model.HOTELS):
        rec(close(ck.cell(10, 2 + j).value, c["3"][h]), f"{tag}check 3 rooms tie-out {h}: {ck.cell(10, 2 + j).value}")
        rec(close(ck.cell(11, 2 + j).value, c["4a"][h]), f"{tag}check 4a fee per contract {h}: {ck.cell(11, 2 + j).value}")
        rec(close(ck.cell(13, 2 + j).value, c["4c"][h]), f"{tag}check 4c to claim back {h}: {ck.cell(13, 2 + j).value}")
        rec(close(ck.cell(14, 2 + j).value, c["5"][h]), f"{tag}check 5 register − ledger depreciation {h}: {ck.cell(14, 2 + j).value}")
    rec(close(ck["E15"].value, c["6"]), f"{tag}check 6 journal entry debits − credits: {ck['E15'].value}")
    for j, h in enumerate(model.HOTELS):
        rec(close(ck.cell(16, 2 + j).value, c["4d"][h]), f"{tag}check 4d EBITDA after recovery {h}: {ck.cell(16, 2 + j).value}")
    ct = wbv["Contract"]
    for i, h in enumerate(model.HOTELS):
        for col, key in ((4, "Base fee"), (5, "Incentive fee"), (6, "Fee per contract")):
            rec(close(ct.cell(12 + i, col).value, c["fee_parts"][h][key]), f"{tag}Contract {h} {key}: excel {ct.cell(12 + i, col).value}")
    return len(log) - n0


# ---------------------------------------------------------------- 5. change test
CHANGES = [  # (sheet, cell, new value) and the same change for Python: (table, match, field, value)
    (("KPIs", "E6", 0.70), ("hotels", {"hotel": "Resort"}, "occupancy", 0.70)),
    (("Contract", "B5", 0.04), ("contract", {"term": "base_fee_pct_of_revenue"}, "value", 0.04)),
    (("Fixed assets", "D7", 50), ("assets", {"hotel": "City", "asset": "Building"}, "useful_life_years", 50)),
    (("Returns", "D5", 240000.0), ("balances", {"hotel": "Island"}, "receivables_eur", 240000)),
]


def change_test(inp):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        log.append("SKIP  change test: LibreOffice (soffice) not found")
        return
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "changed.xlsx")
    wb = load_workbook(XLSX)
    for (sheet, cell, value), _ in CHANGES:
        wb[sheet][cell].value = value
    wb.save(src)
    out = os.path.join(tmp, "out")
    subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out, src],
                   capture_output=True, timeout=180)
    wbv = load_workbook(os.path.join(out, "changed.xlsx"), data_only=True)
    res = model.run_all(model.with_changes(inp, [c for _, c in CHANGES]))
    log.append("      changed: Resort occupancy 80%→70%; base fee 3%→4%; City building life 40→50 years; Island receivables 120k→240k")
    n = compare_results(wbv, res, tag="[changed] ")
    log.append(f"      {n} results compared after the change")
    shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    inp = model.load_inputs()
    wbf = load_workbook(XLSX)                 # formulas
    wbv = load_workbook(XLSX, data_only=True)  # values
    sections = [("1. INPUTS", lambda: check_inputs(wbf, inp)),
                ("2. FORMULAS", lambda: check_formulas(wbf)),
                ("3–4. RESULTS AND THE WORKBOOK'S OWN CHECKS", lambda: compare_results(wbv, model.run_all(inp))),
                ("5. CHANGE TEST", lambda: change_test(inp))]
    for title, fn in sections:
        log.append(f"\n{title}")
        fn()
    total = sum(1 for line in log if line.startswith(("OK", "FAIL")))
    log.append(f"\n{total - fails} of {total} checks passed; {fails} failed.")
    os.makedirs("output", exist_ok=True)
    with open("output/check_log.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(log).strip() + "\n")
    print("\n".join(line for line in log if not line.startswith("OK")))
    raise SystemExit(1 if fails else 0)
