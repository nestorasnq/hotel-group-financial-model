"""The same maths as the Excel workbook, written again in plain Python.

check_excel.py compares every Excel result with these functions, so a wrong
formula in Excel (or a wrong line here) shows up as a mismatch.
Step 1: hotel KPIs. Step 2: the hotel P&L and its checks.
Step 3: fixed assets, the depreciation journal entry, working capital and ROCE."""
import csv
import copy

HOTELS = ["City", "Resort", "Island"]
COST_LINES = ["Staff costs", "Food and beverage cost", "Utilities and energy", "Sales and marketing",
              "Repairs and maintenance", "Other operating costs"]


def read_csv(path):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def load_inputs(folder="data"):
    """Every input the workbook uses, read from the CSV files."""
    return {
        "hotels": read_csv(f"{folder}/hotels.csv"),
        "ledger": read_csv(f"{folder}/ledger_extract.csv"),
        "mapping": read_csv(f"{folder}/account_mapping.csv"),
        "contract": read_csv(f"{folder}/management_agreement.csv"),
        "assets": read_csv(f"{folder}/fixed_assets.csv"),
        "balances": read_csv(f"{folder}/balances_2025.csv"),
        "year_end": 2025,
    }


# ---------------------------------------------------------------- Step 1
def step1_kpis(inp):
    out = {}
    for h in inp["hotels"]:
        available = int(h["rooms"]) * int(h["days_open"])
        sold = available * float(h["occupancy"])
        revenue = sold * float(h["adr_eur"])
        out[h["hotel"]] = {"Available": available, "Sold": sold, "Rooms revenue": revenue,
                           "RevPAR": revenue / available}
    a = sum(v["Available"] for v in out.values())
    s = sum(v["Sold"] for v in out.values())
    r = sum(v["Rooms revenue"] for v in out.values())
    out["Group"] = {"Available": a, "Sold": s, "Rooms revenue": r,
                    "Occupancy": s / a, "ADR": r / s, "RevPAR": r / a}
    return out


# ---------------------------------------------------------------- Step 2
def step2_pnl(inp, kpis):
    mapping = {m["account"]: m["pnl_line"] for m in inp["mapping"]}
    terms = {t["term"]: float(t["value"]) for t in inp["contract"]}
    seen, dups, unmapped, totals = set(), [], [], {}
    for r in inp["ledger"]:
        key = (r["hotel"], r["account"])
        if key in seen:                      # one line per hotel and account
            dups.append(r)
            continue
        seen.add(key)
        line = mapping.get(r["account"])
        if line is None:
            unmapped.append(r)
            continue
        totals[(r["hotel"], line)] = totals.get((r["hotel"], line), 0.0) + float(r["amount_eur"])
    pnl = {}
    for h in HOTELS:
        g = lambda line: totals.get((h, line), 0.0)
        rev = g("Rooms revenue") + g("F&B and other revenue")
        opex = sum(g(c) for c in COST_LINES)
        gop = rev - opex
        fee = g("Management fee")
        ebitda = gop - fee
        pnl[h] = {"Rooms revenue": g("Rooms revenue"), "F&B and other revenue": g("F&B and other revenue"),
                  "Total revenue": rev, **{c: g(c) for c in COST_LINES}, "Total operating costs": opex,
                  "GOP (gross operating profit)": gop, "GOP margin": gop / rev if rev else 0,
                  "Management fee (as invoiced)": fee, "EBITDA": ebitda, "EBITDA margin": ebitda / rev if rev else 0,
                  "Depreciation": g("Depreciation"), "EBIT": ebitda - g("Depreciation")}
    grp = {k: sum(pnl[h][k] for h in HOTELS) for k in pnl["City"]}
    grp["GOP margin"] = grp["GOP (gross operating profit)"] / grp["Total revenue"]
    grp["EBITDA margin"] = grp["EBITDA"] / grp["Total revenue"]
    pnl["Group"] = grp
    for h in HOTELS + ["Group"]:
        pnl[h]["Share of group GOP"] = pnl[h]["GOP (gross operating profit)"] / grp["GOP (gross operating profit)"]
    fee_contract = {h: terms["base_fee_pct_of_revenue"] * pnl[h]["Total revenue"]
                    + terms["incentive_fee_pct_of_gop"] * pnl[h]["GOP (gross operating profit)"] for h in HOTELS}
    checks = {
        "1a": len(unmapped),
        "1b": 0.0 if not unmapped else -sum(float(u["amount_eur"]) for u in unmapped),
        "2a": len(dups),
        "2b": sum(float(d["amount_eur"]) for d in dups),
        "2c": int(dups[0]["line"]) if dups else "none",
        "3": {h: pnl[h]["Rooms revenue"] - kpis[h]["Rooms revenue"] for h in HOTELS},
        "4a": fee_contract,
        "4b": {h: pnl[h]["Management fee (as invoiced)"] for h in HOTELS},
        "4c": {h: pnl[h]["Management fee (as invoiced)"] - fee_contract[h] for h in HOTELS},
        "4d": {h: pnl[h]["EBITDA"] + pnl[h]["Management fee (as invoiced)"] - fee_contract[h] for h in HOTELS},
        "fee_parts": {h: {"Base fee": terms["base_fee_pct_of_revenue"] * pnl[h]["Total revenue"],
                          "Incentive fee": terms["incentive_fee_pct_of_gop"] * pnl[h]["GOP (gross operating profit)"],
                          "Fee per contract": fee_contract[h]} for h in HOTELS},
    }
    return pnl, checks


# ---------------------------------------------------------------- Step 3
def step3_assets(inp):
    year = inp["year_end"]
    register, by_hotel = [], {h: {"Cost": 0.0, "Depreciation 2025": 0.0, "Depreciation so far": 0.0,
                                  "Net book value": 0.0} for h in HOTELS}
    for a in inp["assets"]:
        cost, life, start = float(a["cost_eur"]), int(a["useful_life_years"]), int(a["in_use_from_year"])
        years = year - start + 1
        per_year = cost / life if life else 0.0
        so_far = min(per_year * years, cost)
        nbv = cost - so_far
        register.append({"hotel": a["hotel"], "asset": a["asset"], "Years in use": years,
                         "Depreciation per year": per_year, "Depreciation so far": so_far, "Net book value": nbv})
        b = by_hotel[a["hotel"]]
        b["Cost"] += cost; b["Depreciation 2025"] += per_year; b["Depreciation so far"] += so_far; b["Net book value"] += nbv
    by_hotel["Group"] = {k: sum(by_hotel[h][k] for h in HOTELS) for k in by_hotel["City"]}
    je = {"debits": by_hotel["Group"]["Depreciation 2025"], "credits": by_hotel["Group"]["Depreciation 2025"]}
    return register, by_hotel, je


def step3_returns(inp, pnl, by_hotel):
    bal = {b["hotel"]: b for b in inp["balances"]}
    out = {}
    for h in HOTELS + ["Group"]:
        if h == "Group":
            ar = sum(float(bal[x]["receivables_eur"]) for x in HOTELS)
            inv = sum(float(bal[x]["inventory_eur"]) for x in HOTELS)
            ap = sum(float(bal[x]["payables_eur"]) for x in HOTELS)
        else:
            ar, inv, ap = (float(bal[h]["receivables_eur"]), float(bal[h]["inventory_eur"]), float(bal[h]["payables_eur"]))
        p = pnl[h]
        rev, fb, supp = p["Total revenue"], p["Food and beverage cost"], p["Total operating costs"] - p["Staff costs"]
        wc = ar + inv - ap
        dso, dio, dpo = ar / rev * 365, inv / fb * 365, ap / supp * 365
        ce = by_hotel[h]["Net book value"] + wc
        out[h] = {"Receivables (€)": ar, "Inventory: food and drink (€)": inv, "Payables to suppliers (€)": ap,
                  "Working capital (€)": wc, "Revenue (€)": rev, "Food and beverage cost (€)": fb,
                  "Supplier costs (€)": supp, "DSO: days customers take to pay": dso,
                  "DIO: days stock sits": dio, "DPO: days we take to pay": dpo, "Cash cycle (days)": dso + dio - dpo,
                  "EBIT (€)": p["EBIT"], "Net book value (€)": by_hotel[h]["Net book value"],
                  "Capital employed (€)": ce, "ROCE": p["EBIT"] / ce}
    return out


def run_all(inp):
    """Everything the workbook calculates, from the inputs alone."""
    kpis = step1_kpis(inp)
    pnl, checks = step2_pnl(inp, kpis)
    register, by_hotel, je = step3_assets(inp)
    returns = step3_returns(inp, pnl, by_hotel)
    checks["5"] = {h: by_hotel[h]["Depreciation 2025"] - pnl[h]["Depreciation"] for h in HOTELS}
    checks["6"] = je["debits"] - je["credits"]
    return {"kpis": kpis, "pnl": pnl, "checks": checks, "register": register, "by_hotel": by_hotel,
            "je": je, "returns": returns}


def with_changes(inp, changes):
    """A copy of the inputs with some values changed (used by the change test)."""
    new = copy.deepcopy(inp)
    for table, match, field, value in changes:
        for row in new[table]:
            if all(row[k] == v for k, v in match.items()):
                row[field] = str(value)
    return new


if __name__ == "__main__":
    res = run_all(load_inputs())
    g, r = res["pnl"]["Group"], res["returns"]["Group"]
    print(f"Group 2025: rooms revenue €{res['kpis']['Group']['Rooms revenue']:,.0f}; GOP €{g['GOP (gross operating profit)']:,.0f} "
          f"({g['GOP margin']:.1%}); EBITDA €{g['EBITDA']:,.0f}; EBIT €{g['EBIT']:,.0f}; ROCE {r['ROCE']:.2%}")
