# Hotel Group Financial Model: FP&A and Hotel Asset Management (Excel + Python)

**In plain words:** a made-up group of three hotels in Greece. I turn its accounting export into a profit-and-loss statement (P&L) for each hotel, measure how each hotel performs, check the numbers for errors, and show how much each hotel earns on the money invested in it. A separate Python script re-checks every number in the Excel workbook.

**Keywords:** FP&A · financial analysis · P&L · EBITDA · KPIs (occupancy, ADR, RevPAR) · working capital (DSO, DIO, DPO) · return on capital (ROCE) · depreciation · journal entry · data checks · Excel · Python

- **The hotels and their numbers are made up.** The economic data used in later steps (growth, inflation, interest rates) is real, and each figure names its source.
- **Status: version 0.1. Steps 1–3 of 10 are done.** Each new step adds one small piece to the workbook.

**Planned next, one step at a time:**

| Step | What it adds |
|---|---|
| 4 | Due diligence: adjusted EBITDA (quality of earnings), net debt and working capital, and what they mean for the price of a hotel |
| 5 | The 2026 budget and forecast, from real European Commission data |
| 6 | A renovation: NPV and IRR |
| 7 | Valuing a hotel: a five-year DCF, with and without a loan |
| 8 | Buying one hotel and selling another |
| 9 | A one-page report with recommendations |
| 10 | Company treasury: a cash plan, a bank reconciliation, and interest-rate and FX hedges |

| Hotel | Where | Rooms | Days open |
|---|---|---|---|
| City | Thessaloniki, open all year | 100 | 365 |
| Resort | Halkidiki, summer season | 300 | 150 |
| Island | Crete, long season | 100 | 200 |

---

## Step 1 · Hotel KPIs (sheet `KPIs`)

A hotel sells **room-nights**: one room for one night.

| KPI | Formula | What it tells you |
|---|---|---|
| Available room-nights | rooms × days open | what the hotel could sell |
| Occupancy | sold room-nights ÷ available room-nights | how full it was |
| ADR (average daily rate) | rooms revenue ÷ sold room-nights | the average price actually paid per night sold |
| RevPAR (revenue per available room) | rooms revenue ÷ available room-nights = occupancy × ADR | how full **and** how well priced, in one number |

**Results, 2025**

| | Occupancy | ADR | RevPAR | Rooms revenue |
|---|---|---|---|---|
| City | 70.0% | €100 | €70 | €2,555,000 |
| Resort | 80.0% | €300 | €240 | €10,800,000 |
| Island | 60.0% | €200 | €120 | €2,400,000 |
| **Group** | **72.5%** | **€214.21** | **€155.22** | **€15,755,000** |

**Group KPIs are added up first, then divided.** Averaging the three hotels' percentages would give 70.0% occupancy and €200 ADR. That is wrong, because it treats the 300-room resort the same as a 100-room hotel. The sheet shows both, so the difference is visible.

**What it says:** the summer resort earns 69% of the group's rooms revenue. RevPAR counts only the nights a hotel is open, so a seasonal resort and a year-round city hotel are not directly comparable on RevPAR alone.

---

## Step 2 · From an accounting export to a hotel P&L (sheets `Data`, `Mapping`, `Contract`, `P&L`, `Checks`)

1. **The export** (`data/ledger_extract.csv`): each account's 2025 total for each hotel. 15 accounts × 3 hotels = 45 lines expected; it has 46.
2. **Clean before adding:** each hotel and account may appear once. Line 40 repeats line 37 (Island electricity, €250,000), so it is flagged and left out of the totals.
3. **Mapping** (`data/account_mapping.csv`): every account belongs to one P&L line, for example 6000 Salaries + 6010 Social security → Staff costs. Each P&L cell is one `SUMIFS`.
4. **The hotel P&L:** total revenue − operating costs = **GOP** (gross operating profit: what the hotel's daily operations earn) − management fee = **EBITDA** − depreciation = **EBIT**.
5. **The management fee check:** the made-up contract (`data/management_agreement.csv`) pays the operator 3% of total revenue + 10% of GOP. Each hotel's invoiced fee is compared with that.

**Results, 2025 (€)**

| | City | Resort | Island | Group |
|---|---|---|---|---|
| Total revenue | 3,200,000 | 15,000,000 | 3,000,000 | 21,200,000 |
| GOP | 1,120,000 | 6,300,000 | 900,000 | 8,320,000 |
| GOP margin | 35.0% | 42.0% | 30.0% | 39.2% |
| Management fee (as invoiced) | 240,000 | 1,080,000 | 180,000 | 1,500,000 |
| EBITDA | 880,000 | 5,220,000 | 720,000 | 6,820,000 |
| EBIT | 580,000 | 3,720,000 | 320,000 | 4,620,000 |

**What it says:** the resort makes 75.7% of the group's GOP. The City's operator charged the base fee at 4% instead of 3%: **€32,000 to claim back**. The P&L shows what was actually charged, because that is what the books hold; after recovery the City's EBITDA would be €912,000.

**The four checks** (sheet `Checks`):

| Check | Result |
|---|---|
| 1. Every account maps to a P&L line, and the export total equals the P&L total | OK |
| 2. One line per hotel and account | 1 duplicate found (line 40, €250,000), removed |
| 3. Rooms revenue in the ledger = rooms revenue from Step 1 (two sources agree) | OK, €0 difference |
| 4. Management fee invoiced = fee per contract | City: €32,000 to claim back |

**Two problems were planted on purpose** in the made-up data to show that the checks catch them: the duplicate line 40 and the City's fee invoice.

### Data specification: `data/ledger_extract.csv`
| Column | Meaning | Rule |
|---|---|---|
| `line` | line number in the export | unique, 1, 2, 3, … |
| `hotel` | City, Resort or Island | must be one of the three |
| `account` | 4-digit account code | must exist in `account_mapping.csv` |
| `account_name` | the account's name | for reading only |
| `amount_eur` | the account's 2025 total, in euros | positive; revenue accounts start with 4, cost accounts with 5, 6 or 7 |

Each hotel and account appears **once**. A second line for the same hotel and account is treated as a duplicate and left out of the totals.

---

## Step 3 · Depreciation, working capital and return on capital (sheets `Fixed assets`, `Returns`)

**Depreciation** spreads an asset's cost over its useful life: cost ÷ life each year. Land is never depreciated. **Net book value** = cost − depreciation so far.

| 31 Dec 2025 (€) | Cost | Depreciation 2025 | Depreciation so far | Net book value |
|---|---|---|---|---|
| City | 11,000,000 | 300,000 | 2,500,000 | 8,500,000 |
| Resort | 55,000,000 | 1,500,000 | 17,500,000 | 37,500,000 |
| Island | 14,500,000 | 400,000 | 2,000,000 | 12,500,000 |
| **Group** | **80,500,000** | **2,200,000** | **22,000,000** | **58,500,000** |

Each hotel has land, a building (40-year life) and furniture and equipment (10-year life); the register (`data/fixed_assets.csv`) lists all nine. Its depreciation equals the ledger's in Step 2 (check 5). The year's **journal entry** debits 7000 Depreciation expense and credits 1990 Accumulated depreciation, €2,200,000 on each side (check 6: debits = credits).

**Working capital** = receivables + inventory − payables, from the year-end balances (`data/balances_2025.csv`). In days: DSO = receivables ÷ revenue × 365; DIO = inventory ÷ food and beverage cost × 365; DPO = payables ÷ supplier costs × 365 (supplier costs = operating costs − staff costs).

**ROCE** = EBIT ÷ capital employed, where capital employed = net book value + working capital.

| 2025 | Working capital | DSO | DIO | DPO | Cash cycle | ROCE |
|---|---|---|---|---|---|---|
| City | €95,000 | 29.7 | 27.4 | 60.8 | −3.8 | 6.75% |
| Resort | −€180,000 | 4.9 | 5.6 | 32.4 | −22.0 | 9.97% |
| Island | −€20,000 | 14.6 | 18.3 | 49.8 | −16.9 | 2.56% |
| **Group** | **−€105,000** | | | | | **7.91%** |

**What it says:** for every €100 tied up in it, the Resort earns €9.97 a year and the Island €2.56. The Island is new and the least full: the question an asset manager asks is whether it earns enough on what was invested.

---

## How every number is checked (`check_excel.py`)
| Check | What it does |
|---|---|
| 1. Inputs | every blue input cell equals the CSV file it came from |
| 2. Formulas | every number that is not an input is a formula (no typed-in results), and every row or hotel column of a table uses the same formula as the first one |
| 3. Results | every result in the workbook equals `model.py`, which does the same maths again in plain Python |
| 4. The workbook's own checks | the `Checks` sheet shows exactly what `model.py` expects |
| 5. Change test | a copy gets four inputs changed (Resort occupancy, base fee, a building's life, Island receivables); LibreOffice recalculates it, and Excel and Python must reach the same new numbers |

**Result: 494 of 494 pass** (`output/check_log.txt`). As a test of the checker itself, a copy with a wrong RevPAR formula and one typed-in number was run through it; both were caught, including the typed number that happened to equal the right answer.

## How to run
```
pip install -r requirements.txt
python build_excel.py      # writes output/hotel_group_model.xlsx (formulas only)
# open the workbook in Excel and save it once (or recalculate it with LibreOffice)
python check_excel.py      # the five checks; writes output/check_log.txt (the change test needs LibreOffice)
```

## How it was made
I designed the analysis and made the methodological decisions, and used AI assistance to implement the code. Every number is checked twice: by the workbook's own `Checks` sheet and by `check_excel.py`.

## Files
| File | What it is |
|---|---|
| `data/hotels.csv` | the made-up inputs: rooms, days open, occupancy, ADR |
| `data/ledger_extract.csv` | the made-up accounting export, 46 lines (one planted duplicate) |
| `data/account_mapping.csv` | account → P&L line |
| `data/management_agreement.csv` | the made-up fee terms, with sources for the typical ranges |
| `data/fixed_assets.csv` | the made-up fixed-asset register (9 assets) |
| `data/balances_2025.csv` | the made-up year-end receivables, stock and payables |
| `build_excel.py` | builds the workbook; every result is an Excel formula, inputs in blue |
| `model.py` | the same maths in plain Python |
| `check_excel.py` | the five checks above |
| `output/hotel_group_model.xlsx` | the workbook |
| `output/check_log.txt` | the result of the last check |

## Honest limits (so far)
- The hotels are made up; their round numbers sit in a realistic range for Greece.
- One year only.
- The export holds yearly totals per account, not individual postings.
- The P&L stops at EBIT. Real hotel reports also deduct owner costs such as insurance and property tax before EBITDA; this model leaves them out.
- The incentive fee is calculated on GOP. Many real contracts use GOP after the base fee (adjusted GOP).
- The fee terms are made up but sit in the typical range reported by HVS: base fees of 2–4% of total revenue (3% most common) and an incentive fee of about 10% of GOP.
- Working capital uses year-end balances. 31 December is the off-season for the resorts, so the days can look very different from August. Guests' advance deposits are not modelled.
- Net book value and ROCE use book values: land stays at what was paid for it.
