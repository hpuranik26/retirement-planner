# Retirement Withdrawal Planner

A 35-year retirement income and tax projection tool that produces a fully formatted Excel workbook from user-supplied inputs. Built for high-net-worth couples retiring with a mix of tax-deferred, taxable, and tax-free accounts.

## What it does

- Projects year-by-year portfolio balances, withdrawals, taxes, and net spendable income across a configurable retirement horizon
- Optimizes withdrawal order: **Tax-Deferred → Taxable (LTCG) → Roth**
- Fills a configurable federal bracket (22% or 24%) with Roth conversions each year to minimize lifetime taxes
- Applies **2026 tax law** (IRS Rev. Proc. 2025-32, OBBBA, CMS, Oregon HB 3753) with a versioned tax law database — add future years as IRS publishes them
- Tracks Social Security claiming strategies with a 3-scenario comparison (Early / FRA / Delay)
- Models ACA health insurance (pre-65) and Medicare (post-65) costs with per-spouse and per-dependent tracking
- Calculates the ACA subsidy cliff (400% FPL) with dynamic household size as dependents age off the plan at 26
- Flags IRMAA surcharges only for Medicare-eligible spouses
- Projects RMDs (SECURE 2.0: age 75 for born 1960+) and flags shortfalls

## Output — 8 Excel sheets

| Sheet | Contents |
|---|---|
| **Inputs** | All amber input cells — edit here, never overwritten by the script |
| **Tax Tables** | Exact bracket values, IRMAA tiers, LTCG thresholds used in this run |
| **Assumptions** | All parameters used + 35-year summary totals |
| **Annual Summary** | Full year-by-year table (36 columns across 7 groups) |
| **SS Scenarios** | Lifetime SS comparison: Early / FRA / Delay Primary to 70 |
| **Insurance Detail** | Per-spouse ACA and Medicare costs by year, children on plan |
| **RMD Projections** | Required Minimum Distributions from age 75, QCD notes |
| **Strategy Summary** | Plain-language playbook and quarterly checklist |

## Requirements

```bash
pip3 install openpyxl
```

Python 3.8+ required. No other dependencies.

## Usage

**First run — creates the input template:**
```bash
python3 build_retirement_planner.py MyScenario.xlsx
```
Opens `MyScenario.xlsx` with an Inputs tab pre-filled with defaults. Edit the amber cells and save.

**Subsequent runs — reads inputs, writes all output sheets:**
```bash
python3 build_retirement_planner.py MyScenario.xlsx
```

The **Inputs tab is never overwritten**. Run the script as many times as needed after editing inputs.

## Key inputs (all editable in the Inputs tab)

| Section | Inputs |
|---|---|
| Family | Spouse ages, birth years, up to 3 child birth years (enter 0 if none) |
| Portfolio | Tax-deferred, taxable, and Roth balances |
| Strategy | Gross withdrawal rate, Roth conversion bracket, retirement start year, length |
| Tax Law | Tax law year — uses that year's IRS/CMS tables; falls back to latest available |
| Social Security | FRA monthly benefit and claiming age for each spouse |
| Returns | Growth rate, price return, dividend yield, inflation rate, SS COLA |
| Health Insurance | ACA base rate, ACA inflation, Medicare base premium, Medicare inflation |

## Tax law database

Tax tables live in `TAX_LAW_DB` keyed by year. To add a new year when IRS publishes (typically each October):

```python
TAX_LAW_DB = {
    2025: { ... },
    2026: { ... },
    2027: {           # ← add this block each October
        'fed_brackets':   [...],
        'std_ded_mfj':    33_000,
        'senior_bonus':    6_150,
        'or_brackets':    [...],
        'or_std_ded':     9_920,
        'ltcg_brackets':  [...],
        'niit_threshold': 250_000,
        'irmaa':          [...],
        'qcd_limit':      110_000,
        'rmd_start_age':       75,
    }
}
```

If you enter a future tax law year not yet in the database, the script falls back to the latest available year and flags it clearly in the output title and Strategy Summary tab.

## Tax law sources (2026)

| Item | Source |
|---|---|
| Federal brackets, standard deduction, LTCG thresholds | IRS Rev. Proc. 2025-32 |
| Senior bonus deduction ($6,000/spouse at 65+) | One Big Beautiful Bill Act (OBBBA) |
| Oregon brackets and standard deduction ($9,680 MFJ) | Oregon HB 3753 + OR DOR 2026 |
| IRMAA tiers and Part B/D surcharges | CMS 2026 announcement |
| ACA subsidy cliff (400% FPL, 2-person: $86,560) | HHS 2026 poverty guidelines |
| RMD start age (75, born 1960+) | SECURE 2.0 Act |
