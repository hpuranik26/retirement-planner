# Changelog

All notable changes to the Retirement Withdrawal Planner are documented here.

---

## [Unreleased / Session 3] — 2026-06-11

### Added
- **Charts sheet** (new tab, positioned last in workbook) with a combo chart:
  - Stacked column bars per year: Portfolio Withdrawal (blue), Social Security Income (green), Total Tax (red)
  - Orange line overlay for Net After-Tax Cash on a secondary y-axis
  - Data table in columns A–F backing the chart (Year, Ending Portfolio Balance, Portfolio Withdrawal, SS Income, Total Tax, Net After-Tax Cash)
- `openpyxl.chart` imports: `BarChart`, `LineChart`, `Reference`, `SeriesLabel`

### Changed
- **Roth Conversion Scenario Comparison** sheet now shows **Net After-Tax Cash** (`after_tax_cash`) instead of Net Spendable (`after_tax_net`) — comparison reflects tax impact only, without insurance costs distorting the comparison

---

## [Session 2] — prior session

### Added
- **3-phase withdrawal model**: Go-Go / Slow-Go / No-Go with configurable rates and years
- **Phase transition taper**: linear glide from Go-Go → Slow-Go over `transition_years` to smooth the spending cliff
- **Tax law database** (`TAX_LAW_DB`): versioned 2025 and 2026 entries with automatic fallback to latest available year
- **2026 tax law** (IRS Rev. Proc. 2025-32): updated federal brackets, LTCG thresholds, IRMAA tiers
- **OBBBA senior bonus** ($6,000/spouse at 65+) added to standard deduction
- **Oregon 2026** brackets and standard deduction ($9,680 MFJ) per HB 3753
- **ACA subsidy cliff** tracking (400% FPL hard cliff restored post-ARPA/IRA expiry Jan 2026)
- **Dynamic household size** for ACA FPL calculation — children age off the plan at 26
- **Roth Conversion Scenario Comparison** sheet: No Conversion vs 22% vs 24% side-by-side
- **Monte Carlo** sheet: 2,000 simulations, log-normal returns, bad-sequence (−35% year 3) scenario
- **Tax Tables** sheet: exact bracket values, IRMAA tiers, LTCG thresholds used in each run
- **SS Scenarios** sheet: lifetime SS comparison across Early (62/62), FRA (67/67), Delay Primary (70/67)
- **Insurance Detail** sheet: per-spouse ACA and Medicare costs by year, children on plan
- **RMD Projections** sheet: SECURE 2.0 age-75 start, IRS ULT factors, QCD notes
- **Strategy Summary** sheet: plain-language playbook and quarterly checklist
- `LABEL_ALIASES` for backwards-compatible reading of old Inputs sheets
- Stale sheet cleanup on each run (removes tabs from old script versions)

### Changed
- Removed dead `active_rate` field from simulation rows
- Updated docstring and inline comments throughout

---

## [v1.0] — initial release

### Added
- Core simulation engine: year-by-year projection of Trad IRA, Roth IRA, and Taxable brokerage
- Withdrawal order: Tax-Deferred → Taxable principal → Roth
- Roth conversion to fill configurable federal bracket ceiling
- Federal MFJ tax engine with inflation-adjusted brackets
- Oregon state income tax (ordinary + 9.9% dividends)
- SS taxability via IRS 2-tier provisional income test (IRC §86)
- LTCG / NIIT on dividends and taxable principal (Fed 15% + OR 9.9% + NIIT 3.8%)
- IRMAA surcharge calculation (2% annual threshold inflation, 2-year lookback)
- Social Security: FRA-based claiming with delay credits (62–70), spousal benefit floor (50% of primary PIA)
- ACA Gold premiums with age factors and 5%/yr inflation
- Medicare (Part B + Part D + Medigap G) with per-person tracking
- RMD enforcement per SECURE 2.0 (age 75)
- **Inputs** sheet: amber input cells, never overwritten
- **Assumptions** sheet: all parameters + 35-year summary totals
- **Annual Summary** sheet: full year-by-year table with group headers and color coding
- `README.md` with full usage documentation
