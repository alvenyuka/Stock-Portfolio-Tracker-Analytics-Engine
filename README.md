# Stock-Portfolio-Tracker-Analytics-Engine

> Would a portfolio model's own checks catch a basic ledger error? An Excel 365 portfolio tracker with CAPM, value-at-risk and concentration analytics, reviewed by an independent Python validator that found the workbook adds sales to positions instead of netting them, an error all 23 of its built-in checks pass.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel 365](https://img.shields.io/badge/Excel-365-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![workbook validated](https://github.com/alvenyuka/Stock-Portfolio-Tracker-Analytics-Engine/actions/workflows/ci.yml/badge.svg)](https://github.com/alvenyuka/Stock-Portfolio-Tracker-Analytics-Engine/actions/workflows/ci.yml)

![Stock portfolio dashboard: KPI strip, sector allocation, holdings table](MainDashboard.png)

## The problem

Portfolio and risk reports are only as good as the ledger underneath them. Spreadsheet models usually
check themselves by comparing one total with another, which confirms the sheets agree but not that the
accounting is right. An error in how transactions are booked can pass every internal check and still
flow into every return and risk figure.

I built a 16-stock paper portfolio tracker in Excel 365, then reviewed it the way an auditor would: with a
separate program that recomputes every figure from the raw inputs and never trusts a "PASS" cell.

## What I found

| Review result | Value |
|---|---:|
| Transactions in the ledger | 112 (80 buys, 32 sells) |
| Figures re-derived outside Excel that match the workbook | **28 of 28** |
| Built-in workbook checks passing | 23 of 23 |
| Units held as the workbook counts them vs correctly netted | **2,441.55 vs 1,008.21** |
| Known definitional defects the validator reports on every run | 4 |

- **The workbook adds sales instead of subtracting them.** Every sell row carries positive units, so each
  sale increases the position and the cost basis. Market value, returns, beta and every risk ratio inherit
  the error, so **none of the workbook's performance figures are quoted here**.
- **Internal consistency is not correctness.** The 23 built-in checks and the 28 independent re-derivations
  all agree, because they test whether the workbook computes what its own formulas say, not whether those
  formulas book transactions correctly.
- **Three risk measures are mislabelled**: the "maximum drawdown" is the worst single holding's return, the
  volatility behind Sharpe and VaR leaves out stock-specific risk, and the tracking error is negative. The
  validator names each one rather than letting a green badge hide it.

**What I would tell a risk or finance team:** before relying on a spreadsheet model, recompute its key
figures from the transaction ledger independently, and test the booking logic itself, not just whether the
totals tie.

## How it works

1. **Excel 365 workbook**: a 112-row transaction ledger, live prices via `STOCKHISTORY`, and Dashboard,
   Analytics and Risk Analytics sheets covering CAPM, parametric VaR and expected shortfall, sector
   concentration and a 10-stock watchlist scorer. No VBA, macros or add-ins.
2. **Independent validator** (`validate_portfolio.py`) reads only the inputs, recomputes 28 derived figures
   in Python, and prints the known defects with measured numbers on every run.
3. **The validator is tested**: its tests corrupt a copy of the workbook one cell at a time and confirm each
   fault is reported. Both run on every push.

## Run it

```bash
pip install -r requirements.txt
python validate_portfolio.py      # review the committed workbook
python -m pytest                  # confirm the validator catches injected faults
```

To explore the workbook, open `Stock Portfolio.xlsx` in Excel 365 and run Data, Refresh All to reconnect live
prices.

## Limitations

- **Sales are not netted**, so every performance and risk figure in the workbook is wrong until the ledger
  logic is fixed and the figures restated. That is the next step.
- **Paper portfolio**: the transactions are hypothetical, not executed trades.
- **Risk ratios are a snapshot across holdings**, not computed from a portfolio value series over time, and
  only the parametric form of VaR is built.

## More detail

The full review, sheet by sheet, with every defect and what each metric really measures, is in
[`docs/METHODOLOGY.md`](docs/METHODOLOGY.md).

## License

MIT. See [`LICENSE`](LICENSE).

## Connect

Built by Alven Yuka, CPA Finalist and Accounting Specialist at GIZ, Nairobi.

📫 [alvenyuka2@gmail.com](mailto:alvenyuka2@gmail.com) · 💼 [LinkedIn](https://www.linkedin.com/in/alven-yuka-610b78174/) · 🐙 [GitHub](https://github.com/alvenyuka)
