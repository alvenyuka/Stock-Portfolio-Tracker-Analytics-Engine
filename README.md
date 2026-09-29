# Stock-Portfolio-Tracker-Analytics-Engine

> How much has a 16-stock portfolio really earned, and where is its risk? An Excel 365 tracker rebuilt from its trade ledger: 32.2% a year since 2019 (money-weighted), with 45% of the value in one stock, checked by an independent Python validator that re-derives 45 of 45 figures.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel 365](https://img.shields.io/badge/Excel-365-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![workbook validated](https://github.com/alvenyuka/Stock-Portfolio-Tracker-Analytics-Engine/actions/workflows/ci.yml/badge.svg)](https://github.com/alvenyuka/Stock-Portfolio-Tracker-Analytics-Engine/actions/workflows/ci.yml)

![Stock portfolio dashboard: KPI strip, sector allocation, holdings table](MainDashboard.png)

## The problem

Portfolio reports are only as good as the ledger underneath them. A tracker that books transactions wrongly
can still pass every internal check, because those checks only confirm that one sheet agrees with another.
An investor or risk committee then sees returns, weights and risk figures that look precise and are wrong.

This workbook tracks a paper portfolio of 16 stocks bought and partly sold between 2019 and 2025. Every
figure is built from the trade ledger and live prices, and a separate program recomputes each one from the
raw inputs without trusting a single spreadsheet cell.

## What I found

Valued on 2026-09-29, from 112 trades (80 buys, 32 sells):

| Measure | Result |
|---|---:|
| Market value of current holdings | $211,335 |
| Annual return since 2019, money-weighted (XIRR) | **32.2%** |
| Total return including realised gains | 229.2% |
| 12-month return of current holdings vs SPY | 35.7% vs 15.4% |
| Volatility, 12 months of daily returns | 26.0% |
| Maximum drawdown, last 12 months | -14.8% |
| Sharpe ratio (12 months) | 1.21 |
| Largest single holding (NVDA) | **45.1% of value** |

- **The original workbook added sales to positions instead of netting them.** My independent review found it,
  and I rebuilt the accounting: holdings net of sales, average-cost basis and realised P&L. The corrected
  portfolio is worth $211,335, not the $302,909 the old sheets showed.
- **Concentration is the real risk.** Once sales are netted, NVDA is 45% of the portfolio and the effective
  number of holdings is about 4. Most of the return and most of the risk come from one position.
- **The validator caught Excel returning a false zero.** For one losing holding, Excel's XIRR silently
  returned 0% instead of the true -13.6%. The workbook now supplies a starting estimate and shows an error
  rather than a fake zero if XIRR ever fails again.

**What I would tell an investment committee:** the performance is real but driven by one stock; trim NVDA or
set a position limit, and recompute any spreadsheet model's figures from its transaction ledger before
relying on them.

## How it works

1. **Ledger to holdings.** Units held are units bought minus units sold; cost basis is net units at average
   purchase cost; each sale books realised P&L.
2. **Returns.** Money-weighted XIRR over the dated purchases and sales, with today's value as the final cash
   flow, for every holding and for the portfolio.
3. **Risk.** A *Portfolio Series* sheet values today's holdings over the last 12 months of daily closes, then
   measures volatility, downside deviation, drawdown, historical VaR, and beta, tracking error and
   information ratio against SPY.
4. **Checks.** 26 in-sheet checks, including units reconciling to the ledger and realised plus unrealised
   equalling total gain. The Python validator rebuilds all 45 figures from raw inputs, and its tests break the
   workbook on purpose to prove each fault is caught. Both run on every push.

## Run it

```bash
pip install -r requirements.txt
python validate_portfolio.py      # rebuild every figure from the ledger and prices
python -m pytest                  # confirm the validator catches injected faults
```

To explore the workbook, open `Stock Portfolio.xlsx` in Excel 365 and run Data, Refresh All for live prices.

## Limitations

- **Paper portfolio**: the trades are hypothetical, and fees, taxes and spreads are not modelled.
- **The risk figures back-cast today's holdings** over the last 12 months. They describe the current
  portfolio, not the path of the historical one.
- **One holding (0R2N, 2% of value) has no price history**, so the 12-month series covers 15 of 16 holdings.

## More detail

Every sheet, formula choice and check, and the full list of what the review found and fixed, is in
[`docs/METHODOLOGY.md`](docs/METHODOLOGY.md).

## License

MIT. See [`LICENSE`](LICENSE).

## Connect

Built by Alven Yuka, CPA Finalist and Accounting Specialist at GIZ, Nairobi.

📫 [alvenyuka2@gmail.com](mailto:alvenyuka2@gmail.com) · 💼 [LinkedIn](https://www.linkedin.com/in/alven-yuka-610b78174/) · 🐙 [GitHub](https://github.com/alvenyuka)
