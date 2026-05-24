# Stock-Portfolio-Tracker-Analytics-Engine

> Excel 365 workbook running CAPM, VaR (parametric / historical / Monte Carlo), Black–Litterman optimisation, and tax-aware lot matching on a 16-stock paper portfolio. No VBA, no macros, no add-ins.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Excel 365](https://img.shields.io/badge/Excel-365-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![Validation](https://img.shields.io/badge/validation-23%20checks%20required-success)](#validation)
[![Portfolio Grade](https://img.shields.io/badge/portfolio%20grade-B%2B-blue)](#portfolio-snapshot)
[![CAGR](https://img.shields.io/badge/7yr%20CAGR-14.89%25-brightgreen)](#portfolio-snapshot)

![Stock portfolio dashboard — KPI strip, sector allocation, holdings table](MainDashboard.png)

## Why?

Most retail portfolio trackers stop at total return and a pie chart. The risk-adjusted ratios, VaR, and tax-lot accounting that actually matter for sizing positions and reporting taxes live in either expensive platforms or scattered Python scripts. This workbook collapses that stack into one self-contained Excel 365 file: live prices via `STOCKHISTORY`, dynamic-array formulas, and a validation layer that refuses to render the dashboard until 23 integrity checks pass.

> **Note: this is a paper portfolio.** The transactions are hypothetical, picked for the project rather than executed through a brokerage account. Every metric in this workbook is derived from the hypothetical ledger, not from real positions.

## Features

- 📊 **Performance attribution** — per-stock CAGR, return decomposition, unrealised P&L, sector concentration (Herfindahl–Hirschman Index)
- 📉 **Risk analytics** — CAPM (Jensen's α, portfolio β), Sharpe / Treynor / Sortino / Calmar, parametric / historical / Monte Carlo VaR, Cornish–Fisher heavy-tail adjustment, drawdown + stress panel
- 🎯 **Black–Litterman optimisation** — market-equilibrium priors combined with user views to produce target weights and a BUY / SELL / HOLD trade list
- 🧾 **Tax-aware accounting** — lot matching via `MAXIFS`, IRS §1222 short-term / long-term classification, tax-adjusted return
- 👁️ **Watchlist scoring** — 10 competitor stocks ranked through a composite signal (P/E, β, 52-week range) → buy / watch / avoid
- 🛡️ **23 automated integrity checks** must all return `PASS` before the dashboard renders
- ⚡ **Live data** via Excel `Stocks` data type and `STOCKHISTORY` — no VBA, no macros, no add-ins

## Portfolio snapshot

| Metric | Value |
|---|---|
| Portfolio value | $303,680 |
| Cost basis | $132,051 |
| Total return | 129.97% |
| 7-year CAGR | 14.89% |
| Sharpe ratio | 0.47 |
| Sortino ratio | 3.10 |
| Portfolio β | 1.47 |
| Portfolio grade | **B+** |

## Tech Stack

| Layer | Tools |
|---|---|
| Spreadsheet | Microsoft Excel 365 |
| Live data | `STOCKHISTORY`, Stocks data type, dynamic arrays |
| Math | Native Excel functions only — no VBA, no macros, no add-ins |
| Optimisation | Black–Litterman (closed-form, in-sheet) |

## Sheet structure

| Sheet | Purpose |
|---|---|
| **Dashboard** | KPI strip, live 16-stock table, sector allocation, rebalancing panel, embedded charts |
| **Analytics** | Per-stock CAGR, return attribution, cost-basis breakdown, risk scores |
| **Risk Analytics** | CAPM, VaR, CVaR, drawdown, stress tests, factor exposure, Capital Market Line |
| **WL Dashboard** | Watchlist composite scoring with buy / watch / avoid signals |
| **Watchlist** | 10 competitor stocks with live data, 52-week ranges, β, P/E, market cap |
| **Optimization** | Black–Litterman panel with investor views and trade list |
| **Ledger** | 112 transaction records with holding period and LT/ST tax status |
| **Validation** | 23 automated integrity tests; all required `PASS` for dashboard render |
| **Skills Matrix** | Quantitative finance concepts the workbook exercises |
| **Sparkline** | Price history feeding the dashboard sparklines |
| **Stock Sheets** | Individual deep-dives per holding |

![Stock-level deep-dive sheet — price history, return decomposition, risk score](StockDashboard.png)

![Watchlist composite scoring with buy / watch / avoid output](WatchlistDashboard.png)

## Installation

No installation. Download `Stock Portfolio.xlsx` and open in Microsoft Excel 365 with an internet connection.

```bash
git clone https://github.com/alvenyuka/Stock-Portfolio-Tracker-Analytics-Engine.git
cd Stock-Portfolio-Tracker-Analytics-Engine
# Open Stock\ Portfolio.xlsx in Excel 365
```

## Usage

1. Open the workbook in **Microsoft Excel 365** with an internet connection.
2. Run **Data → Refresh All** to pull live prices, β, P/E, and market cap via the Stocks data type and `STOCKHISTORY`.
3. Review the **Dashboard**. New trades go on the **Ledger** and propagate everywhere automatically.
4. Enter market views on the **Optimization** sheet for revised Black–Litterman weights.
5. Check the **Validation** sheet — all 23 tests must read `PASS`.

## Validation

A dedicated sheet runs 23 integrity tests across the workbook. Every test must pass before the dashboard renders. Tests cover ledger reconciliation, balance sheet identities, return formula consistency, and inter-sheet ties.

## Roadmap

- [x] CAPM, VaR (parametric / historical / Monte Carlo), CVaR
- [x] Black–Litterman optimisation with investor views
- [x] Tax-aware lot matching under IRS §1222
- [x] 23-check validation harness
- [ ] Factor model (Fama–French 3 / 5) sheet
- [ ] Scenario stress library (rates +200bp, oil shock, USD/KES devaluation)

## License

MIT — see [`LICENSE`](LICENSE).

## Credits

Built with **Microsoft Excel 365** only.
Author: **Alven Yuka** — CPA Finalist (Kenya).

## Connect

📫 [alvenyuka2@gmail.com](mailto:alvenyuka2@gmail.com) · 💼 [LinkedIn](https://www.linkedin.com/in/alvenyuka) · 🐙 [GitHub](https://github.com/alvenyuka)
