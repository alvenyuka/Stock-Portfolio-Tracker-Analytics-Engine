![Stock Portfolio Analytics](banner.svg)

<p align="center">
  <img alt="Excel 365" src="https://img.shields.io/badge/Excel%20365-217346?logo=microsoftexcel&logoColor=white">
  <img alt="Black-Litterman" src="https://img.shields.io/badge/Black--Litterman-1F4E79">
  <img alt="VaR / CVaR" src="https://img.shields.io/badge/VaR%20%2F%20CVaR-2E5C8A">
  <img alt="Monte Carlo" src="https://img.shields.io/badge/Monte%20Carlo-5B7553">
  <img alt="License" src="https://img.shields.io/badge/license-MIT-blue">
</p>

---

## Overview

An Excel 365 workbook implementing the standard portfolio risk and performance analytics on a paper portfolio of 16 stocks held over a 7-year horizon. Built natively with `STOCKHISTORY`, the Stocks data type, and dynamic arrays. No VBA, no macros, no add-ins.

## Note: this is a paper portfolio

The transactions are hypothetical, picked for the project rather than executed through a brokerage account. Every metric in this workbook is derived from the hypothetical ledger, not from real positions.

## Portfolio snapshot

| Metric | Value |
|---|---|
| Portfolio value | $303,680 |
| Cost basis | $132,051 |
| Total return | 129.97% |
| 7-year CAGR | 14.89% |
| Sharpe ratio | 0.47 |
| Sortino ratio | 3.10 |
| Portfolio beta | 1.47 |
| Portfolio grade | B+ |

## What's in the workbook

### Performance attribution
Per-stock CAGR, return decomposition, unrealised gains and losses, sector concentration via the Herfindahl-Hirschman Index. Daily P&L and KPI cards on the dashboard, fed from live quotes via the Excel Stocks data type and `STOCKHISTORY`.

### Risk analytics
CAPM (Jensen's alpha and portfolio beta), the standard risk-adjusted return ratios (Sharpe, Treynor, Sortino, Calmar), and three flavours of Value-at-Risk: parametric, historical, and Monte Carlo. Cornish-Fisher VaR is included for heavy-tailed adjustment. Drawdown analysis and a stress-test panel sit alongside.

### Watchlist scoring
Ten competitor stocks ranked through a composite signal model (P/E, beta, 52-week range), with buy, watch, and avoid outputs.

### Portfolio optimisation
A Black-Litterman panel combining market-equilibrium priors with user-entered views to produce target weights and a BUY / SELL / HOLD trade list.

### Tax-aware accounting
Lot matching through `MAXIFS`, classifying each position as short-term or long-term under IRS section 1222 for the tax-adjusted return calculation.

### Validation
A dedicated sheet runs 23 integrity tests across the workbook. Every test must pass before the dashboard renders. Tests cover ledger reconciliation, balance sheet identities, return formula consistency, and inter-sheet ties.

## Sheet structure

| Sheet | Purpose |
|---|---|
| Dashboard | KPI strip, live 16-stock table, sector allocation, rebalancing panel, embedded charts |
| Analytics | Per-stock CAGR, return attribution, cost-basis breakdown, risk scores |
| Risk Analytics | CAPM, VaR, CVaR, drawdown, stress tests, factor exposure, Capital Market Line |
| WL Dashboard | Watchlist composite scoring with buy / watch / avoid signals |
| Watchlist | 10 competitor stocks with live data, 52-week ranges, beta, P/E, market cap |
| Optimization | Black-Litterman panel with investor views and trade list |
| Ledger | 112 transaction records with holding period and LT/ST tax status |
| Validation | 23 automated integrity tests, all required PASS for dashboard render |
| Skills Matrix | Quantitative finance concepts the workbook exercises |
| Sparkline | Price history feeding the dashboard sparklines |
| Stock Sheets | Individual deep-dives per holding |

## How to use

Open in Microsoft Excel 365 with an internet connection. Run Data, then Refresh All, to pull live prices, beta, P/E, and market cap. Review the Dashboard. New trades go on the Ledger and propagate everywhere automatically. Enter market views on the Optimization sheet for revised Black-Litterman weights. Check the Validation sheet; all 23 tests should read PASS.

## Built with

Microsoft Excel 365 only. Dynamic arrays, Stocks data type, `STOCKHISTORY`, standard formulas. No VBA, no macros, no add-ins.

## License

MIT.

Author: Alven Yuka.
