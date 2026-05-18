# Stock Portfolio Tracker & Analytics Engine

An Excel 365 workbook for tracking and analysing a paper portfolio of 16 stocks over a 7-year horizon, with the standard risk and performance analytics built on top.

## This is a paper portfolio

The transactions in this workbook are hypothetical. They were picked for the project rather than executed through a brokerage account. The point of the workbook is to work through quantitative finance methods on a plausible set of holdings, not to track a real position. Every metric is computed from the hypothetical ledger.

## Snapshot

| Metric | Value |
|---|---|
| Portfolio value | $303,680 |
| Cost basis | $132,051 |
| Total return | 129.97% |
| 7-year CAGR | 14.89% |
| Sharpe ratio | 0.47 |
| Sortino ratio | 3.10 |
| Portfolio beta | 1.47 |

## What's in it

A dashboard with KPI cards and daily P&L, fed from the Excel Stocks data type and `STOCKHISTORY` (no VBA, no macros). A transaction ledger of 112 buy/sell records across the 16 stocks over 7 years. Per-stock analytics for CAGR, return attribution, unrealised gains and losses, and sector concentration via an HHI index.

On the risk side: CAPM with Jensen's alpha and beta, Sharpe, Treynor, Sortino, and Calmar ratios, parametric, historical, and Monte Carlo VaR, CVaR, and a Cornish-Fisher VaR variant for heavy-tailed adjustment. Drawdown analysis and a stress-test panel sit alongside the standard reports.

A watchlist of 10 competitor stocks runs a composite scoring model (P/E, beta, 52-week range) with buy/watch/avoid signals. A Black-Litterman panel combines market-equilibrium priors with user-entered views to suggest target weights.

A tax-aware ledger does lot matching via `MAXIFS` and classifies positions as short- or long-term under IRS §1222. A validation sheet runs 23 integrity checks across the workbook and reports pass/fail. The dashboard will not render until all 23 pass.

## Sheet structure

| Sheet | Description |
|---|---|
| Dashboard | KPI strip, live 16-stock table, sector allocation, rebalancing panel, three embedded charts |
| Analytics | Per-stock CAGR, return attribution, cost basis breakdown, risk scores |
| Risk Analytics | CAPM, VaR, CVaR, drawdown analysis, stress tests, factor exposure, Capital Market Line |
| WL Dashboard | Watchlist scoring model with composite signals |
| Watchlist | 10 competitor stocks with live data, 52-week ranges, beta, P/E, market cap |
| Optimization | Black-Litterman with investor view panel and BUY/SELL/HOLD trade list |
| Ledger | 112 transaction records with holding period and LT/ST tax status |
| Validation | 23 automated integrity tests; all must pass for the dashboard to render |
| Skills Matrix | Quantitative finance skills the workbook exercises |
| Sparkline | Price history feeding the dashboard sparklines |
| Stock Sheets | Individual deep-dives per holding |

## How to use it

Open in Microsoft Excel 365 with an internet connection. Use Data → Refresh All to pull live prices, beta, P/E, and market cap. Review the Dashboard sheet. New trades go on the Ledger and the rest propagates. Enter your own market views on the Optimization sheet for revised Black-Litterman weights. Check the Validation sheet; all 23 tests should read PASS.

## Built with

Microsoft Excel 365 only. Dynamic arrays, Stocks data type, `STOCKHISTORY`, standard formulas. No VBA, no macros, no add-ins.

## License

MIT, for educational and portfolio purposes.

Author: Alven Yuka
