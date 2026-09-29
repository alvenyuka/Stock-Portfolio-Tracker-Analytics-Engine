# Stock-Portfolio-Tracker-Analytics-Engine: full methodology

> The detailed write-up behind the short [README](../README.md): how each figure is built, how it is checked,
> and what the independent review found and fixed. Figures are as of the valuation date 2026-09-29, the last
> recalculation saved in the workbook. Live prices move them on every refresh.

## The portfolio

A paper portfolio of 16 stocks across 10 industries. The ledger holds 112 dated trades from 2019 to 2025:
80 purchases and 32 sales, each with price, units and amount. Prices come from the Excel Stocks data type and
`STOCKHISTORY`. No VBA, macros or add-ins.

## Sheets

| Sheet | What it does |
|---|---|
| Ledger | The 112 trades. The only hand-entered data in the workbook |
| Dashboard | Holdings table built from the ledger with dynamic arrays, KPI strip, sector allocation, rebalancing view |
| Analytics | Per-holding cost, value, unrealised and realised P&L, return and IRR; sector concentration; scorecard |
| Risk Analytics | CAPM, risk ratios, parametric and historical VaR, diversification, per-holding risk decomposition |
| Spartkine | 12 months of daily closes per holding, feeding the sparklines and the Portfolio Series |
| Portfolio Series | Today's holdings valued over the last 12 months, and the risk measures taken from that series |
| Validation | 26 in-sheet checks with a pass count that treats a cell in error as a failure |
| WL Dashboard, Watchlist | A 10-stock watchlist scored on P/E, beta and 52-week range |
| Stock sheets | Price history and charts for 10 of the holdings |

## How each figure is built

**Holdings.** Units held = units bought - units sold, per stock (`Dashboard!D7`). Cost basis = units held x
average purchase cost, where average cost is total purchase amount / total units bought. Unrealised P&L =
market value - cost basis. Realised P&L = sale proceeds - units sold x average cost (`Analytics!I10:I25`).

**Returns.** Every annual return is a money-weighted XIRR: purchases are outflows, sales are inflows, and
today's market value is the final inflow, all dated. This replaces an earlier fixed seven-year CAGR, which
treated purchases made as late as 2025 as if held since 2019. XIRR is given a starting estimate in the
direction of the gain or loss, and a result that fails to converge shows as an error rather than zero.
Total return = (market value + sale proceeds - purchases) / purchases.

**The 12-month series.** `Portfolio Series` multiplies each holding's units by its daily close and sums them,
giving a daily value for today's portfolio over the last 12 months (250 daily returns). From it:

| Measure | Definition | Value |
|---|---|---:|
| 12-month return | last value / first value - 1 | 35.73% |
| Volatility | sample standard deviation of daily returns x sqrt(252) | 25.97% |
| Downside deviation | sqrt(mean of squared negative daily returns) x sqrt(252) | 17.31% |
| Maximum drawdown | worst fall from a running peak | -14.84% |
| Historical VaR, 1-day, 95% | minus the 5th percentile daily return | 2.67% |
| Beta vs SPY | slope of portfolio on SPY daily returns | 1.64 |
| Tracking error vs SPY | stdev of daily active returns x sqrt(252) | 16.91% |
| Information ratio | (12-month return - SPY 12-month return) / tracking error | 1.20 |

One holding, 0R2N (the London listing of RTX Corporation, 2.1% of value), returns no price history from
`STOCKHISTORY`, so the series covers 15 holdings and 97.9% of market value.

**CAPM and ratios.** Risk-free rate 4.25% and market return 10.0% are labelled inputs. Portfolio beta is the
value-weighted average of the Stocks data type betas (1.82). The ratios use the 12-month return and 12-month
risk measures, so numerator and denominator cover the same window:

| Ratio | Definition | Value |
|---|---|---:|
| CAPM expected return | Rf + beta x (Rm - Rf) | 14.69% |
| Jensen's alpha | 12-month return - CAPM expected return | 21.04% |
| Sharpe | (12-month return - Rf) / volatility | 1.21 |
| Sortino | (12-month return - Rf) / downside deviation | 1.82 |
| Calmar | 12-month return / abs(maximum drawdown) | 2.41 |
| Parametric VaR 95%, 1 year | CAPM return - 1.645 x volatility | -28.0% |

**Concentration.** After netting sales, NVDA is 45.1% of value, the effective number of holdings
(1 / holding-level HHI) is 4.05, and the sector HHI is 0.41.

## How it is checked

**In the workbook.** 26 checks on the Validation sheet: weights and sector weights sum to 100%, totals tie
across sheets, units equal bought minus sold, no holding is negative, realised plus unrealised equals total
gain, and the risk ratios agree between sheets and sit in plausible ranges. The pass count adds cells in error
to the failures, so a broken formula cannot hide behind ALL PASS.

**Outside the workbook.** `validate_portfolio.py` reads only raw inputs (the ledger, current prices, daily
closes, SPY closes, betas, Rf, Rm and the valuation date) and rebuilds 45 figures: every holding's units,
value, cost, unrealised and realised P&L, return, weight and IRR; the portfolio totals and IRR; the sector
table and HHI; every point of the 12-month series and each measure taken from it; and every CAPM figure and
ratio. It runs in CI on every push.

**The validator is tested.** `tests/` copies the workbook, changes one cached value, and asserts that exact
figure is reported as wrong. Fifteen tests, including the defects below: sales added instead of netted, a
seven-year CAGR in place of the XIRR, a volatility not measured from prices, a missing realised P&L and a cell
in error. `tests/xlsx_surgery.py` edits cached values in the sheet XML directly, because saving with openpyxl
would discard every cached value.

## What the review found and fixed

An earlier version of this workbook, and of this README, published figures built on these defects. All are
now fixed and the figures above are restated.

| Defect | Effect | Fix |
|---|---|---|
| Sales were added to positions (`SUMIF` over all units) | Units, value, cost, weights, returns and every ratio were wrong; value showed $302,909 | Net of sales, average-cost basis, realised P&L |
| Annual return was a fixed seven-year CAGR | Purchases from 2025 were treated as held since 2019 | Money-weighted XIRR over dated trades |
| XIRR silently returned 0 for BA | A holding losing 13.6% a year showed as break-even | Starting estimate, and an error instead of zero |
| Volatility was beta x a hardcoded 0.155 | No stock-specific risk; Sharpe, VaR and CML were off | Measured from 12 months of daily returns |
| "Max drawdown" was the worst holding's return | Calmar was not a Calmar ratio | Peak-to-trough fall of the 12-month series |
| Tracking error was a stdev minus a return | Negative tracking error, wrong-signed information ratio | Measured against SPY |
| Sortino disagreed between sheets (37.2 vs 2.6) | Two definitions under one name | One definition, one source cell |
| GOOG's price row read AMZN's prices | GOOG's sparkline showed another stock | Corrected reference |
| Today's % change and daily P&L mixed per-share and position values | Both figures meaningless | Change / previous close; units x change |
| A benchmark cell was empty | "Outperforming" was always shown | Benchmark set to the market return input |
| The pass count skipped cells in error | ALL PASS with a broken Calmar cell | Errors counted as failures |
| Labels overstated what was measured | "Style factors", "Max 1Y loss", two different "effective diversification" figures | Labels now describe the calculation |

## Limitations

- **Paper portfolio.** The trades are hypothetical; fees, taxes, slippage and spreads are not modelled.
- **Back-cast, not history.** The 12-month series values today's holdings over the past year. It describes
  the risk of the current portfolio, not the realised path of the portfolio as it was traded.
- **One holding without price history** is excluded from the series (2.1% of value).
- **Parametric VaR assumes normal returns**, which understates tail risk for a portfolio this concentrated;
  the historical VaR alongside it is the better guide.
- **Data depends on Microsoft 365.** The Stocks data type and `STOCKHISTORY` need an internet connection and
  a Microsoft 365 subscription; on first open without them, live cells show `#VALUE!` until refreshed.
