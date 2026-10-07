# Stock Portfolio Tracker and Risk Analytics: methodology

> The detailed write-up behind the short [README](../README.md): how each figure is built, how it is checked,
> and what the independent reviews found and fixed. Figures are as of the valuation date 2026-09-29, held in
> one input cell (`'Portfolio Series'!B24`, named `ValuationDate`) that every IRR and price window ends on.
> Changing that date and refreshing moves every figure; refreshing alone updates current prices.

## The portfolio

A paper portfolio of 16 stocks across 10 industries. The ledger holds 112 dated trades from 2019 to 2025:
80 purchases and 32 sales, each with price, units and amount. Every trade is dated 1 January of its year, a
nominal date for a paper portfolio rather than an exchange trading day. Prices come from the Excel Stocks data type and
`STOCKHISTORY`. No VBA, macros or add-ins.

## Sheets

| Sheet | What it does |
|---|---|
| Ledger | The 112 trades. The only hand-entered data in the workbook |
| Dashboard | Holdings table built from the ledger with dynamic arrays, KPI strip, sector allocation, rebalancing view |
| Analytics | Per-holding cost, value, unrealised and realised P&L, return and IRR; sector concentration; scorecard |
| Risk Analytics | CAPM, risk ratios, parametric and historical VaR, diversification, per-holding risk decomposition |
| Price History | 12 months of daily closes per holding, feeding the sparklines and the Portfolio Series |
| Portfolio Series | Today's holdings valued over the last 12 months, and the risk measures taken from that series |
| Validation | 26 in-sheet checks with a pass count that treats a cell in error as a failure |
| WL Dashboard, Watchlist | A 10-stock watchlist scored on P/E, beta and 52-week range |
| Stock sheets | Price history and charts for the 10 watchlist stocks (AMD is also a holding) |

## How each figure is built

**Holdings.** Units held = units bought - units sold, per stock (`Dashboard!D7`). Cost basis = units held x
average purchase cost, where average cost is total purchase amount / total units bought. Unrealised P&L =
market value - cost basis. Realised P&L = sale proceeds - units sold x average cost (`Analytics!I10:I25`).
Average cost pools every purchase of a stock, including purchases made after a sale. A running average, which
uses only the purchases made before each sale, is the more common convention; the validator prints both as an
advisory. Realised P&L is $14,944 pooled against $15,686 running, and cost basis $42,916 against $43,658;
total gain is the same under both.

**Returns.** Every annual return is a money-weighted XIRR: purchases are outflows, sales are inflows, and
the market value on the valuation date is the final inflow, all dated. This replaces an earlier fixed seven-year CAGR, which
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

**CAPM and ratios.** Risk-free rate 4.25% and market return 10.0% are labelled inputs (blue). Two betas are
shown and labelled: the value-weighted average of the Stocks data type betas (1.82), which CAPM, the excess
return and Treynor use, and the 12-month regression beta against SPY (1.64). The ratios use the 12-month return
and 12-month risk measures, so numerator and denominator cover the same window:

| Ratio | Definition | Value |
|---|---|---:|
| CAPM expected return | Rf + beta x (Rm - Rf) | 14.69% |
| Excess over CAPM expected return | 12-month return - CAPM expected return (alpha at the assumed Rm) | 21.04% |
| Sharpe | (12-month return - Rf) / volatility | 1.21 |
| Sortino | (12-month return - Rf) / downside deviation | 1.82 |
| Calmar | 12-month return / abs(maximum drawdown) | 2.41 |
| Parametric VaR 95%, 1 year | CAPM return - 1.645 x volatility | -28.0% |
| SPY volatility | sample stdev of SPY daily returns x sqrt(252) | 13.02% |
| Capital market line return | Rf + (SPY 12-month return - Rf) / SPY volatility x portfolio volatility | 26.40% |
| M-squared (excess over SPY) | Sharpe x SPY volatility + Rf - SPY 12-month return | 4.68% |

The parametric VaR uses the CAPM expected return as its mean; the historical VaR uses realised daily returns.
The capital market line and M-squared use SPY's realised return and volatility over the same 12 months.

**Money at risk.** Rows 70 to 76 of Risk Analytics put the exposure in dollars: the largest industry's value
($131,600), the loss if it fell by the shock input (30%: $39,480, 18.7% of the portfolio), the worst 12-month
drawdown of current holdings ($25,279) and the 1-day historical VaR ($5,645). The VaR is measured on the 15
holdings with price history and applied to the full portfolio value.

**Concentration.** After netting sales, NVDA is 45.1% of value, the effective number of holdings
(1 / holding-level HHI) is 4.05, and the sector HHI is 0.41.

## How it is checked

**In the workbook.** 26 checks on the Validation sheet: weights and sector weights sum to 100%, totals tie
across sheets, units equal bought minus sold, no holding is negative, realised plus unrealised equals total
gain, and the risk ratios agree between sheets and sit in plausible ranges. The pass count adds cells in error
to the failures, so a broken formula cannot hide behind ALL PASS.

**Outside the workbook.** `validate_portfolio.py` reads only raw inputs (the ledger, current prices, daily
closes, SPY closes, betas, Rf, Rm and the valuation date) and rebuilds 55 figures: every holding's units,
value, cost, unrealised and realised P&L, return, weight and IRR; the portfolio totals and IRR; the sector
table and HHI; every point of the 12-month series and each measure taken from it; and every CAPM figure and
ratio; the SPY volatility, capital market line and M-squared; and the money-at-risk block. It also checks
that betas are listed against the right tickers and that no formula still calls `TODAY()`. It prints the
headline figures it re-derived, so each number in the README can be matched on screen. It runs in CI on every
push.

**The validator is tested.** `tests/` copies the workbook, changes one cached value, and asserts that exact
figure is reported as wrong. Eighteen tests, including the defects below: sales added instead of netted, a
seven-year CAGR in place of the XIRR, a volatility not measured from prices, a missing realised P&L and a cell
in error, an M-squared computed with a typed market volatility, a wrong money-at-risk figure and betas listed
in the wrong order. `tests/xlsx_surgery.py` edits cached values in the sheet XML directly, because saving with openpyxl
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

A second review in October 2026 found and fixed the following. Only M-squared and the capital market line
changed value; the valuation date was set to the date of the last recalculation, so no other figure moved.

| Defect | Effect | Fix |
|---|---|---|
| Market volatility still typed as 0.155 in the CML and M-squared | M-squared showed 13.04% | SPY's realised volatility and return over the same 12 months: 4.68% |
| Every IRR and price window ended on `TODAY()`, and the labelled valuation-date cell was unused | Figures changed daily and could not be reproduced | One valuation-date input, named `ValuationDate`, used by all 54 formulas that had called `TODAY()` |
| The XOM sheet had no exchange or ticker | Its price series and two charts were empty | XNYS / XOM entered; fills on the next refresh |
| Market caps labelled "T" were billions | AMD showed as "$992T" | Formats corrected ($992.3B; watchlist total $7.31T) |
| The README's money-at-risk figures existed only in Python | A reviewer could not find them in the workbook | Money-at-risk block on Risk Analytics, checked by the validator |
| Leftover "CAGR" labels, an unlabelled second beta, an assumed Rm called a benchmark | Labels contradicted the formulas | Relabelled: IRR, vendor vs SPY beta, "assumed market return" |
| The in-workbook README described a $10,000-a-year plan with a 35% technology cap, GICS sectors and auto-extending formulas | None of it matched the ledger or the sheets | Rewritten to describe the ledger, the industry source and the 16-row blocks |

## Rebalancing scenarios

The README's recommendation, reduce the NVIDIA position, is priced by `scenarios.py` (tests in
`tests/test_scenarios.py`). It reads the holdings, average costs and daily closes the validator reads, and
measures any set of units with the validator's formulas. A test requires the unchanged portfolio to reproduce
the workbook's own sector-shock loss, drawdown, VaR, effective number of holdings and volatility, so the
scenarios rest on the validated arithmetic rather than a second implementation.

Three policies, each keeping the total value:

| Policy | Rule |
|---|---|
| NVIDIA capped at 25% | NVIDIA sold down to 25% of value; the proceeds spread over the other holdings in proportion to their value |
| No holding above 20% | the same rule applied to every holding, repeated until nothing is above the cap |
| Semiconductors capped at 40% | NVIDIA and AMD scaled down together, keeping their mix; the freed value spread over the rest |

Cost of each rebalance: realised gain = units sold x (price - average cost), tax = 15% of the net realised gain
(never negative), trading cost = 0.10% of the value bought and sold. Both rates are assumptions, constants at
the top of `scenarios.py`. Results (`outputs/scenarios.json`):

| | Current | NVIDIA 25% | Holdings 20% | Semis 40% |
|---|---:|---:|---:|---:|
| Loss if semiconductors fell 30% | 39,480 | 30,730 | 25,360 | 25,360 |
| 1-day historical VaR 95% | 5,645 | 4,591 | 4,312 | 4,189 |
| Worst 12-month drawdown | 25,279 | 22,256 | 22,185 | 22,698 |
| Volatility, 12-month back-cast | 26.0% | 23.3% | 21.0% | 20.0% |
| Beta vs SPY | 1.64 | 1.55 | 1.44 | 1.37 |
| Effective number of holdings | 4.1 | 7.0 | 8.7 | 7.6 |
| 12-month back-cast return | 35.7% | 39.9% | 34.9% | 26.5% |
| Tax + trading cost | 0 | 6,131 | 7,657 | 6,723 |

Per dollar of cost, the semiconductor cap removes the most shock loss (about $0.48 of cost per dollar removed,
against $0.70 for the NVIDIA cap alone, whose proceeds partly flow into AMD). The back-cast returns describe how
each mix would have behaved over the last year, not a forecast.

## Limitations

- **Paper portfolio.** The trades are hypothetical; fees, taxes, slippage and spreads are not modelled in the
  workbook. The rebalancing scenarios add an assumed tax and trading cost for the trades they propose.
- **Back-cast, not history.** The 12-month series values today's holdings over the past year. It describes
  the risk of the current portfolio, not the realised path of the portfolio as it was traded.
- **One holding without price history** is excluded from the series (2.1% of value).
- **Fixed 16-row blocks.** Holdings spill from the ledger, but the Dashboard's per-holding columns, Analytics
  rows 10 to 25 and Risk Analytics rows 39 to 54 hold 16 stocks and must be extended before a 17th is added.
- **Trade dates are nominal** (1 January each year), so IRRs are approximate to within the days to the first
  trading session.
- **Parametric VaR assumes normal returns**, which understates tail risk for a portfolio this concentrated;
  the historical VaR alongside it is the better guide.
- **Data depends on Microsoft 365.** The Stocks data type and `STOCKHISTORY` need an internet connection and
  a Microsoft 365 subscription; on first open without them, live cells show `#VALUE!` until refreshed.
