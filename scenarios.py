"""What would reducing the concentration cost, and what would it buy?

The README's finding is that the portfolio's return is mostly one bet: NVIDIA is 45% of
its value and semiconductors 62%, and it names trimming NVIDIA as the lever that changes
the risk most. This script puts numbers on that recommendation. It rebalances the
current holdings under three policies, then measures each one with the same arithmetic
validate_portfolio.py uses (and checks that the current portfolio reproduces the
validator's figures exactly):

    NVIDIA capped at 25%         sell NVIDIA down to 25% of value, spread the proceeds
                                 over the other holdings in proportion to their value
    no holding above 20%         the same rule applied to every holding
    semiconductors capped at 40% scale both semiconductor holdings down together

For each policy it reports the money at risk (largest-sector shock, 1-day historical VaR,
worst 12-month drawdown), concentration (largest holding and sector, effective number of
holdings), volatility and beta against SPY, and what getting there would cost.

Assumptions, stated here because they drive the cost lines and are not in the workbook:

    TAX_RATE      15% on net realised gains (a common long-term capital-gains rate; set it
                  to the investor's own)
    TRADING_COST  0.10% of the value traded, both sides

The risk figures use the same 12-month back-cast as the workbook: today's (or the
rebalanced) units valued at each of the last 251 closes. RTX's London listing has no
price history and stays out of the back-cast, as it does in the workbook. The back-cast
describes how the holdings would have behaved, not returns that were earned.

Writes outputs/scenarios.json and figures/rebalancing.png.

Usage:  python scenarios.py [path/to/Book.xlsx]
"""
from __future__ import annotations

import json
import math
import statistics
import sys
from pathlib import Path

from validate_portfolio import (
    DASH_FIRST,
    DASH_LAST,
    DEFAULT,
    MONEY,
    SPARK_FIRST,
    TRADING_DAYS,
    load,
    num,
    percentile_inc,
    positions,
    read_ledger,
)

TAX_RATE = 0.15
TRADING_COST = 0.001
NVIDIA = "NVDA"
SEMIS = "Semiconductors & Semiconductor Equipment"
ROOT = Path(__file__).parent


def read_holdings(wb) -> tuple[list[dict], list[float], float]:
    """The current holdings with their sector, net units, price, average cost and daily closes.

    Units and average cost come from the Ledger (net of sales), as in the validator. Returns
    the holdings, the SPY closes over the same days, and the sector-shock input.
    """
    dash, spark, ps = wb["Dashboard"], wb["Price History"], wb["Portfolio Series"]
    pos = positions(read_ledger(wb["Ledger"]))

    # Count the days that have a date in row 2 of Price History (columns B onwards).
    ndays = 0
    for col in range(2, spark.max_column + 1):
        if num(spark.cell(2, col).value) is not None:
            ndays += 1

    holdings = []
    for i, dash_row in enumerate(range(DASH_FIRST, DASH_LAST + 1)):
        ticker = dash.cell(dash_row, 13).value
        closes = []
        for col in range(2, 2 + ndays):
            closes.append(num(spark.cell(SPARK_FIRST + i, col).value))
        if closes[0] is None:
            closes = None    # no price history for this holding, so it stays out of the back-cast
        holdings.append({
            "ticker": ticker,
            "sector": dash.cell(dash_row, 3).value,
            "units": pos[ticker]["net_units"],
            "price": num(dash.cell(dash_row, 5).value),
            "avg_cost": pos[ticker]["avg_cost"],
            "closes": closes,
        })

    spy = []
    for col in range(2, 2 + ndays):
        spy.append(num(ps.cell(7, col).value))
    shock = num(wb["Risk Analytics"][MONEY["shock"]].value)
    return holdings, spy, shock


def profile(holdings: list[dict], units: dict, spy: list[float], shock: float) -> dict:
    """Concentration and money-at-risk figures for one set of units (ticker -> units)."""
    # market value of each holding, and of each sector
    market_value = {}
    sector_value = {}
    for h in holdings:
        value = units[h["ticker"]] * h["price"]
        market_value[h["ticker"]] = value
        sector_value[h["sector"]] = sector_value.get(h["sector"], 0.0) + value
    total = sum(market_value.values())

    weights = {}
    for ticker, value in market_value.items():
        weights[ticker] = value / total
    top_ticker = max(market_value, key=market_value.get)
    top_sector = max(sector_value, key=sector_value.get)
    squared_weights = []
    for w in weights.values():
        squared_weights.append(w * w)
    sum_squared_weights = sum(squared_weights)

    # 12-month back-cast: these units valued at each day's close (holdings with no history are left out)
    covered = []
    for h in holdings:
        if h["closes"] is not None:
            covered.append(h)
    daily_value = []
    for day in range(len(spy)):
        holding_values = []
        for h in covered:
            holding_values.append(units[h["ticker"]] * h["closes"][day])
        daily_value.append(sum(holding_values))

    returns = []
    for day in range(1, len(daily_value)):
        returns.append(daily_value[day] / daily_value[day - 1] - 1)
    spy_returns = []
    for day in range(1, len(spy)):
        spy_returns.append(spy[day] / spy[day - 1] - 1)

    # max drawdown: the worst fall from the highest value seen so far
    peak = daily_value[0]
    max_drawdown = 0.0
    max_drawdown_usd = 0.0
    for v in daily_value:
        peak = max(peak, v)
        max_drawdown = min(max_drawdown, v / peak - 1)
        max_drawdown_usd = min(max_drawdown_usd, v - peak)

    # beta = covariance of the portfolio with SPY / variance of SPY
    mean_portfolio = statistics.fmean(returns)
    mean_spy = statistics.fmean(spy_returns)
    covariances = []
    for portfolio_return, spy_return in zip(returns, spy_returns):
        covariances.append((portfolio_return - mean_portfolio) * (spy_return - mean_spy))
    spy_variances = []
    for spy_return in spy_returns:
        spy_variances.append((spy_return - mean_spy) ** 2)
    beta = sum(covariances) / sum(spy_variances)

    return {
        "market_value": total,
        "largest_holding": top_ticker,
        "largest_holding_share": weights[top_ticker],
        "nvidia_share": weights.get(NVIDIA, 0.0),
        "largest_sector": top_sector,
        "largest_sector_share": sector_value[top_sector] / total,
        "semiconductor_share": sector_value.get(SEMIS, 0.0) / total,
        "effective_holdings": 1 / sum_squared_weights,
        "sector_shock_loss": shock * sector_value[top_sector],
        "semiconductor_shock_loss": shock * sector_value.get(SEMIS, 0.0),
        "volatility": statistics.stdev(returns) * math.sqrt(TRADING_DAYS),
        "max_drawdown": max_drawdown,
        "max_drawdown_usd": max_drawdown_usd,
        "var_95_1d_usd": -percentile_inc(returns, 0.05) * total,
        "beta_vs_spy": beta,
        "return_12m_backcast": daily_value[-1] / daily_value[0] - 1,
    }


def cap_weights(values: dict, cap: float, only: set | None = None) -> dict:
    """Target values with each holding (or each holding in `only`) at most `cap` of the total.

    The excess is spread over the holdings still below the cap in proportion to their value,
    and the rule is reapplied until nothing is above it. The total is unchanged.
    """
    if only is None and cap * len(values) < 1:
        raise ValueError(f"a {cap:.0%} cap cannot hold for {len(values)} holdings: they must sum to 100%")
    total = sum(values.values())
    limit = cap * total
    target = dict(values)
    capped = set()

    while True:
        # holdings still above the limit (the 1e-9 absorbs floating-point noise)
        over = []
        for ticker, value in target.items():
            in_scope = only is None or ticker in only
            if in_scope and ticker not in capped and value > limit + 1e-9:
                over.append(ticker)
        if not over:
            return target

        for ticker in over:
            target[ticker] = limit
            capped.add(ticker)

        # give the freed value to the holdings that are not capped, in proportion to their value
        free = []
        for ticker in target:
            if ticker not in capped:
                free.append(ticker)
        excess = total - sum(target.values())
        free_values = []
        for ticker in free:
            free_values.append(target[ticker])
        free_total = sum(free_values)
        for ticker in free:
            target[ticker] += excess * target[ticker] / free_total


def cap_group(values: dict, group: set, cap: float) -> dict:
    """Target values with the holdings in `group` together at most `cap` of the total.

    The group is scaled down as a whole (keeping its internal mix) and the freed value is
    spread over the other holdings in proportion to their value.
    """
    total = sum(values.values())
    limit = cap * total
    group_values = []
    for ticker, value in values.items():
        if ticker in group:
            group_values.append(value)
    group_value = sum(group_values)
    if group_value <= limit:
        return dict(values)

    scale = limit / group_value
    others_value = total - group_value
    freed = group_value - limit
    target = {}
    for ticker, value in values.items():
        if ticker in group:
            target[ticker] = value * scale
        else:
            target[ticker] = value + freed * value / others_value
    return target


def trade_costs(holdings: list[dict], target: dict) -> dict:
    """Value bought and sold to reach `target` (ticker -> value), tax on net realised gains, and trading cost."""
    sold = bought = gains = 0.0
    for h in holdings:
        now = h["units"] * h["price"]
        change = target[h["ticker"]] - now
        if change < 0:
            units_sold = -change / h["price"]
            sold += -change
            gains += units_sold * (h["price"] - h["avg_cost"])
        else:
            bought += change
    tax = TAX_RATE * max(gains, 0.0)
    trading = TRADING_COST * (sold + bought)
    return {"sold": sold, "bought": bought, "realised_gain": gains, "tax": tax,
            "trading_cost": trading, "total_cost": tax + trading}


def run(path: Path = DEFAULT) -> dict:
    """Profile the current portfolio and the three rebalancing policies."""
    wb = load(path)
    holdings, spy, shock = read_holdings(wb)

    values = {}
    semis = set()
    for h in holdings:
        values[h["ticker"]] = h["units"] * h["price"]
        if h["sector"] == SEMIS:
            semis.add(h["ticker"])

    # each policy is a target market value per ticker
    policies = {
        "Current portfolio": values,
        "NVIDIA capped at 25%": cap_weights(values, 0.25, only={NVIDIA}),
        "No holding above 20%": cap_weights(values, 0.20),
        "Semiconductors capped at 40%": cap_group(values, semis, 0.40),
    }

    result = {"assumptions": {"tax_rate_on_net_realised_gains": TAX_RATE,
                              "trading_cost_share_of_value_traded": TRADING_COST,
                              "sector_shock": shock},
              "scenarios": {}}
    for name, target in policies.items():
        units = {}
        for h in holdings:
            units[h["ticker"]] = target[h["ticker"]] / h["price"]
        scenario = profile(holdings, units, spy, shock)
        scenario.update(trade_costs(holdings, target))
        result["scenarios"][name] = scenario
    return result


def plot(result: dict, path: Path) -> None:
    """Money at risk and the one-off cost of each policy, side by side."""
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    scenarios = result["scenarios"]
    names = list(scenarios)

    def column(key, sign=1):
        """One figure under each policy (sign=-1 turns a negative drawdown into a positive loss)."""
        figures = []
        for name in names:
            figures.append(sign * scenarios[name][key])
        return figures

    # (legend label, values under each policy, colour)
    series = [("Loss if semiconductors fell 30%", column("semiconductor_shock_loss"), "#c0392b"),
              ("Worst 12-month drawdown", column("max_drawdown_usd", sign=-1), "#dd8452"),
              ("1-day VaR (95%)", column("var_95_1d_usd"), "#4c72b0"),
              ("One-off cost: tax + trading", column("total_cost"), "#7f7f7f")]
    fig, ax = plt.subplots(figsize=(9.5, 4.6))
    width = 0.2
    for i, (label, vals, colour) in enumerate(series):
        # four bars per policy, centred on the policy's tick
        xs = []
        for k in range(len(names)):
            xs.append(k + (i - 1.5) * width)
        thousands = []
        for v in vals:
            thousands.append(v / 1000)
        ax.bar(xs, thousands, width, label=label, color=colour)
    ax.set_xticks(range(len(names)))
    ax.set_xticklabels(names, fontsize=9)
    ax.set_ylabel("US$ thousands")
    ax.set_title("Money at risk under each rebalancing policy, and what the rebalance costs", loc="left", fontsize=11)
    ax.legend(frameon=False, fontsize=8, ncol=2)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def main(argv) -> int:
    """Run the scenarios, print the comparison and write the JSON and chart."""
    path = Path(argv[1]) if len(argv) > 1 else DEFAULT
    result = run(path)
    # (row label, key in each scenario, number format)
    table_rows = [("market value", "market_value", "{:,.0f}"),
                  ("largest holding share", "largest_holding_share", "{:.1%}"),
                  ("semiconductor share", "semiconductor_share", "{:.1%}"),
                  ("effective number of holdings", "effective_holdings", "{:.1f}"),
                  ("loss if semiconductors fell 30%", "semiconductor_shock_loss", "{:,.0f}"),
                  ("1-day VaR 95%", "var_95_1d_usd", "{:,.0f}"),
                  ("worst 12-month drawdown", "max_drawdown_usd", "{:,.0f}"),
                  ("volatility (12m back-cast)", "volatility", "{:.1%}"),
                  ("beta vs SPY", "beta_vs_spy", "{:.2f}"),
                  ("12-month return (back-cast)", "return_12m_backcast", "{:.1%}"),
                  ("value sold", "sold", "{:,.0f}"),
                  ("tax on realised gains", "tax", "{:,.0f}"),
                  ("trading cost", "trading_cost", "{:,.0f}")]
    names = list(result["scenarios"])

    header = f"{'':34s}"
    for name in names:
        header += f"{name[:22]:>24s}"
    print(header)
    for label, key, number_format in table_rows:
        line = f"{label:34s}"
        for name in names:
            cell = number_format.format(result["scenarios"][name][key])
            line += f"{cell:>24s}"
        print(line)

    assumptions = result["assumptions"]
    print(f"\nAssumptions: tax {assumptions['tax_rate_on_net_realised_gains']:.0%} on net realised gains, "
          f"trading cost {assumptions['trading_cost_share_of_value_traded']:.2%} of value traded, "
          f"sector shock {assumptions['sector_shock']:.0%}.")
    (ROOT / "outputs").mkdir(exist_ok=True)
    (ROOT / "outputs" / "scenarios.json").write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    (ROOT / "figures").mkdir(exist_ok=True)
    plot(result, ROOT / "figures" / "rebalancing.png")
    print("wrote outputs/scenarios.json and figures/rebalancing.png")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
