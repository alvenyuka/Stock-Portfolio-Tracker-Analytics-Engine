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
    ndays = sum(1 for c in range(2, spark.max_column + 1) if num(spark.cell(2, c).value) is not None)
    holdings = []
    for i, r in enumerate(range(DASH_FIRST, DASH_LAST + 1)):
        tk = dash.cell(r, 13).value
        closes = [num(spark.cell(SPARK_FIRST + i, c).value) for c in range(2, 2 + ndays)]
        holdings.append({
            "ticker": tk,
            "sector": dash.cell(r, 3).value,
            "units": pos[tk]["net"],
            "price": num(dash.cell(r, 5).value),
            "avg_cost": pos[tk]["avg"],
            "closes": closes if closes[0] is not None else None,
        })
    spy = [num(ps.cell(7, c).value) for c in range(2, 2 + ndays)]
    shock = num(wb["Risk Analytics"][MONEY["shock"]].value)
    return holdings, spy, shock


def profile(holdings: list[dict], units: dict, spy: list[float], shock: float) -> dict:
    """Concentration and money-at-risk figures for one set of units (ticker -> units)."""
    mv = {h["ticker"]: units[h["ticker"]] * h["price"] for h in holdings}
    total = sum(mv.values())
    weights = {t: v / total for t, v in mv.items()}
    sectors = {}
    for h in holdings:
        sectors[h["sector"]] = sectors.get(h["sector"], 0.0) + mv[h["ticker"]]
    top_tk = max(mv, key=mv.get)
    top_sector = max(sectors, key=sectors.get)

    covered = [h for h in holdings if h["closes"] is not None]
    value = [sum(units[h["ticker"]] * h["closes"][k] for h in covered) for k in range(len(spy))]
    rets = [value[k] / value[k - 1] - 1 for k in range(1, len(value))]
    srets = [spy[k] / spy[k - 1] - 1 for k in range(1, len(spy))]
    peak, mdd, mdd_usd = value[0], 0.0, 0.0
    for v in value:
        peak = max(peak, v)
        mdd = min(mdd, v / peak - 1)
        mdd_usd = min(mdd_usd, v - peak)
    mr, ms = statistics.fmean(rets), statistics.fmean(srets)
    beta = sum((a - mr) * (b - ms) for a, b in zip(rets, srets)) / sum((b - ms) ** 2 for b in srets)
    return {
        "market_value": total,
        "largest_holding": top_tk,
        "largest_holding_share": weights[top_tk],
        "nvidia_share": weights.get(NVIDIA, 0.0),
        "largest_sector": top_sector,
        "largest_sector_share": sectors[top_sector] / total,
        "semiconductor_share": sectors.get(SEMIS, 0.0) / total,
        "effective_holdings": 1 / sum(w * w for w in weights.values()),
        "sector_shock_loss": shock * sectors[top_sector],
        "semiconductor_shock_loss": shock * sectors.get(SEMIS, 0.0),
        "volatility": statistics.stdev(rets) * math.sqrt(TRADING_DAYS),
        "max_drawdown": mdd,
        "max_drawdown_usd": mdd_usd,
        "var_95_1d_usd": -percentile_inc(rets, 0.05) * total,
        "beta_vs_spy": beta,
        "return_12m_backcast": value[-1] / value[0] - 1,
    }


def cap_weights(values: dict, cap: float, only: set | None = None) -> dict:
    """Target values with each holding (or each holding in `only`) at most `cap` of the total.

    The excess is spread over the holdings still below the cap in proportion to their value,
    and the rule is reapplied until nothing is above it. The total is unchanged.
    """
    if only is None and cap * len(values) < 1:
        raise ValueError(f"a {cap:.0%} cap cannot hold for {len(values)} holdings: they must sum to 100%")
    total = sum(values.values())
    target = dict(values)
    capped = set()
    while True:
        over = [t for t, v in target.items() if (only is None or t in only) and t not in capped
                and v > cap * total + 1e-9]
        if not over:
            return target
        for t in over:
            target[t] = cap * total
            capped.add(t)
        free = [t for t in target if t not in capped]
        excess = total - sum(target.values())
        base = sum(target[t] for t in free)
        for t in free:
            target[t] += excess * target[t] / base


def cap_group(values: dict, group: set, cap: float) -> dict:
    """Target values with the holdings in `group` together at most `cap` of the total.

    The group is scaled down as a whole (keeping its internal mix) and the freed value is
    spread over the other holdings in proportion to their value.
    """
    total = sum(values.values())
    in_group = sum(v for t, v in values.items() if t in group)
    if in_group <= cap * total:
        return dict(values)
    scale = cap * total / in_group
    others = total - in_group
    freed = in_group - cap * total
    return {t: (v * scale if t in group else v + freed * v / others) for t, v in values.items()}


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
    values = {h["ticker"]: h["units"] * h["price"] for h in holdings}
    semis = {h["ticker"] for h in holdings if h["sector"] == SEMIS}
    policies = {
        "Current portfolio": values,
        "NVIDIA capped at 25%": cap_weights(values, 0.25, only={NVIDIA}),
        "No holding above 20%": cap_weights(values, 0.20),
        "Semiconductors capped at 40%": cap_group(values, semis, 0.40),
    }
    out = {"assumptions": {"tax_rate_on_net_realised_gains": TAX_RATE, "trading_cost_share_of_value_traded": TRADING_COST,
                           "sector_shock": shock},
           "scenarios": {}}
    for name, target in policies.items():
        units = {h["ticker"]: target[h["ticker"]] / h["price"] for h in holdings}
        out["scenarios"][name] = {**profile(holdings, units, spy, shock), **trade_costs(holdings, target)}
    return out


def plot(result: dict, path: Path) -> None:
    """Money at risk and the one-off cost of each policy, side by side."""
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    names = list(result["scenarios"])
    s = result["scenarios"]
    series = [("Loss if semiconductors fell 30%", [s[n]["semiconductor_shock_loss"] for n in names], "#c0392b"),
              ("Worst 12-month drawdown", [-s[n]["max_drawdown_usd"] for n in names], "#dd8452"),
              ("1-day VaR (95%)", [s[n]["var_95_1d_usd"] for n in names], "#4c72b0"),
              ("One-off cost: tax + trading", [s[n]["total_cost"] for n in names], "#7f7f7f")]
    fig, ax = plt.subplots(figsize=(9.5, 4.6))
    width = 0.2
    for i, (label, vals, colour) in enumerate(series):
        xs = [k + (i - 1.5) * width for k in range(len(names))]
        ax.bar(xs, [v / 1000 for v in vals], width, label=label, color=colour)
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
    rows = [("market value", "market_value", "{:,.0f}"),
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
    print(f"{'':34s}" + "".join(f"{n[:22]:>24s}" for n in names))
    for label, key, fmt in rows:
        print(f"{label:34s}" + "".join(f"{fmt.format(result['scenarios'][n][key]):>24s}" for n in names))
    a = result["assumptions"]
    print(f"\nAssumptions: tax {a['tax_rate_on_net_realised_gains']:.0%} on net realised gains, "
          f"trading cost {a['trading_cost_share_of_value_traded']:.2%} of value traded, sector shock {a['sector_shock']:.0%}.")
    (ROOT / "outputs").mkdir(exist_ok=True)
    (ROOT / "outputs" / "scenarios.json").write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    (ROOT / "figures").mkdir(exist_ok=True)
    plot(result, ROOT / "figures" / "rebalancing.png")
    print("wrote outputs/scenarios.json and figures/rebalancing.png")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
