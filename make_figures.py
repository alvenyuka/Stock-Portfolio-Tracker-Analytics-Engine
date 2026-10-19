"""
make_figures.py: draw the README charts from the workbook's stored values.

Reads the cached results in Stock Portfolio.xlsx (the same values validate_portfolio.py
rebuilds and checks), so the charts match the validated figures. Opening the workbook in
Excel would refresh prices; reading it here does not.

Writes:
  figures/portfolio_vs_spy.png   12-month value of today's holdings against SPY, both indexed to 100
  figures/sector_exposure.png    market value by sector

Run: python make_figures.py   (needs matplotlib)
"""
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

from validate_portfolio import DEFAULT, SECTOR_FIRST, SECTOR_LAST, as_date, load, num  # noqa: E402

OUT = Path(__file__).parent / "figures"
BLUE, GREY, INK = "#2b6cb0", "#a0aec0", "#2d3748"


def main() -> None:
    """Draw the README charts from the workbook's cached values."""
    wb = load(DEFAULT)
    spark, ps, an = wb["Price History"], wb["Portfolio Series"], wb["Analytics"]

    # Row 2 of Price History holds the dates, starting in column B.
    dates = []
    for col in range(2, spark.max_column + 1):
        cell_value = spark.cell(2, col).value
        if num(cell_value) is not None:
            dates.append(as_date(cell_value))
    ndays = len(dates)

    # Portfolio Series: row 5 is the value of today's holdings, row 7 is SPY.
    value = []
    spy = []
    for col in range(2, 2 + ndays):
        value.append(num(ps.cell(5, col).value))
        spy.append(num(ps.cell(7, col).value))
    OUT.mkdir(exist_ok=True)

    # index both lines to 100 on the first day so they can share one axis
    portfolio_index = []
    spy_index = []
    for v in value:
        portfolio_index.append(v / value[0] * 100)
    for v in spy:
        spy_index.append(v / spy[0] * 100)

    fig, ax = plt.subplots(figsize=(8.5, 4.2))
    ax.plot(dates, portfolio_index, color=BLUE, lw=2,
            label=f"Current holdings ({value[-1] / value[0] - 1:+.1%})")
    ax.plot(dates, spy_index, color=GREY, lw=2, label=f"SPY ({spy[-1] / spy[0] - 1:+.1%})")
    ax.set_ylabel("Indexed to 100")
    ax.set_title(f"12 months to {dates[-1]:%d %b %Y}: current holdings against SPY", loc="left", fontsize=11, color=INK)
    ax.legend(frameon=False, loc="upper left")
    ax.spines[["top", "right"]].set_visible(False)
    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(OUT / "portfolio_vs_spy.png", dpi=150)
    plt.close(fig)

    # (sector name, market value), smallest first so the largest bar is at the top
    sectors = []
    for row in range(SECTOR_FIRST, SECTOR_LAST + 1):
        name = an.cell(row, 1).value
        market_value = num(an.cell(row, 2).value)
        if name and market_value is not None:
            sectors.append((name, market_value))
    sectors.sort(key=lambda sector: sector[1])

    names = []
    thousands = []
    values = []
    for name, market_value in sectors:
        names.append(name)
        thousands.append(market_value / 1000)
        values.append(market_value)
    total = sum(values)
    largest = values[-1]
    colours = []
    for market_value in values:
        colours.append(BLUE if market_value == largest else GREY)    # highlight the largest sector

    fig, ax = plt.subplots(figsize=(8.5, 4.6))
    bars = ax.barh(names, thousands, color=colours)
    for bar, market_value in zip(bars, values):
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2, f"{market_value / total:.1%}",
                va="center", fontsize=9)
    ax.set_xlabel("Market value, $ thousands")
    ax.set_title(f"Market value by sector (total ${total:,.0f})", loc="left", fontsize=11, color=INK)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig(OUT / "sector_exposure.png", dpi=150)
    plt.close(fig)
    print("wrote figures/portfolio_vs_spy.png and figures/sector_exposure.png")


if __name__ == "__main__":
    main()
