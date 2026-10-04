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

    dates = [as_date(v) for v in (spark.cell(2, c).value for c in range(2, spark.max_column + 1)) if num(v) is not None]
    n = len(dates)
    value = [num(ps.cell(5, c).value) for c in range(2, 2 + n)]
    spy = [num(ps.cell(7, c).value) for c in range(2, 2 + n)]
    OUT.mkdir(exist_ok=True)

    fig, ax = plt.subplots(figsize=(8.5, 4.2))
    ax.plot(dates, [v / value[0] * 100 for v in value], color=BLUE, lw=2,
            label=f"Current holdings ({value[-1] / value[0] - 1:+.1%})")
    ax.plot(dates, [v / spy[0] * 100 for v in spy], color=GREY, lw=2, label=f"SPY ({spy[-1] / spy[0] - 1:+.1%})")
    ax.set_ylabel("Indexed to 100")
    ax.set_title(f"12 months to {dates[-1]:%d %b %Y}: current holdings against SPY", loc="left", fontsize=11, color=INK)
    ax.legend(frameon=False, loc="upper left")
    ax.spines[["top", "right"]].set_visible(False)
    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(OUT / "portfolio_vs_spy.png", dpi=150)
    plt.close(fig)

    sectors = [(an.cell(r, 1).value, num(an.cell(r, 2).value)) for r in range(SECTOR_FIRST, SECTOR_LAST + 1)
               if an.cell(r, 1).value and num(an.cell(r, 2).value) is not None]
    sectors.sort(key=lambda t: t[1])
    total = sum(v for _, v in sectors)
    fig, ax = plt.subplots(figsize=(8.5, 4.6))
    bars = ax.barh([s for s, _ in sectors], [v / 1000 for _, v in sectors],
                   color=[BLUE if v == sectors[-1][1] else GREY for _, v in sectors])
    for bar, (_, v) in zip(bars, sectors):
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2, f"{v / total:.1%}", va="center", fontsize=9)
    ax.set_xlabel("Market value, $ thousands")
    ax.set_title(f"Market value by sector (total ${total:,.0f})", loc="left", fontsize=11, color=INK)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig(OUT / "sector_exposure.png", dpi=150)
    plt.close(fig)
    print("wrote figures/portfolio_vs_spy.png and figures/sector_exposure.png")


if __name__ == "__main__":
    main()
