"""Write the four CSVs the Power BI report reads, from the workbook's stored values.

The report (StockPortfolio.pbip) reads powerbi-data/*.csv through the DataFolder parameter. Exporting from the
same cells validate_portfolio.py checks keeps the report, the workbook and the README on one set of figures.

Run: python export_powerbi_data.py   (after saving the workbook in Excel, so its cached values are current)
"""
import csv
import sys
from pathlib import Path

import validate_portfolio as vp

ROOT = Path(__file__).parent
OUT = ROOT / "powerbi-data"

HOLDING_FIRST_ROW, HOLDING_LAST_ROW = 7, 22      # Dashboard holdings table, one row per stock
HOLDING_COLS = {                                  # Dashboard column -> CSV column
    13: "Ticker", 3: "Industry", 4: "Units", 5: "CurrentPrice", 6: "MarketValue", 7: "TodaysChangeDollar",
    8: "TodaysChangePct", 9: "GainLoss", 11: "ReturnPct", 12: "WeightPct", 14: "CostBasis", 15: "AvgBuyPrice",
    16: "PriceToCost", 17: "AnnReturn", 18: "RiskScore", 19: "Signal",
}
MARKET_VALUE_COL = 6                                # Dashboard column F
SECTOR_FIRST_ROW = 30                             # Dashboard sector allocation, columns B to E
KPI_CELLS = {                                     # KPI column -> Dashboard cell
    "PortfolioValue": "B3", "CostBasis": "D3", "UnrealisedPnL": "E3", "TotalReturnPct": "F3",
    "AnnReturnIRR": "G3", "BestStock": "H3", "WorstStock": "I3", "HHIIndex": "J3", "DailyPnL": "M3",
    "SharpeRatio": "N3", "WinRatePct": "O3", "AvgReturnPct": "P3", "MaxDrawdownPct": "Q3", "PortfolioBeta": "R3",
    "TreynorRatio": "S3", "Top3ConcentrationPct": "E25", "SortinoRatio": "Q25", "CalmarRatio": "R25",
    "PortfolioGrade": "S25",
}


def clean(v):
    """Strip the arrow and symbol prefixes the dashboard adds for display (for example '▲ NVDA', '▶ HOLD')."""
    if isinstance(v, str):
        return v.lstrip("▲▼▶🔻🔺 ").strip()
    return v


def write(name, header, rows):
    """Write one CSV into powerbi-data/: the header row, then the data rows."""
    with open(OUT / name, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def main(path=vp.DEFAULT):
    wb = vp.load(Path(path))
    dash = wb["Dashboard"]
    OUT.mkdir(exist_ok=True)

    holdings = []
    for r in range(HOLDING_FIRST_ROW, HOLDING_LAST_ROW + 1):
        row = []
        for col in HOLDING_COLS:
            row.append(clean(dash.cell(r, col).value))
        ticker = row[0]    # Ticker is the first CSV column
        if ticker is None:
            sys.exit(f"Dashboard row {r} has no ticker in column M; the holdings table has moved")
        holdings.append(row)
    write("Holdings.csv", list(HOLDING_COLS.values()), holdings)

    sectors = []
    r = SECTOR_FIRST_ROW
    while dash.cell(r, 2).value not in (None, ""):
        sectors.append([dash.cell(r, 2).value, dash.cell(r, 3).value,
                        dash.cell(r, 4).value, dash.cell(r, 5).value])
        r += 1
    write("SectorAllocation.csv", ["Industry", "MarketValue", "WeightPct", "HHIContrib"], sectors)

    trades = vp.read_ledger(wb["Ledger"])
    ledger_rows = []
    for t in trades:
        ledger_rows.append([t["date"].isoformat(), t["kind"].title(), t["price"], t["units"],
                            t["amount"], t["ticker"]])
    write("Ledger.csv", ["Date", "TransactionType", "Price", "Units", "TransactionAmount", "Ticker"], ledger_rows)

    kpi_row = []
    for cell in KPI_CELLS.values():
        kpi_row.append(clean(dash[cell].value))
    write("KPIs.csv", list(KPI_CELLS), [kpi_row])

    market_values = []
    for r in range(HOLDING_FIRST_ROW, HOLDING_LAST_ROW + 1):
        market_values.append(clean(dash.cell(r, MARKET_VALUE_COL).value))
    mv = sum(market_values)
    if abs(mv - dash["B3"].value) > 0.01:
        sys.exit(f"holdings sum to {mv:,.2f} but the dashboard reports {dash['B3'].value:,.2f}")
    print(f"wrote {OUT.name}/: {len(holdings)} holdings, {len(sectors)} sectors, {len(trades)} trades; "
          f"market value ${mv:,.0f}")


if __name__ == "__main__":
    main(*sys.argv[1:])
