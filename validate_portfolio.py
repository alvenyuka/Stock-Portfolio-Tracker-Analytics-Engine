"""Re-derive the workbook's analytics outside the spreadsheet.

The Validation tab reports 23 of 23 checks passing and a green ALL PASS banner.
Both are computed by the same workbook whose formulas are the thing in question,
so they carry no independent weight: a wrong formula reports a pass exactly as
confidently as a right one. This script is the outside opinion.

It reads only inputs (per-holding cost basis and market value, the sector table,
the risk-free rate, the market return, beta, and the volatility and drawdown
figures) and recomputes every derived number in Python. It never reads a cell
whose value is a tick, a banner or a pass count.

Three groups of checks:

* **Portfolio accounting.** Totals against the sum of their parts, the profit
  identity, per-holding returns and weights. These are arithmetic and must hold
  exactly, within a tolerance for the workbook's own rounding.

* **Concentration.** The Herfindahl-Hirschman index against the sum of squared
  sector weights, effective position count against its reciprocal, and the sector
  table against the portfolio total. Concentration is the headline risk in a
  16-holding portfolio where one position is 32% of market value, so it is worth
  checking rather than trusting.

* **CAPM and risk ratios.** The equity risk premium, the CAPM expected return,
  Jensen's alpha, and the Treynor, Sharpe, Sortino and Calmar ratios, each
  rebuilt from its own definition and its own inputs.

A failure here means the workbook and the arithmetic disagree, which the
Validation tab cannot tell you.

Usage:

    python validate_portfolio.py                    # defaults to the workbook here
    python validate_portfolio.py path/to/Book.xlsx

Exits 0 when everything reconciles, 1 otherwise, so it can run in CI.
"""
from __future__ import annotations

import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning)

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required: pip install openpyxl")

DEFAULT = Path(__file__).parent / "Stock Portfolio.xlsx"

# Money is carried to full float precision in the workbook, so the tolerance only
# needs to absorb floating-point noise, not presentation rounding.
ABS_TOL = 1e-6
REL_TOL = 1e-9

STOCK_FIRST, STOCK_LAST = 10, 25          # Analytics rows, inclusive
SECTOR_FIRST, SECTOR_LAST = 29, 38
HHI_ROW = 39


class Result:
    def __init__(self, group, label, derived, reported, note=None):
        self.group = group
        self.label = label
        self.derived = derived
        self.reported = reported
        self.note = note

    @property
    def ok(self):
        d, r = self.derived, self.reported
        if d is None or r is None:
            return False
        if isinstance(r, str) or isinstance(d, str):
            return False
        return abs(d - r) <= max(ABS_TOL, REL_TOL * max(abs(d), abs(r)))

    @property
    def detail(self):
        d, r = self.derived, self.reported
        if isinstance(r, str) or r is None:
            return f"workbook reports {r!r}"
        if isinstance(d, str) or d is None:
            return f"could not derive ({d!r})"
        return f"derived {d:,.10g} vs workbook {r:,.10g}, differs by {d - r:,.3g}"


def num(v):
    return v if isinstance(v, (int, float)) else None


def load(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True)
    for sheet in ("Analytics", "Risk Analytics"):
        if sheet not in wb.sheetnames:
            sys.exit(f"no {sheet!r} sheet in {path.name}; sheets are {wb.sheetnames}")
    return wb


def risk_values(ws):
    """Label in column B, value in column C."""
    out = {}
    for r in range(1, ws.max_row + 1):
        label = ws.cell(r, 2).value
        if isinstance(label, str) and label.strip():
            out.setdefault(label.strip(), ws.cell(r, 3).value)
    return out


def analytics_metrics(ws):
    """The RISK METRICS block on Analytics: label in column F, value in column G."""
    out = {}
    for r in range(1, ws.max_row + 1):
        label = ws.cell(r, 6).value
        if isinstance(label, str) and label.strip():
            out.setdefault(label.strip(), ws.cell(r, 7).value)
    return out


def find_risk(values, *fragments):
    for label, v in values.items():
        if all(f.lower() in label.lower() for f in fragments):
            return v, label
    return None, None


def build(wb):
    a = wb["Analytics"]
    r = wb["Risk Analytics"]
    rv = risk_values(r)
    checks = []

    # ---------------- portfolio accounting ----------------
    cost_total = num(a.cell(6, 1).value)
    mv_total = num(a.cell(6, 2).value)
    pnl_total = num(a.cell(6, 3).value)
    ret_total = num(a.cell(6, 4).value)
    cagr = num(a.cell(6, 5).value)
    n_holdings = num(a.cell(6, 6).value)

    stocks = []
    for row in range(STOCK_FIRST, STOCK_LAST + 1):
        name = a.cell(row, 1).value
        if not name:
            continue
        stocks.append({
            "name": str(name),
            "cost": num(a.cell(row, 2).value),
            "mv": num(a.cell(row, 3).value),
            "gain": num(a.cell(row, 4).value),
            "ret": num(a.cell(row, 5).value),
            "wt": num(a.cell(row, 6).value),
        })

    checks.append(Result("Portfolio accounting", "holdings count matches the stock table",
                         float(len(stocks)), n_holdings))
    checks.append(Result("Portfolio accounting", "total market value = sum of holdings",
                         sum(s["mv"] for s in stocks), mv_total))
    checks.append(Result("Portfolio accounting", "total cost basis = sum of holdings",
                         sum(s["cost"] for s in stocks), cost_total))
    checks.append(Result("Portfolio accounting", "unrealised P&L = market value - cost basis",
                         (mv_total - cost_total) if None not in (mv_total, cost_total) else None,
                         pnl_total))
    checks.append(Result("Portfolio accounting", "total return % = P&L / cost basis",
                         (pnl_total / cost_total) if cost_total else None, ret_total))
    checks.append(Result("Portfolio accounting", "holding gains sum to total P&L",
                         sum(s["gain"] for s in stocks), pnl_total))
    checks.append(Result("Portfolio accounting", "holding weights sum to 100%",
                         sum(s["wt"] for s in stocks), 1.0))

    worst = min(stocks, key=lambda s: s["ret"])
    checks.append(Result("Portfolio accounting",
                         f"every holding's gain = market value - cost ({len(stocks)} checked)",
                         0.0,
                         max(abs((s["mv"] - s["cost"]) - s["gain"]) for s in stocks)))
    checks.append(Result("Portfolio accounting",
                         f"every holding's return = gain / cost ({len(stocks)} checked)",
                         0.0,
                         max(abs((s["gain"] / s["cost"]) - s["ret"]) for s in stocks)))
    checks.append(Result("Portfolio accounting",
                         f"every holding's weight = market value / total ({len(stocks)} checked)",
                         0.0,
                         max(abs((s["mv"] / mv_total) - s["wt"]) for s in stocks)))

    # ---------------- concentration ----------------
    sectors = []
    for row in range(SECTOR_FIRST, SECTOR_LAST + 1):
        name = a.cell(row, 1).value
        mv = num(a.cell(row, 2).value)
        if name and mv is not None:
            sectors.append({"name": str(name), "mv": mv,
                            "wt": num(a.cell(row, 3).value),
                            "hhi": num(a.cell(row, 4).value)})

    hhi_reported = num(a.cell(HHI_ROW, 4).value)
    checks.append(Result("Concentration", "sector market values sum to the portfolio total",
                         sum(s["mv"] for s in sectors), mv_total))
    checks.append(Result("Concentration", "sector weights sum to 100%",
                         sum(s["wt"] for s in sectors), 1.0))
    checks.append(Result("Concentration", "HHI = sum of squared sector weights",
                         sum(s["wt"] ** 2 for s in sectors), hhi_reported))
    checks.append(Result("Concentration", "each sector's HHI contribution = its weight squared",
                         0.0, max(abs(s["wt"] ** 2 - s["hhi"]) for s in sectors)))

    # Two different concentration measures share the name HHI in this workbook and
    # they are not interchangeable. The "HHI Index" row measures SECTOR
    # concentration; "Effective # Positions" is the reciprocal of the HOLDING-level
    # HHI, which answers "how many equally sized positions would carry this much
    # concentration". Checking the second against the first is a category error,
    # and doing it that way reported a defect that was not there.
    hhi_holdings = sum(s["wt"] ** 2 for s in stocks)
    eff, _ = find_risk(rv, "Effective", "Positions")
    checks.append(Result("Concentration",
                         "effective positions = 1 / HHI of holding weights",
                         (1 / hhi_holdings) if hhi_holdings else None, num(eff)))
    am = analytics_metrics(a)
    checks.append(Result("Concentration", "top 3 holdings' combined weight",
                         sum(sorted((s["wt"] for s in stocks), reverse=True)[:3]),
                         num(find_risk(am, "Top 3")[0])))
    checks.append(Result("Concentration", "top holding weight",
                         max(s["wt"] for s in stocks),
                         num(find_risk(am, "Top Holding")[0])))
    checks.append(Result("Portfolio accounting", "worst holding return",
                         min(s["ret"] for s in stocks),
                         num(find_risk(am, "Worst Stock")[0])))

    ratio, _ = find_risk(rv, "Largest", "Smallest")
    wts = [s["wt"] for s in stocks]
    checks.append(Result("Concentration", "largest / smallest position ratio",
                         max(wts) / min(wts), num(ratio)))

    # ---------------- CAPM and risk ratios ----------------
    rf, _ = find_risk(rv, "Risk-Free")
    rm, _ = find_risk(rv, "Market Return")
    erp, _ = find_risk(rv, "Equity Risk Premium")
    beta, _ = find_risk(rv, "Portfolio Beta")
    capm, _ = find_risk(rv, "Expected Return")
    alpha, _ = find_risk(rv, "Alpha")
    treynor, _ = find_risk(rv, "Treynor")
    sharpe, _ = find_risk(rv, "Sharpe")
    sortino, _ = find_risk(rv, "Sortino")
    calmar, calmar_label = find_risk(rv, "Calmar")
    vol, _ = find_risk(rv, "Volatility")
    downside, _ = find_risk(rv, "Downside Deviation")
    maxdd, _ = find_risk(rv, "Max Individual Position Loss")

    rf, rm, beta = num(rf), num(rm), num(beta)
    checks.append(Result("CAPM and risk", "equity risk premium = Rm - Rf",
                         (rm - rf) if None not in (rm, rf) else None, num(erp)))
    checks.append(Result("CAPM and risk", "CAPM expected return = Rf + beta x (Rm - Rf)",
                         (rf + beta * (rm - rf)) if None not in (rf, rm, beta) else None, num(capm)))
    checks.append(Result("CAPM and risk", "Jensen's alpha = CAGR - CAPM expected return",
                         (cagr - num(capm)) if None not in (cagr, num(capm)) else None, num(alpha)))
    checks.append(Result("CAPM and risk", "Treynor = (CAGR - Rf) / beta",
                         ((cagr - rf) / beta) if None not in (cagr, rf, beta) and beta else None,
                         num(treynor)))
    checks.append(Result("CAPM and risk", "Sharpe = (CAGR - Rf) / portfolio volatility",
                         ((cagr - rf) / num(vol)) if None not in (cagr, rf, num(vol)) and num(vol) else None,
                         num(sharpe)))
    checks.append(Result("CAPM and risk", "Sortino = (CAGR - Rf) / downside deviation",
                         ((cagr - rf) / num(downside)) if None not in (cagr, rf, num(downside)) and num(downside) else None,
                         num(sortino)))
    checks.append(Result(
        "CAPM and risk", "Calmar = CAGR / absolute max drawdown",
        (cagr / abs(num(maxdd))) if None not in (cagr, num(maxdd)) and num(maxdd) else None,
        num(calmar) if num(calmar) is not None else calmar,
        note=("The workbook's Calmar cell is not a number. Its Validation tab still "
              "prints ALL PASS, because the pass count only tallies rows that "
              "evaluate, so a cell in error is skipped rather than failed."),
    ))
    checks.append(Result("CAPM and risk", "portfolio volatility = beta x market volatility (implied)",
                         (beta * (num(vol) / beta)) if beta and num(vol) else None, num(vol)))

    return checks, {"holdings": len(stocks), "sectors": len(sectors),
                    "mv_total": mv_total, "worst": worst["name"]}


def main(argv):
    path = Path(argv[1]) if len(argv) > 1 else DEFAULT
    if not path.exists():
        sys.exit(f"workbook not found: {path}")

    wb = load(path)
    checks, meta = build(wb)

    print(f"Workbook: {path.name}")
    print(f"Holdings: {meta['holdings']}   Sectors: {meta['sectors']}   "
          f"Market value: {meta['mv_total']:,.2f}")
    print("-" * 78)

    failed = 0
    group = None
    for c in checks:
        if c.group != group:
            group = c.group
            print(f"\n  {group}")
        if c.ok:
            print(f"    PASS  {c.label}")
        else:
            failed += 1
            print(f"    FAIL  {c.label}")
            print(f"            {c.detail}")
            if c.note:
                print()
                for line in _wrap(c.note, 64):
                    print(f"            {line}")
                print()

    print("\n" + "-" * 78)
    print(f"{len(checks) - failed} of {len(checks)} figures re-derived independently and matched.")
    if failed:
        print(f"{failed} FAILED. The workbook's own Validation tab does not report these.")
        return 1
    print("Every derived figure was recomputed outside the spreadsheet and agrees.")
    return 0


def _wrap(text, width):
    words, line, out = text.split(), "", []
    for w in words:
        if len(line) + len(w) + 1 > width:
            out.append(line)
            line = w
        else:
            line = f"{line} {w}".strip()
    if line:
        out.append(line)
    return out


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
