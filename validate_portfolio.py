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

What this script cannot tell you, stated here because an earlier version of it
read as though it could. Every check above compares the workbook against its own
definitions. None of them can find a figure that is computed correctly and means
something other than its label says. Four such defects are known, so they are
reported explicitly, under KNOWN DEFECT, with the numbers measured rather than
described. They do not fail the run, because they are not arithmetic errors and
because correcting them rewrites published performance figures, which is the
owner's decision and not this script's. The one that matters most is the first:
the Dashboard adds sell transactions to positions instead of netting them, so
every figure in the workbook, including the ones re-derived and matched above,
is built on the wrong number of units.

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

# The workbook's market-volatility assumption, hardcoded in Risk Analytics C25,
# C66 and C68. Unlike Rf and Rm it is not a labelled input cell, which is why it
# is named here rather than back-solved from the figure it produces.
MARKET_VOLATILITY = 0.155

# The workbook annualises over a fixed 7-year horizon, in Analytics E6 and in
# Dashboard Q7, for every holding alike. The ledger runs 2019 to 2025 and
# individual lots were bought as late as 2024, so this is a lump-sum equivalent
# rather than a money-weighted return. Named here so the assumption is visible.
CAGR_YEARS = 7

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


class Defect:
    """A finding the arithmetic checks cannot express.

    Every Result above asks whether the workbook computes what its own formula
    says. A Defect asks whether the formula computes what its label claims, and
    the answer is fixed and known, so there is nothing to compare. It is printed
    with its measured numbers and does not affect the exit code: correcting any
    of these changes published performance figures, and that is a decision for
    the owner of the model rather than for a validator.
    """

    def __init__(self, label, detail):
        self.label = label
        self.detail = detail


def ascii_label(text):
    """The workbook's row labels carry Greek letters and en dashes. Printing them
    verbatim raises on a console using a legacy code page, so a defect report
    would crash on exactly the machine the workbook is edited on."""
    if not isinstance(text, str):
        return str(text)
    return text.encode("ascii", "replace").decode("ascii")


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
    # CAGR is the numerator of Jensen's alpha, Treynor, Sharpe, Sortino, Calmar,
    # the VaR series and M-squared, and it was the one figure this script read
    # without rederiving. The horizon is hardcoded at 7 years in the workbook.
    checks.append(Result("Portfolio accounting",
                         f"CAGR = (1 + total return) ^ (1/{CAGR_YEARS}) - 1",
                         ((1 + ret_total) ** (1 / CAGR_YEARS) - 1) if ret_total is not None else None,
                         cagr))

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
    # The previous version of this check derived `beta * (vol / beta)`, which is
    # `vol` for any non-zero beta, so it compared a cell against itself and
    # could not fail for any volatility whatsoever. The assumption it was meant
    # to pin is the 0.155 market volatility, which is hardcoded three times in
    # the workbook and, unlike Rf and Rm, is not a labelled input cell.
    checks.append(Result("CAPM and risk",
                         f"portfolio volatility = beta x {MARKET_VOLATILITY} market volatility",
                         (beta * MARKET_VOLATILITY) if beta is not None else None,
                         num(vol)))

    return checks, {"holdings": len(stocks), "sectors": len(sectors),
                    "mv_total": mv_total, "worst": worst["name"]}


def build_defects(wb):
    """The four things the checks above are structurally unable to catch."""
    a = wb["Analytics"]
    r = wb["Risk Analytics"]
    rv = risk_values(r)
    defects = []

    # ---- 1. the ledger does not reconcile to the holdings --------------------
    if "Ledger" in wb.sheetnames:
        ledger = wb["Ledger"]
        gross = net = 0.0
        buys = sells = 0
        for row in range(1, ledger.max_row + 1):
            kind = ledger.cell(row, 4).value
            units = num(ledger.cell(row, 6).value)
            if not isinstance(kind, str) or units is None:
                continue
            kind = kind.strip().lower()
            if kind == "buy":
                buys += 1
                gross += units
                net += units
            elif kind == "sell":
                sells += 1
                gross += units
                net -= units
        dashboard_units = sum(
            num(a.cell(row, 2).value) or 0.0 for row in range(STOCK_FIRST, STOCK_LAST + 1)
        )
        if buys and sells:
            defects.append(Defect(
                "the holdings do not net the ledger's sales",
                f"{buys} buys and {sells} sells, every one of them recorded with "
                f"positive units, are added together by a plain SUMIF on "
                f"Dashboard!D7. Gross units across the ledger: {gross:,.2f}. Net "
                f"of sales: {net:,.2f}. A sale therefore increases both the "
                f"position and the cost basis. Market value, weights, returns, "
                f"CAGR and every risk ratio on this sheet inherit it, including "
                f"the ones reported as matching above, which check the workbook "
                f"against itself and not against the ledger."
            ))

    # ---- 2. tracking error is a standard deviation minus a return -----------
    # "Tracking Error" also appears inside the Information Ratio's own label,
    # which comes first on the sheet, so match on the row that starts with it.
    te, te_label = find_risk(rv, "tracking error (")
    ir, _ = find_risk(rv, "information ratio")
    alpha, _ = find_risk(rv, "jensen")
    if num(te) is not None and num(te) < 0:
        defects.append(Defect(
            "tracking error is negative, and the information ratio inherits its sign",
            f"{ascii_label(te_label)!r} is {num(te):,.6f}. It is computed as the cross-sectional "
            f"standard deviation of the holdings' returns minus the portfolio's own "
            f"return, which subtracts a return level from a dispersion measure. A "
            f"standard deviation cannot be negative. The information ratio divides "
            f"Jensen's alpha ({num(alpha):,.6f}) by it and reports "
            f"{num(ir):,.6f}, positive, on a portfolio the same sheet marks as "
            f"having negative alpha, negative M-squared and a below-CML position. "
            f"Tracking error needs a benchmark return series, which the workbook "
            f"does not have."
        ))

    # ---- 3. max drawdown is the worst holding's total return ----------------
    maxdd, dd_label = find_risk(rv, "max", "position")
    calmar, _ = find_risk(rv, "calmar")
    if num(maxdd) is not None:
        defects.append(Defect(
            "the Calmar denominator is not a drawdown",
            f"{ascii_label(dd_label)!r} is {num(maxdd):,.6f}, the worst single holding's total "
            f"return since inception. A maximum drawdown is a peak-to-trough "
            f"decline of the portfolio's value over time, and no such series is "
            f"computed anywhere in the workbook. Calmar "
            f"({num(calmar) if num(calmar) is not None else calmar}) is CAGR "
            f"divided by it, and is then scored against benchmark bands that "
            f"belong to a real Calmar ratio."
        ))

    # ---- 4. volatility contains no idiosyncratic risk -----------------------
    vol, vol_label = find_risk(rv, "volatility")
    if num(vol) is not None:
        defects.append(Defect(
            "Sharpe and the VaR series use systematic volatility only",
            f"{ascii_label(vol_label)!r} is beta times a hardcoded {MARKET_VOLATILITY} market "
            f"volatility, so it carries no idiosyncratic risk at all. For a "
            f"16-holding book whose largest position is about a third of market "
            f"value, total volatility is materially higher, so Sharpe is "
            f"overstated and the VaR and CVaR tail losses are understated. It "
            f"also makes Sharpe algebraically identical to Treynor divided by "
            f"{MARKET_VOLATILITY}, so the two are not independent measures."
        ))

    return defects


def main(argv):
    path = Path(argv[1]) if len(argv) > 1 else DEFAULT
    if not path.exists():
        sys.exit(f"workbook not found: {path}")

    wb = load(path)
    checks, meta = build(wb)
    defects = build_defects(wb)

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

    if defects:
        print("\n  Known defects, reported but not failed")
        for d in defects:
            print(f"    KNOWN DEFECT  {d.label}")
            print()
            for line in _wrap(d.detail, 64):
                print(f"            {line}")
            print()

    print("\n" + "-" * 78)
    print(f"{len(checks) - failed} of {len(checks)} figures re-derived independently and matched.")
    if defects:
        print(f"{len(defects)} known defect(s) reported above. These are not "
              f"arithmetic errors, so nothing here fails on them, and they are "
              f"not fixed here because correcting them rewrites published "
              f"performance figures.")
    if failed:
        print(f"{failed} FAILED. The workbook's own Validation tab does not report these.")
        return 1
    print("Every derived figure was recomputed outside the spreadsheet and agrees "
          "with the workbook's own definitions. That is a consistency result, not "
          "a statement that the definitions are the right ones. See the known "
          "defects above.")
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
