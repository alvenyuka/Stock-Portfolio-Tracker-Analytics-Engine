"""Re-derive the workbook's figures from its raw inputs, outside the spreadsheet.

The Validation tab reports its own checks passing. Those checks are computed by
the same workbook whose formulas are in question, so they carry no independent
weight. This script is the outside opinion.

It reads only inputs:

* the Ledger: date, ticker, buy or sell, price, units and amount of every trade;
* each holding's current price, and the Stocks data type betas;
* the 12 months of daily closes on the Price History sheet, and the SPY closes;
* the risk-free rate, the market return and the valuation date (one input cell,
  named ValuationDate, that every IRR and price window ends on).

From those it rebuilds holdings net of sales, cost basis at average purchase
cost, realised and unrealised P&L, money-weighted returns (XIRR), the 12-month
back-cast of the current holdings, every risk ratio and the money-at-risk block,
then compares each with the figure the workbook reports. It never reads a tick, a
banner or a pass count.

Cost basis follows the workbook: average cost over all purchases of a stock. A
running (moving) average, which only uses purchases made before each sale, is
printed alongside as an advisory, because the two differ for every holding that
was bought again after a sale.

Usage:

    python validate_portfolio.py                    # defaults to the workbook here
    python validate_portfolio.py path/to/Book.xlsx

Exits 0 when everything reconciles, 1 otherwise, so it can run in CI.
"""
from __future__ import annotations

import datetime as dt
import math
import statistics
import sys
import warnings
import zipfile
from collections import defaultdict
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required: pip install openpyxl")

DEFAULT = Path(__file__).parent / "Stock Portfolio.xlsx"
SHEETS = ("Ledger", "Dashboard", "Analytics", "Risk Analytics", "Price History", "Portfolio Series")

# Money is carried at full float precision, so the tolerance only absorbs
# floating-point noise. XIRR is iterative in Excel, so it gets its own tolerance.
ABS_TOL = 1e-6
REL_TOL = 1e-9
IRR_TOL = 1e-6
TRADING_DAYS = 252

LEDGER_FIRST_ROW = 3
DASH_FIRST, DASH_LAST = 7, 22          # Dashboard holdings rows
STOCK_FIRST, STOCK_LAST = 10, 25       # Analytics holdings rows
SECTOR_FIRST, SECTOR_LAST = 29, 38
HHI_ROW = 39
SPARK_FIRST = 3                        # Price History rows 3..18, same order as the Dashboard

# Risk Analytics, column C
RA = {"rf": "C6", "rm": "C7", "erp": "C8", "beta": "C9", "capm": "C10", "alpha": "C11",
      "treynor": "C15", "ir": "C16", "te": "C17", "sharpe": "C18", "sortino": "C19",
      "calmar": "C20", "mean": "C24", "vol": "C25", "var95": "C26", "var99": "C27",
      "cvar": "C28", "downside": "C30", "hvar": "C31", "eff_pos": "C34", "spread": "C35"}
# Portfolio Series, column B
PS = {"covered": "B10", "share": "B11", "ndays": "B12", "ret12": "B13", "vol": "B14",
      "downside": "B15", "mdd": "B16", "hvar": "B17", "spy12": "B18", "beta_spy": "B19",
      "te": "B20", "active": "B21", "ir": "B22", "spy_vol": "B23", "asof": "B24"}
# Risk Analytics: capital market line and money at risk
CML = {"cml": "C66", "spread": "C67", "m2": "C68"}
MONEY = {"shock": "C72", "top_sector": "C73", "loss": "C74", "mdd_usd": "C75", "hvar_usd": "C76"}


class Result:
    """One figure re-derived from the ledger or prices, compared with the figure the workbook reports."""

    def __init__(self, group, label, derived, reported, tol=None, note=None):
        self.group = group
        self.label = label
        self.derived = derived
        self.reported = reported
        self.tol = tol
        self.note = note

    @property
    def ok(self):
        """True when both are numbers and agree within `tol`, or within the default absolute and relative tolerance."""
        derived, reported = self.derived, self.reported
        if not isinstance(derived, (int, float)) or not isinstance(reported, (int, float)):
            return False
        gap = abs(derived - reported)
        if self.tol is not None:
            return gap <= self.tol
        return gap <= max(ABS_TOL, REL_TOL * max(abs(derived), abs(reported)))

    @property
    def detail(self):
        """One-line description of the comparison, for the report."""
        derived, reported = self.derived, self.reported
        if not isinstance(reported, (int, float)):
            return f"workbook reports {reported!r}, which is not a number"
        if not isinstance(derived, (int, float)):
            return f"could not derive ({derived!r})"
        return f"derived {derived:,.10g} vs workbook {reported:,.10g}, differs by {derived - reported:,.3g}"


def num(v):
    """The value if it is a real number (not a bool), else None."""
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def as_date(v):
    """A date from a datetime, a date, or an Excel serial number; None otherwise."""
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)):
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    return None


def xirr(flows):
    """Excel's XIRR: the rate r solving sum(amount / (1 + r) ^ (years since the first flow)) = 0.

    Found by bisection between -99.99% and 1000%; returns None if no rate in that range solves it.
    """
    flows = sorted(flows)
    first_date = flows[0][0]
    years = []
    amounts = []
    for date, amount in flows:
        years.append((date - first_date).days / 365.0)
        amounts.append(amount)

    def net_present_value(rate):
        present_values = []
        for amount, t in zip(amounts, years):
            present_values.append(amount / (1 + rate) ** t)
        return sum(present_values)

    low, high = -0.9999, 10.0
    npv_low, npv_high = net_present_value(low), net_present_value(high)
    if npv_low * npv_high > 0:
        return None
    for _ in range(300):
        mid = (low + high) / 2
        npv_mid = net_present_value(mid)
        if npv_low * npv_mid <= 0:
            high, npv_high = mid, npv_mid
        else:
            low, npv_low = mid, npv_mid
    return (low + high) / 2


def percentile_inc(xs, p):
    """Excel PERCENTILE.INC: linear interpolation between order statistics."""
    ordered = sorted(xs)
    position = p * (len(ordered) - 1)
    lower = math.floor(position)
    upper = min(lower + 1, len(ordered) - 1)
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


def load(path: Path):
    """Open the workbook with cached values and stop with a message if a required sheet is missing."""
    wb = openpyxl.load_workbook(path, data_only=True)
    missing = [s for s in SHEETS if s not in wb.sheetnames]
    if missing:
        sys.exit(f"missing sheet(s) {missing} in {path.name}; sheets are {wb.sheetnames}")
    return wb


def read_ledger(ws):
    """Every buy and sell row of the Ledger sheet as a dict of date, kind, price, units, amount and ticker."""
    trades = []
    for row in range(LEDGER_FIRST_ROW, ws.max_row + 1):
        kind = ws.cell(row, 4).value
        if not isinstance(kind, str) or kind.strip().lower() not in ("buy", "sell"):
            continue    # skips blank rows, headers and anything that is not a trade
        trades.append({"date": as_date(ws.cell(row, 2).value),
                       "kind": kind.strip().lower(),
                       "price": num(ws.cell(row, 5).value),
                       "units": num(ws.cell(row, 6).value),
                       "amount": num(ws.cell(row, 7).value),
                       "ticker": ws.cell(row, 8).value})
    return trades


def new_position():
    """A blank per-ticker record to add trades into."""
    return {"units_bought": 0.0, "amount_bought": 0.0,
            "units_sold": 0.0, "amount_sold": 0.0, "flows": []}


def positions(trades):
    """Per-ticker totals from the trades: units and amounts bought and sold, net units, pooled average cost, cost basis, realised gain and cash flows."""
    by_ticker = defaultdict(new_position)

    for trade in trades:
        position = by_ticker[trade["ticker"]]
        if trade["kind"] == "buy":
            position["units_bought"] += trade["units"]
            position["amount_bought"] += trade["amount"]
            # money leaving the investor's pocket is negative
            position["flows"].append((trade["date"], -trade["amount"]))
        else:
            position["units_sold"] += trade["units"]
            position["amount_sold"] += trade["amount"]
            position["flows"].append((trade["date"], trade["amount"]))

    for position in by_ticker.values():
        position["net_units"] = position["units_bought"] - position["units_sold"]
        if position["units_bought"]:
            position["avg_cost"] = position["amount_bought"] / position["units_bought"]
        else:
            position["avg_cost"] = 0.0
        position["cost_basis"] = position["net_units"] * position["avg_cost"]
        position["realised_pnl"] = position["amount_sold"] - position["units_sold"] * position["avg_cost"]
    return by_ticker


def running_average(trades):
    """Average cost that only uses purchases made before each sale.

    Returns, per ticker, the units still held, their cost and the realised P&L so far.
    """
    by_ticker = defaultdict(lambda: {"units": 0.0, "cost": 0.0, "realised": 0.0})

    for trade in sorted(trades, key=lambda t: t["date"]):
        state = by_ticker[trade["ticker"]]
        if trade["kind"] == "buy":
            state["units"] += trade["units"]
            state["cost"] += trade["amount"]
        else:
            avg_cost_so_far = state["cost"] / state["units"]
            cost_of_units_sold = trade["units"] * avg_cost_so_far
            state["realised"] += trade["amount"] - cost_of_units_sold
            state["cost"] -= cost_of_units_sold
            state["units"] -= trade["units"]
    return by_ticker


def uses_today(path):
    """Cells whose formula still calls TODAY() (the valuation-date cell excepted)."""
    hits = 0
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                hits += z.read(name).decode("utf-8", "ignore").count("TODAY()")
    return hits


def add_check(checks, group, label, derived, reported, tol=None, note=None):
    """Append one derived-versus-reported comparison to the list of checks."""
    checks.append(Result(group, label, derived, reported, tol=tol, note=note))


def add_gap_check(checks, group, label, pairs, tol=None):
    """Append one check that covers a whole list of (derived, reported) pairs.

    The reported figure is the largest absolute gap between any pair and the derived
    figure is 0, so the check passes only when every pair agrees. If any value is
    missing, the gap is the text "missing value", which always fails.
    """
    gaps = []
    for derived, reported in pairs:
        if num(derived) is not None and num(reported) is not None:
            gaps.append(abs(derived - reported))
    if len(gaps) == len(pairs):
        largest_gap = max(gaps)
    else:
        largest_gap = "missing value"
    add_check(checks, group, label, 0.0, largest_gap, tol=tol)


# ---------------------------------------------------------------------------
# One function per check group. Each adds its checks to the list in the order
# the report prints them, and returns the figures the later groups need.
# ---------------------------------------------------------------------------

def check_holdings(checks, wb, trades, pos, asof):
    """Group 1: units, cost basis, market value, P&L, weight and XIRR of each holding.

    Returns the holdings (one dict per Dashboard row, with our figures and the
    workbook's) and the total market value.
    """
    group = "Holdings from the ledger"
    dash, an = wb["Dashboard"], wb["Analytics"]

    amount_pairs = []
    for t in trades:
        amount_pairs.append((t["price"] * t["units"], t["amount"]))
    add_gap_check(checks, group, f"every trade amount = price x units ({len(trades)} trades)", amount_pairs)

    # The Dashboard and Analytics sheets list the holdings in the same order, so row i of one is row i of the other.
    holdings = []
    for offset in range(DASH_LAST - DASH_FIRST + 1):
        dash_row = DASH_FIRST + offset
        analytics_row = STOCK_FIRST + offset
        ticker = dash.cell(dash_row, 13).value
        if ticker not in pos:
            add_check(checks, group, f"holding {ticker!r} on the Dashboard appears in the ledger", None, 1.0)
            continue
        position = pos[ticker]
        price = num(dash.cell(dash_row, 5).value)
        if price is None or not position["cost_basis"]:
            add_check(checks, group, f"holding {ticker!r} has a current price and a cost basis", None, "missing")
            continue
        holdings.append({
            "ticker": ticker,
            "position": position,
            "price": price,
            "market_value": position["net_units"] * price,
            "dashboard_units": num(dash.cell(dash_row, 4).value),
            "workbook_cost_basis": num(an.cell(analytics_row, 2).value),
            "workbook_market_value": num(an.cell(analytics_row, 3).value),
            "workbook_unrealised": num(an.cell(analytics_row, 4).value),
            "workbook_return": num(an.cell(analytics_row, 5).value),
            "workbook_weight": num(an.cell(analytics_row, 6).value),
            "workbook_irr": num(an.cell(analytics_row, 8).value),
            "workbook_realised": num(an.cell(analytics_row, 9).value),
        })

    market_values = []
    for h in holdings:
        market_values.append(h["market_value"])
    mv_total = sum(market_values)

    units_pairs = []
    market_value_pairs = []
    cost_pairs = []
    unrealised_pairs = []
    realised_pairs = []
    return_pairs = []
    weight_pairs = []
    irr_pairs = []
    lowest_units = holdings[0]["position"]["net_units"]
    for h in holdings:
        position = h["position"]
        unrealised = h["market_value"] - position["cost_basis"]
        units_pairs.append((position["net_units"], h["dashboard_units"]))
        market_value_pairs.append((h["market_value"], h["workbook_market_value"]))
        cost_pairs.append((position["cost_basis"], h["workbook_cost_basis"]))
        unrealised_pairs.append((unrealised, h["workbook_unrealised"]))
        realised_pairs.append((position["realised_pnl"], h["workbook_realised"]))
        return_pairs.append((unrealised / position["cost_basis"], h["workbook_return"]))
        weight_pairs.append((h["market_value"] / mv_total, h["workbook_weight"]))
        # today's market value counts as a final cash inflow
        holding_irr = xirr(position["flows"] + [(asof, h["market_value"])])
        irr_pairs.append((holding_irr, h["workbook_irr"]))
        lowest_units = min(lowest_units, position["net_units"])

    add_gap_check(checks, group, f"every holding's units = bought - sold ({len(holdings)} holdings)", units_pairs)
    add_check(checks, group, "no holding has negative units", 0.0, max(0.0, -lowest_units))
    add_gap_check(checks, group, "every holding's market value = net units x price", market_value_pairs)
    add_gap_check(checks, group, "every holding's cost basis = net units x average purchase cost", cost_pairs)
    add_gap_check(checks, group, "every holding's unrealised P&L = market value - cost basis", unrealised_pairs)
    add_gap_check(checks, group, "every holding's realised P&L = proceeds - units sold x average cost", realised_pairs)
    add_gap_check(checks, group, "every holding's return = unrealised P&L / cost basis", return_pairs)
    add_gap_check(checks, group, "every holding's weight = market value / total", weight_pairs)
    add_gap_check(checks, group, "every holding's annual return = XIRR of its trades plus today's value",
                  irr_pairs, tol=IRR_TOL)
    return holdings, mv_total


def check_totals(checks, wb, trades, holdings, mv_total, asof):
    """Group 2: portfolio-level value, cost, P&L, total return and XIRR.

    Returns the figures the report reuses.
    """
    group = "Portfolio totals"
    dash, an = wb["Dashboard"], wb["Analytics"]

    buy_amounts = []
    sell_amounts = []
    all_flows = []
    for t in trades:
        if t["kind"] == "buy":
            buy_amounts.append(t["amount"])
            all_flows.append((t["date"], -t["amount"]))
        else:
            sell_amounts.append(t["amount"])
            all_flows.append((t["date"], t["amount"]))
    buys = sum(buy_amounts)
    sells = sum(sell_amounts)

    cost_bases = []
    realised_pnls = []
    for h in holdings:
        cost_bases.append(h["position"]["cost_basis"])
        realised_pnls.append(h["position"]["realised_pnl"])
    cost_total = sum(cost_bases)
    realised_total = sum(realised_pnls)

    workbook_unrealised = num(an["C6"].value)
    workbook_realised = num(dash["P25"].value)
    total_gain = mv_total + sells - buys
    portfolio_irr = xirr(all_flows + [(asof, mv_total)])

    add_check(checks, group, "total market value = sum of holdings", mv_total, num(an["B6"].value))
    add_check(checks, group, "total cost basis = sum of holdings", cost_total, num(an["A6"].value))
    add_check(checks, group, "unrealised P&L = market value - cost basis", mv_total - cost_total, workbook_unrealised)
    add_check(checks, group, "realised P&L = sum of holdings' realised P&L", realised_total, workbook_realised)
    add_check(checks, group, "total return = (value + sale proceeds - purchases) / purchases",
              total_gain / buys, num(an["D6"].value))
    add_check(checks, group, "total gain = realised + unrealised P&L", total_gain,
              (workbook_unrealised or 0) + (workbook_realised or 0))
    add_check(checks, group, "portfolio annual return = XIRR of every trade plus today's value",
              portfolio_irr, num(an["E6"].value), tol=IRR_TOL)
    return {"cost_total": cost_total, "realised_total": realised_total,
            "total_gain": total_gain, "irr": portfolio_irr}


def check_concentration(checks, wb, holdings, mv_total):
    """Group 3: sector totals, HHI and the spread between the largest and smallest position.

    Returns the sectors, the holding weights and the effective number of positions.
    """
    group = "Concentration"
    an, ra = wb["Analytics"], wb["Risk Analytics"]

    sectors = []
    for row in range(SECTOR_FIRST, SECTOR_LAST + 1):
        name = an.cell(row, 1).value
        market_value = num(an.cell(row, 2).value)
        if name and market_value is not None:
            sectors.append({"name": name, "market_value": market_value,
                            "weight": num(an.cell(row, 3).value)})

    sector_values = []
    sector_weights = []
    sector_squared_weights = []
    for sector in sectors:
        sector_values.append(sector["market_value"])
        sector_weights.append(sector["weight"])
        sector_squared_weights.append(sector["weight"] ** 2)
    sector_mv_sum = sum(sector_values)
    sector_weight_sum = sum(sector_weights)
    sector_hhi = sum(sector_squared_weights)

    weights = []
    for h in holdings:
        weights.append(h["market_value"] / mv_total)
    squared_weights = []
    for w in weights:
        squared_weights.append(w * w)
    holding_hhi = sum(squared_weights)
    effective_positions = 1 / holding_hhi

    add_check(checks, group, "sector market values sum to the portfolio total", sector_mv_sum, mv_total)
    add_check(checks, group, "sector weights sum to 100%", sector_weight_sum, 1.0)
    add_check(checks, group, "HHI = sum of squared sector weights", sector_hhi, num(an.cell(HHI_ROW, 4).value))
    add_check(checks, group, "effective positions = 1 / HHI of holding weights",
              effective_positions, num(ra[RA["eff_pos"]].value))
    add_check(checks, group, "largest / smallest position ratio",
              max(weights) / min(weights), num(ra[RA["spread"]].value))
    return sectors, weights, effective_positions


def check_backcast(checks, wb, holdings, mv_total):
    """Group 4: today's holdings valued at each of the last 12 months of closes.

    Returns the daily returns and the risk figures built from them.
    """
    group = "12-month back-cast of current holdings"
    spark, ps = wb["Price History"], wb["Portfolio Series"]

    # Row 2 of Price History holds the dates, starting in column B.
    dates = []
    for col in range(2, spark.max_column + 1):
        cell_value = spark.cell(2, col).value
        if num(cell_value) is not None:
            dates.append(cell_value)
    ndays = len(dates)

    daily_value = [0.0] * ndays
    covered = 0
    covered_mv = 0.0
    for i, h in enumerate(holdings):
        closes = []
        for col in range(2, 2 + ndays):
            closes.append(spark.cell(SPARK_FIRST + i, col).value)
        if num(closes[0]) is None:
            continue    # no price history for this holding (RTX in London), so it stays out
        covered += 1
        covered_mv += h["market_value"]
        for day in range(ndays):
            daily_value[day] += h["position"]["net_units"] * (num(closes[day]) or 0.0)

    daily_returns = []
    for day in range(1, ndays):
        daily_returns.append(daily_value[day] / daily_value[day - 1] - 1)

    reported_value = []
    for col in range(2, 2 + ndays):
        reported_value.append(num(ps.cell(5, col).value))

    ret12 = daily_value[-1] / daily_value[0] - 1
    vol = statistics.stdev(daily_returns) * math.sqrt(TRADING_DAYS)

    squared_losses = []
    for r in daily_returns:
        if r < 0:
            squared_losses.append(r * r)
    downside = math.sqrt(sum(squared_losses) / len(daily_returns)) * math.sqrt(TRADING_DAYS)

    # max drawdown: the worst fall from the highest value seen so far
    peak = daily_value[0]
    max_drawdown = 0.0
    max_drawdown_usd = 0.0
    for v in daily_value:
        peak = max(peak, v)
        max_drawdown = min(max_drawdown, v / peak - 1)
        max_drawdown_usd = min(max_drawdown_usd, v - peak)
    hist_var = -percentile_inc(daily_returns, 0.05)

    add_gap_check(checks, group, f"daily value = net units x close, summed ({ndays} days)",
                  list(zip(daily_value, reported_value)))
    add_check(checks, group, "holdings with price history", float(covered), num(ps[PS["covered"]].value))
    add_check(checks, group, "share of market value covered", covered_mv / mv_total, num(ps[PS["share"]].value))
    add_check(checks, group, "12-month return = last value / first value - 1", ret12, num(ps[PS["ret12"]].value))
    add_check(checks, group, "volatility = sample stdev of daily returns x sqrt(252)", vol, num(ps[PS["vol"]].value))
    add_check(checks, group, "downside deviation = sqrt(mean of squared negative returns) x sqrt(252)",
              downside, num(ps[PS["downside"]].value))
    add_check(checks, group, "max drawdown = worst fall from a running peak", max_drawdown, num(ps[PS["mdd"]].value))
    add_check(checks, group, "historical VaR (95%, 1-day) = minus the 5th percentile return",
              hist_var, num(ps[PS["hvar"]].value))
    return {"ndays": ndays, "covered": covered, "daily_returns": daily_returns, "ret12": ret12,
            "vol": vol, "downside": downside, "max_drawdown": max_drawdown,
            "max_drawdown_usd": max_drawdown_usd, "hist_var": hist_var}


def check_spy(checks, wb, backcast):
    """The SPY part of the back-cast group: SPY return, beta, tracking error, information ratio.

    Returns the SPY figures, or None when a SPY close is missing (that is itself a failed check).
    """
    group = "12-month back-cast of current holdings"
    ps = wb["Portfolio Series"]
    ndays = backcast["ndays"]
    daily_returns = backcast["daily_returns"]

    spy = []
    for col in range(2, 2 + ndays):
        spy.append(num(ps.cell(7, col).value))
    for close in spy:
        if close is None:
            add_check(checks, group, "SPY closes are present for every day", None, "missing")
            return None

    spy_returns = []
    for day in range(1, ndays):
        spy_returns.append(spy[day] / spy[day - 1] - 1)
    spy12 = spy[-1] / spy[0] - 1

    mean_portfolio = statistics.fmean(daily_returns)
    mean_spy = statistics.fmean(spy_returns)
    covariances = []
    active_returns = []
    for portfolio_return, spy_return in zip(daily_returns, spy_returns):
        covariances.append((portfolio_return - mean_portfolio) * (spy_return - mean_spy))
        active_returns.append(portfolio_return - spy_return)
    spy_variances = []
    for spy_return in spy_returns:
        spy_variances.append((spy_return - mean_spy) ** 2)
    beta = sum(covariances) / sum(spy_variances)
    tracking_error = statistics.stdev(active_returns) * math.sqrt(TRADING_DAYS)

    add_check(checks, group, "SPY 12-month return", spy12, num(ps[PS["spy12"]].value))
    add_check(checks, group, "beta vs SPY = slope of daily returns", beta, num(ps[PS["beta_spy"]].value))
    add_check(checks, group, "tracking error = stdev of (portfolio - SPY) x sqrt(252)",
              tracking_error, num(ps[PS["te"]].value))
    add_check(checks, group, "information ratio = active return / tracking error",
              (backcast["ret12"] - spy12) / tracking_error, num(ps[PS["ir"]].value))
    return {"spy12": spy12, "beta": beta, "tracking_error": tracking_error, "returns": spy_returns}


def check_capm(checks, wb, holdings, weights, backcast):
    """Group 5: CAPM, alpha and the risk ratios built from the back-cast.

    Returns the risk-free rate, which the next group needs.
    """
    group = "CAPM and risk ratios"
    ra = wb["Risk Analytics"]
    ret12, vol = backcast["ret12"], backcast["vol"]

    rf = num(ra[RA["rf"]].value)
    rm = num(ra[RA["rm"]].value)

    # The workbook's beta list is matched to our holdings by row, so the order must be the same.
    first_row = 39
    order_mismatches = 0
    betas = []
    for i, h in enumerate(holdings):
        if ra.cell(first_row + i, 2).value != h["ticker"]:
            order_mismatches += 1
        betas.append(num(ra.cell(first_row + i, 4).value))
    add_check(checks, group, "risk decomposition lists the holdings in Dashboard order (betas matched by ticker)",
              0.0, float(order_mismatches))

    weighted_betas = []
    for weight, holding_beta in zip(weights, betas):
        weighted_betas.append(weight * holding_beta)
    beta = sum(weighted_betas)
    capm = rf + beta * (rm - rf)

    add_check(checks, group, "equity risk premium = Rm - Rf", rm - rf, num(ra[RA["erp"]].value))
    add_check(checks, group, "portfolio beta = weighted average of holding betas", beta, num(ra[RA["beta"]].value))
    add_check(checks, group, "CAPM expected return = Rf + beta x (Rm - Rf)", capm, num(ra[RA["capm"]].value))
    add_check(checks, group, "excess over CAPM expected return (alpha at the assumed Rm) = 12m return - CAPM return",
              ret12 - capm, num(ra[RA["alpha"]].value))
    add_check(checks, group, "Treynor = (12m return - Rf) / beta", (ret12 - rf) / beta, num(ra[RA["treynor"]].value))
    add_check(checks, group, "Sharpe = (12m return - Rf) / volatility", (ret12 - rf) / vol,
              num(ra[RA["sharpe"]].value))
    add_check(checks, group, "Sortino = (12m return - Rf) / downside deviation",
              (ret12 - rf) / backcast["downside"], num(ra[RA["sortino"]].value))
    add_check(checks, group, "Calmar = 12m return / |max drawdown|", ret12 / abs(backcast["max_drawdown"]),
              num(ra[RA["calmar"]].value),
              note=("A cell in error here would still leave a Validation tab that skips "
                    "errors printing ALL PASS; this workbook's tab now counts errors as failures."))
    add_check(checks, group, "parametric VaR 95% = CAPM return - 1.645 x volatility",
              capm - 1.645 * vol, num(ra[RA["var95"]].value))
    add_check(checks, group, "parametric VaR 99% = CAPM return - 2.326 x volatility",
              capm - 2.326 * vol, num(ra[RA["var99"]].value))
    add_check(checks, group, "expected shortfall 95% = CAPM return - 2.063 x volatility",
              capm - 2.063 * vol, num(ra[RA["cvar"]].value))
    return rf


def check_money_at_risk(checks, wb, path, sectors, mv_total, backcast, spy, rf):
    """Group 6: capital market line, M-squared and the dollar risk figures.

    Returns the largest sector and the sector-shock input.
    """
    group = "Market line and money at risk"
    ra, ps = wb["Risk Analytics"], wb["Portfolio Series"]
    ret12, vol = backcast["ret12"], backcast["vol"]

    if spy is not None:
        spy_vol = statistics.stdev(spy["returns"]) * math.sqrt(TRADING_DAYS)
        cml = rf + (spy["spy12"] - rf) / spy_vol * vol
        add_check(checks, group, "SPY volatility = sample stdev of SPY daily returns x sqrt(252)",
                  spy_vol, num(ps[PS["spy_vol"]].value))
        add_check(checks, group, "capital market line return at the portfolio's volatility (SPY, 12m)",
                  cml, num(ra[CML["cml"]].value))
        add_check(checks, group, "portfolio vs CML spread = 12m return - CML return",
                  ret12 - cml, num(ra[CML["spread"]].value))
        add_check(checks, group, "M-squared = Sharpe x SPY volatility + Rf - SPY return",
                  (ret12 - rf) / vol * spy_vol + rf - spy["spy12"], num(ra[CML["m2"]].value))

    top_sector = sectors[0]
    for s in sectors:
        if s["market_value"] > top_sector["market_value"]:
            top_sector = s
    shock = num(ra[MONEY["shock"]].value)

    add_check(checks, group, "largest industry's market value", top_sector["market_value"],
              num(ra[MONEY["top_sector"]].value))
    add_check(checks, group, "loss if the largest industry fell by the shock input",
              (shock or 0) * top_sector["market_value"], num(ra[MONEY["loss"]].value))
    add_check(checks, group, "worst 12-month drawdown in dollars, current holdings",
              backcast["max_drawdown_usd"], num(ra[MONEY["mdd_usd"]].value))
    add_check(checks, group, "1-day historical VaR 95% in dollars, on full portfolio value",
              backcast["hist_var"] * mv_total, num(ra[MONEY["hvar_usd"]].value))
    if path is not None:
        add_check(checks, group, "no formula calls TODAY(); every date window ends on ValuationDate",
                  0.0, float(uses_today(path)))
    return top_sector, shock


def build(wb, path=None):
    """Re-derive every checked figure from the ledger and price history and pair it with the workbook's figure.

    Returns the list of Result checks and a dict of headline figures for the report.
    """
    checks = []
    trades = read_ledger(wb["Ledger"])
    pos = positions(trades)
    asof = as_date(wb["Portfolio Series"][PS["asof"]].value)

    holdings, mv_total = check_holdings(checks, wb, trades, pos, asof)
    totals = check_totals(checks, wb, trades, holdings, mv_total, asof)
    sectors, weights, effective_positions = check_concentration(checks, wb, holdings, mv_total)
    backcast = check_backcast(checks, wb, holdings, mv_total)
    spy = check_spy(checks, wb, backcast)
    rf = check_capm(checks, wb, holdings, weights, backcast)
    top_sector, shock = check_money_at_risk(checks, wb, path, sectors, mv_total, backcast, spy, rf)

    # The figures the report prints under the checks.
    top_holding = holdings[0]
    for h in holdings:
        if h["market_value"] > top_holding["market_value"]:
            top_holding = h

    held_tickers = []
    for h in holdings:
        held_tickers.append(h["ticker"])
    realised_by_ticker = []
    held_costs = []
    for ticker, state in running_average(trades).items():
        realised_by_ticker.append(state["realised"])
        if ticker in held_tickers:
            held_costs.append(state["cost"])
    running_realised = sum(realised_by_ticker)
    running_cost = sum(held_costs)

    first_trade_year = min(t["date"] for t in trades).year
    report = {
        "holdings": len(holdings),
        "sectors": len(sectors),
        "mv_total": mv_total,
        "asof": asof,
        "covered": backcast["covered"],
        "ndays": backcast["ndays"],
        "since": first_trade_year,
        "shock": shock or 0.0,
        "running_realised": running_realised,
        "running_cost": running_cost,
        "pooled_realised": totals["realised_total"],
        "pooled_cost": totals["cost_total"],
        "irr": totals["irr"],
        "ret12": backcast["ret12"],
        "vol": backcast["vol"],
        "max_drawdown": backcast["max_drawdown"],
        "sharpe": (backcast["ret12"] - rf) / backcast["vol"],
        "effective_positions": effective_positions,
        "spy12": None,
        "beta_vs_spy": None,
        "tracking_error": None,
        "total_gain": totals["total_gain"],
        "hist_var_usd": backcast["hist_var"] * mv_total,
        "max_drawdown_usd": backcast["max_drawdown_usd"],
        "top_holding_name": top_holding["ticker"],
        "top_holding_value": top_holding["market_value"],
        "top_holding_share": top_holding["market_value"] / mv_total,
        "top_sector_name": top_sector["name"],
        "top_sector_value": top_sector["market_value"],
        "top_sector_share": top_sector["market_value"] / mv_total,
    }
    if spy is not None:
        report["spy12"] = spy["spy12"]
        report["beta_vs_spy"] = spy["beta"]
        report["tracking_error"] = spy["tracking_error"]
    return checks, report


def main(argv):
    """Run the checks on the workbook, print the report and return the exit code (0 when all pass)."""
    path = Path(argv[1]) if len(argv) > 1 else DEFAULT
    if not path.exists():
        sys.exit(f"workbook not found: {path}")

    wb = load(path)
    checks, report = build(wb, path)

    print(f"Workbook: {path.name}   Valuation date: {report['asof']}")
    print(f"Holdings: {report['holdings']}   Sectors: {report['sectors']}   "
          f"Market value: {report['mv_total']:,.2f}")
    print(f"Back-cast: {report['covered']} of {report['holdings']} holdings have price history, "
          f"{report['ndays']} trading days")
    print("-" * 78)

    failed = 0
    group = None
    for check in checks:
        if check.group != group:
            group = check.group
            print(f"\n  {group}")
        if check.ok:
            print(f"    PASS  {check.label}")
        else:
            failed += 1
            print(f"    FAIL  {check.label}")
            print(f"            {check.detail}")
            if check.note:
                for line in _wrap(check.note, 64):
                    print(f"            {line}")

    print("\n" + "-" * 78)
    print(f"{len(checks) - failed} of {len(checks)} figures re-derived independently and matched.")
    if failed:
        print(f"{failed} FAILED. The workbook disagrees with its own inputs on these.")
        return 1
    print("Every figure was rebuilt from the ledger, prices and inputs, and agrees with the workbook.")
    print_report(report)
    return 0


def print_report(report):
    """Print the headline figures, the money at risk and the cost-basis advisory."""
    since = report["since"]
    shock = report["shock"]
    print()
    print("  Headline figures (re-derived above)")
    print(f"    annual return since {since} (XIRR)                  {report['irr']:>11.1%}")
    if report["spy12"] is not None:
        print(f"    12-month return, current holdings vs SPY       {report['ret12']:>6.1%} vs {report['spy12']:.1%}")
    print(f"    volatility / max drawdown (12m)                {report['vol']:>6.1%} / {report['max_drawdown']:.1%}")
    print(f"    Sharpe ratio                                    {report['sharpe']:>11.2f}")
    if report["beta_vs_spy"] is not None:
        print(f"    beta vs SPY / tracking error                   {report['beta_vs_spy']:>6.2f} / {report['tracking_error']:.1%}")
    print(f"    effective number of holdings                    {report['effective_positions']:>11.1f}")
    print()
    print("  Money at risk (from the figures above)")
    print(f"    total gain since {since} (realised + unrealised)   {report['total_gain']:>12,.0f}")
    print(f"    realised P&L on sales                           {report['pooled_realised']:>12,.0f}")
    print(f"    1-day historical VaR, 95% (on full value)        {report['hist_var_usd']:>12,.0f}")
    print(f"    worst 12-month drawdown, current holdings        {report['max_drawdown_usd']:>12,.0f}")
    print(f"    largest holding {report['top_holding_name']:<6} ({report['top_holding_share']:.1%} of value)          "
          f"{report['top_holding_value']:>12,.0f}")
    print(f"    largest sector ({report['top_sector_share']:.1%} of value)                 "
          f"{report['top_sector_value']:>12,.0f}  {report['top_sector_name']}")
    print(f"    loss if that sector fell {shock:.0%}                     {shock * report['top_sector_value']:>12,.0f}"
          f"  ({shock * report['top_sector_share']:.1%} of the portfolio)")
    print()
    print("  Advisory: cost basis method (not a pass/fail check)")
    print(f"    realised P&L    pooled average {report['pooled_realised']:>10,.0f}   running average {report['running_realised']:>10,.0f}")
    print(f"    cost basis      pooled average {report['pooled_cost']:>10,.0f}   running average {report['running_cost']:>10,.0f}")
    print("    The workbook pools every purchase of a stock; a running average uses only")
    print("    purchases made before each sale. Total gain is the same under both.")


def _wrap(text, width):
    """Split text into lines of at most `width` characters, breaking between words."""
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
