"""The rebalancing scenarios: the current portfolio must reproduce the workbook, and the
capping and cost arithmetic must do what the README says."""
import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

from scenarios import cap_group, cap_weights, run, trade_costs  # noqa: E402

WORKBOOK = REPO / "Stock Portfolio.xlsx"


@pytest.fixture(scope="module")
def result():
    return run(WORKBOOK)


def test_current_portfolio_reproduces_the_workbook(result):
    """Before any scenario can be trusted, the unchanged portfolio has to give the workbook's own figures."""
    wb = openpyxl.load_workbook(WORKBOOK, data_only=True)
    ra, ps = wb["Risk Analytics"], wb["Portfolio Series"]
    now = result["scenarios"]["Current portfolio"]
    assert now["sector_shock_loss"] == pytest.approx(ra["C74"].value)
    assert now["max_drawdown_usd"] == pytest.approx(ra["C75"].value)
    assert now["var_95_1d_usd"] == pytest.approx(ra["C76"].value)
    assert now["effective_holdings"] == pytest.approx(ra["C34"].value)
    assert now["volatility"] == pytest.approx(ps["B14"].value)
    assert now["total_cost"] == 0


def test_every_policy_keeps_the_portfolio_value_and_meets_its_cap(result):
    s = result["scenarios"]
    total = s["Current portfolio"]["market_value"]
    assert all(v["market_value"] == pytest.approx(total) for v in s.values())
    assert s["NVIDIA capped at 25%"]["nvidia_share"] == pytest.approx(0.25)
    assert s["No holding above 20%"]["largest_holding_share"] <= 0.20 + 1e-9
    assert s["Semiconductors capped at 40%"]["semiconductor_share"] == pytest.approx(0.40)


def test_capping_spreads_the_excess_in_proportion():
    target = cap_weights({"A": 60.0, "B": 30.0, "C": 10.0}, cap=0.5)
    assert target["A"] == pytest.approx(50.0)
    assert target["B"] / target["C"] == pytest.approx(3.0)        # the others keep their mix
    assert sum(target.values()) == pytest.approx(100.0)


def test_capping_repeats_until_nothing_is_over():
    """A's excess pushes B over the cap too, so the rule has to run again."""
    target = cap_weights({"A": 70.0, "B": 20.0, "C": 5.0, "D": 5.0}, cap=0.3)
    assert max(target.values()) <= 30.0 + 1e-9
    assert target["C"] == pytest.approx(target["D"])
    assert sum(target.values()) == pytest.approx(100.0)


def test_an_impossible_cap_is_rejected():
    with pytest.raises(ValueError):
        cap_weights({"A": 70.0, "B": 25.0, "C": 5.0}, cap=0.3)        # three holdings cannot all be <= 30%


def test_group_cap_keeps_the_group_mix():
    target = cap_group({"A": 40.0, "B": 20.0, "C": 40.0}, group={"A", "B"}, cap=0.3)
    assert target["A"] + target["B"] == pytest.approx(30.0)
    assert target["A"] / target["B"] == pytest.approx(2.0)


def test_tax_falls_on_net_realised_gains_only():
    holdings = [{"ticker": "A", "units": 10, "price": 100.0, "avg_cost": 40.0},
                {"ticker": "B", "units": 10, "price": 100.0, "avg_cost": 100.0}]
    costs = trade_costs(holdings, {"A": 500.0, "B": 1500.0})       # sell 5 A, buy 5 B
    assert costs["realised_gain"] == pytest.approx(5 * 60.0)
    assert costs["tax"] == pytest.approx(0.15 * 300.0)
    assert costs["trading_cost"] == pytest.approx(0.001 * 1000.0)
    loss = trade_costs([{"ticker": "A", "units": 10, "price": 100.0, "avg_cost": 150.0}], {"A": 500.0})
    assert loss["tax"] == 0
