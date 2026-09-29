"""Tests for validate_portfolio.py, by breaking a copy of the workbook.

A validator that has only ever seen a correct workbook proves nothing. It would
report every figure as matching just as confidently if its comparisons were
inverted, its tolerance were infinite, or it were reading the wrong cells. The
only way to know it works is to hand it a workbook that is wrong and confirm it
says so, and to name which figure is wrong.

Each test copies the workbook, changes one cell, and asserts that specific
failure is reported. They cover every group the validator checks: holdings
rebuilt from the ledger, portfolio totals, concentration, the 12-month back-cast
and the CAPM and risk ratios.

Several reproduce defects this workbook really had: sales added to positions
instead of netted, a volatility that was not measured from prices, and a cell in
error that a pass count skipped rather than failed.
"""
import os
import shutil
import stat
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

from xlsx_surgery import ref, set_cached_value

REPO = Path(__file__).resolve().parent.parent
WORKBOOK = REPO / "Stock Portfolio.xlsx"
SCRIPT = REPO / "validate_portfolio.py"

# Analytics sheet
ROW_TOTALS = 6              # cost basis, market value, P&L, return, IRR, holdings
COL_MARKET_VALUE = 2
COL_IRR = 5
ROW_FIRST_STOCK = 10
COL_STOCK_WEIGHT = 6
COL_REALISED = 9
ROW_HHI = 39
COL_HHI = 4
# Dashboard sheet
ROW_FIRST_HOLDING = 7
COL_UNITS = 4
# Portfolio Series sheet
ROW_VOLATILITY = 14


def run(workbook: Path):
    r = subprocess.run(
        [sys.executable, str(SCRIPT), str(workbook)],
        capture_output=True, text=True, cwd=REPO,
    )
    return r.returncode, r.stdout + r.stderr


def committed_workbook(tmp_path):
    """The workbook as committed, not as it currently sits on disk.

    These tests inject one fault and assert that fault is reported. That only
    means something against a clean baseline. The working copy on a developer's
    machine may be mid-edit and already failing, which would let a fault
    injection test pass without the injected fault doing anything.

    On a clean checkout, including CI, the two are identical.
    """
    out = tmp_path / "committed.xlsx"
    r = subprocess.run(
        ["git", "show", f"HEAD:{WORKBOOK.name}"],
        cwd=REPO, capture_output=True,
    )
    if r.returncode == 0 and r.stdout:
        out.write_bytes(r.stdout)
    else:
        shutil.copyfile(WORKBOOK, out)
    os.chmod(out, stat.S_IWRITE | stat.S_IREAD)
    return out


@pytest.fixture
def book(tmp_path):
    """A writable copy of the committed workbook for one test to damage."""
    counter = {"n": 0}

    def _copy():
        counter["n"] += 1
        base = committed_workbook(tmp_path)
        dst = tmp_path / f"book{counter['n']}.xlsx"
        shutil.copyfile(base, dst)
        os.chmod(dst, stat.S_IWRITE | stat.S_IREAD)
        return dst
    return _copy


@pytest.fixture
def clean_book(tmp_path):
    """The committed workbook, unmodified."""
    return committed_workbook(tmp_path)


def edit(path: Path, sheet: str, edits: dict, *, as_error: bool = False):
    """Change cached cell values in place.

    Deliberately not openpyxl. It holds either formulas or cached results, never
    both, so saving with it discards every cached value in the workbook and the
    validator would then read None everywhere. See tests/xlsx_surgery.py.
    """
    for (row, col), value in edits.items():
        set_cached_value(path, sheet, ref(row, col), value, as_error=as_error)


# --------------------------------------------------------------------------
# Holdings and totals
# --------------------------------------------------------------------------

def test_broken_market_value_total_is_caught(book):
    """Change the portfolio total so it no longer equals the sum of holdings."""
    wb = book()
    edit(wb, "Analytics", {(ROW_TOTALS, COL_MARKET_VALUE): 1})
    code, out = run(wb)
    assert code == 1, out
    assert "total market value = sum of holdings" in out
    assert "FAIL" in out


def test_failure_says_how_far_off_it_is(book):
    """A failure that does not quantify the gap sends the reader back to the
    spreadsheet to work it out themselves."""
    wb = book()
    edit(wb, "Analytics", {(ROW_TOTALS, COL_MARKET_VALUE): 1})
    _, out = run(wb)
    assert "differs by" in out


def test_a_single_holding_weight_error_is_caught(book):
    """The per-holding checks report the largest discrepancy across all 16, so a
    single corrupted row still has to surface."""
    wb = book()
    edit(wb, "Analytics", {(ROW_FIRST_STOCK, COL_STOCK_WEIGHT): 0.99})
    code, out = run(wb)
    assert code == 1, out
    assert "weights sum to 100%" in out or "weight = market value / total" in out


# --------------------------------------------------------------------------
# Concentration
# --------------------------------------------------------------------------

def test_broken_hhi_is_caught(book):
    """HHI must equal the sum of squared sector weights."""
    wb = book()
    edit(wb, "Analytics", {(ROW_HHI, COL_HHI): 0.5})
    code, out = run(wb)
    assert code == 1, out
    assert "HHI = sum of squared sector weights" in out


def test_effective_positions_uses_holding_weights_not_sector_weights(clean_book):
    """Two measures share the name HHI: sector concentration, and the holding
    level one behind effective positions. The unmodified workbook passes only
    when the right basis is used."""
    code, out = run(clean_book)
    assert code == 0, out
    assert "effective positions = 1 / HHI of holding weights" in out


# --------------------------------------------------------------------------
# CAPM and risk ratios
# --------------------------------------------------------------------------

def test_broken_capm_expected_return_is_caught(book):
    wb = book()
    edit(wb, "Risk Analytics", {(10, 3): 0.5})
    code, out = run(wb)
    assert code == 1, out
    assert "CAPM expected return" in out


def test_broken_sharpe_is_caught(book):
    wb = book()
    edit(wb, "Risk Analytics", {(18, 3): 99})
    code, out = run(wb)
    assert code == 1, out
    assert "Sharpe" in out


def test_a_cell_in_error_is_caught_and_explained(book):
    """A division by zero in the Calmar cell. The validator must fail and say
    the workbook's value is not a number, rather than skip the row."""
    wb = book()
    edit(wb, "Risk Analytics", {(20, 3): "#DIV/0!"}, as_error=True)
    code, out = run(wb)
    assert code == 1, out
    assert "Calmar" in out
    assert "not a number" in out


# --------------------------------------------------------------------------
# Defects this workbook really had
# --------------------------------------------------------------------------

def test_sales_added_instead_of_netted_is_caught(book):
    """The original defect: a holding's units counted sales as purchases."""
    wb = book()
    edit(wb, "Dashboard", {(ROW_FIRST_HOLDING, COL_UNITS): 107.16})
    code, out = run(wb)
    assert code == 1, out
    assert "units = bought - sold" in out


def test_a_wrong_realised_pnl_is_caught(book):
    wb = book()
    edit(wb, "Analytics", {(ROW_FIRST_STOCK, COL_REALISED): 0})
    code, out = run(wb)
    assert code == 1, out
    assert "realised P&L = proceeds - units sold x average cost" in out


def test_a_wrong_portfolio_irr_is_caught(book):
    """The old workbook annualised over a fixed seven years; the return must be
    the XIRR of the dated trades."""
    wb = book()
    edit(wb, "Analytics", {(ROW_TOTALS, COL_IRR): 0.1259})
    code, out = run(wb)
    assert code == 1, out
    assert "portfolio annual return = XIRR" in out


def test_a_volatility_not_measured_from_prices_is_caught(book):
    """The old workbook set volatility to beta x a hardcoded 0.155."""
    wb = book()
    edit(wb, "Portfolio Series", {(ROW_VOLATILITY, 2): 0.2283})
    code, out = run(wb)
    assert code == 1, out
    assert "volatility = sample stdev of daily returns" in out


# --------------------------------------------------------------------------
# The unmodified workbook, and input handling
# --------------------------------------------------------------------------

def test_the_committed_workbook_reconciles(clean_book):
    """Every figure in the published workbook reconciles. If this fails, either
    the workbook changed or the validator did, and no other test in this file
    can be interpreted."""
    code, out = run(clean_book)
    assert code == 0, out
    assert "re-derived independently and matched" in out


def test_missing_file_is_reported_clearly(tmp_path):
    code, out = run(tmp_path / "nope.xlsx")
    assert code != 0
    assert "workbook not found" in out


def test_workbook_without_the_expected_sheets_is_reported_clearly(book):
    """Removing a sheet is a structural change, so openpyxl is the right tool
    here even though it drops cached values: the validator should refuse before
    it reads a single figure."""
    wb = book()
    b = openpyxl.load_workbook(wb)
    del b["Analytics"]
    b.save(wb)
    b.close()
    code, out = run(wb)
    assert code != 0
    assert "Analytics" in out
