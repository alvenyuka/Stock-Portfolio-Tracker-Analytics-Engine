"""Tests for validate_portfolio.py, by breaking a copy of the workbook.

A validator that has only ever seen a correct workbook proves nothing. It would
report every figure as matching just as confidently if its comparisons were
inverted, its tolerance were infinite, or it were reading the wrong cells. The
only way to know it works is to hand it a workbook that is wrong and confirm it
says so, and to name which figure is wrong.

Each test copies the workbook, changes one cell, and asserts that specific
failure is reported. They cover the three groups the validator checks:
portfolio accounting, concentration, and the CAPM and risk ratios.

One test deliberately reproduces a real regression. A division by zero in the
Calmar ratio cell leaves the workbook's own Validation tab still printing ALL
PASS, because its pass count tallies only the rows that evaluate and skips the
one in error. That is the exact failure mode this script exists to catch.
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
ROW_TOTALS = 6              # cost basis, market value, P&L, return, CAGR, holdings
COL_MARKET_VALUE = 2
ROW_FIRST_STOCK = 10
COL_STOCK_WEIGHT = 6
ROW_HHI = 39
COL_HHI = 4


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


def risk_row(path: Path, fragment: str):
    """Find a row on Risk Analytics whose column B label contains fragment."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Risk Analytics"]
    try:
        for r in range(1, ws.max_row + 1):
            label = ws.cell(r, 2).value
            if isinstance(label, str) and fragment.lower() in label.lower():
                return r
    finally:
        wb.close()
    raise AssertionError(f"no Risk Analytics row matching {fragment!r}")


# --------------------------------------------------------------------------
# Portfolio accounting
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
    """Regression test for a mistake in this validator rather than the workbook.

    Two different measures share the name HHI here. The HHI Index row is sector
    concentration; effective positions is the reciprocal of the holding-level
    HHI. Checking one against the other reported a defect that did not exist.
    The unmodified workbook must pass, which it only does when the right basis
    is used.
    """
    code, out = run(clean_book)
    assert code == 0, out
    assert "effective positions = 1 / HHI of holding weights" in out


# --------------------------------------------------------------------------
# CAPM and risk ratios
# --------------------------------------------------------------------------

def test_broken_capm_expected_return_is_caught(book):
    wb = book()
    row = risk_row(wb, "Expected Return")
    edit(wb, "Risk Analytics", {(row, 3): 0.5})
    code, out = run(wb)
    assert code == 1, out
    assert "CAPM expected return" in out


def test_broken_sharpe_is_caught(book):
    wb = book()
    row = risk_row(wb, "Sharpe")
    edit(wb, "Risk Analytics", {(row, 3): 99})
    code, out = run(wb)
    assert code == 1, out
    assert "Sharpe" in out


def test_a_cell_in_error_is_caught_and_explained(book):
    """The real regression. A division by zero in the Calmar cell leaves the
    workbook's own Validation tab printing ALL PASS, because its count skips
    rows that do not evaluate. The validator must fail and say why."""
    wb = book()
    row = risk_row(wb, "Calmar")
    edit(wb, "Risk Analytics", {(row, 3): "#DIV/0!"}, as_error=True)
    code, out = run(wb)
    assert code == 1, out
    assert "Calmar" in out
    assert "not a number" in out


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
