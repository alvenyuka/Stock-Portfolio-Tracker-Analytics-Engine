"""Puts tests/ on sys.path so the test module can import its xlsx helper."""
import sys
from pathlib import Path

TESTS = Path(__file__).parent / "tests"
if str(TESTS) not in sys.path:
    sys.path.insert(0, str(TESTS))
