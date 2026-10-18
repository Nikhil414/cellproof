from pathlib import Path

import pytest

from unexcel.compiler import compile_book
from unexcel.model import a1
from unexcel.reader import load
from unexcel.verifier import verify

EXAMPLE = Path(__file__).parent.parent / "examples" / "commission.xlsx"


@pytest.mark.skipif(not EXAMPLE.exists(), reason="save examples/commission.xlsx from real Excel (PLAN Task 12)")
def test_real_excel_answers_match():
    book = load(EXAMPLE)
    verdicts = verify(book, compile_book(book))
    failures = [v for v in verdicts if v.status != "PASS"]
    assert not failures, "\n".join(
        f"{a1(v.key)}: {v.status} expected={v.expected!r} actual={v.actual!r} {v.reason}" for v in failures)
    assert len(verdicts) == 4  # B2, C2, B3, B5


ERRORS_EXAMPLE = Path(__file__).parent.parent / "examples" / "errors.xlsx"


@pytest.mark.skipif(not ERRORS_EXAMPLE.exists(), reason="save examples/errors.xlsx from real Excel (PLAN-2a Task 6)")
def test_real_excel_error_codes_match():
    book = load(ERRORS_EXAMPLE)
    verdicts = verify(book, compile_book(book))
    failures = [v for v in verdicts if v.status != "PASS"]
    assert not failures, "\n".join(
        f"{a1(v.key)}: {v.status} expected={v.expected!r} actual={v.actual!r} {v.reason}" for v in failures)
    assert len(verdicts) == 15  # B1:B9, D2, E1, F1, H2, I1, I2


FUNCTIONS_EXAMPLE = Path(__file__).parent.parent / "examples" / "functions.xlsx"
FUNCTIONS_UNSAVED = Path(__file__).parent.parent / "examples" / "functions_unsaved.xlsx"


def _formula_count(path) -> int:
    return sum(1 for cell in load(path).cells.values() if cell.formula is not None)


@pytest.mark.skipif(not FUNCTIONS_EXAMPLE.exists(),
                    reason="save examples/functions.xlsx from real Excel (PLAN-2b Task 9)")
def test_real_excel_function_answers_match():
    wanted = _formula_count(FUNCTIONS_UNSAVED)  # every formula the generator writes
    if _formula_count(FUNCTIONS_EXAMPLE) != wanted:
        pytest.skip(f"examples/functions.xlsx is older than functions_unsaved.xlsx ({wanted} formulas): "
                    "run examples/make_functions.py, open the result in Excel and Save As examples/functions.xlsx")
    book = load(FUNCTIONS_EXAMPLE)
    verdicts = verify(book, compile_book(book))
    # UNRESOLVED is allowed only where unexcel refuses to guess what Excel's binary search does on unsorted data.
    failures = [v for v in verdicts
                if v.status != "PASS" and not (v.status == "UNRESOLVED" and "sorted data" in v.reason)]
    assert not failures, "\n".join(
        f"{a1(v.key)}: {v.status} expected={v.expected!r} actual={v.actual!r} {v.reason}" for v in failures)
    assert len(verdicts) == wanted
