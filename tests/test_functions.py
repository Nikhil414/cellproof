import pytest
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute


def run(sheets):
    book = book_of(sheets)
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def test_sum_mixes_ranges_scalars_and_formula_cells():
    got, c = run({"S": {"A1": 1, "A2": 2, "A3": "note", "A4": True, "A5": ("=A1*10", 10),
                        "B1": ("=SUM(A1:A5, 100)", 113)}})
    assert c.unresolved == {} and got[("S", 1, 2)] == ("n", 113.0, None)


def test_sum_propagates_errors_and_handles_empty_ranges():
    got, _ = run({"S": {"A1": ("=1/0", "#DIV/0!"), "B1": ("=SUM(A1:A2)", "#DIV/0!"), "C1": ("=SUM(Z1:Z9)", 0)}})
    assert got[("S", 1, 2)] == ("e", None, "#DIV/0!")
    assert got[("S", 1, 3)] == ("n", 0.0, None)


def test_sum_reports_numbers_stored_as_text():
    got, c = run({"S": {"A1": 5, "A2": "15,000", "A3": ("=SUM(A1:A2)", 5)}})
    assert got[("S", 3, 1)] == ("n", 5.0, None)
    assert c.findings == ["S!A2: number stored as text ('15,000') is skipped by Excel"]


def test_if_round_and_or_iferror():
    got, c = run({"S": {
        "A1": 120000, "A2": 2.675,
        "B1": ("=IF(A1>100000, A1*5%, A1*2%)", 6000),
        "B2": ("=ROUND(A2, 2)", 2.68),
        "B3": ("=AND(A1>0, OR(A2<0, TRUE))", True),
        "B4": ("=IFERROR(A1/0, -1)", -1),
        "B5": ('=IF(A1>0, "big", "small")', "big"),
    }})
    assert c.unresolved == {}
    assert got[("S", 1, 2)][1] == pytest.approx(6000)
    assert got[("S", 2, 2)][1] == pytest.approx(2.68)
    assert got[("S", 3, 2)] == ("b", 1.0, None)
    assert got[("S", 4, 2)] == ("n", -1.0, None)
    assert got[("S", 5, 2)] == ("s", None, "big")


def test_if_with_wrong_argument_count_is_unresolved():
    _, c = run({"S": {"A1": 1, "B1": ('=IF(A1>0, "yes", 0)', "yes"), "B2": ("=IF(A1)", 0)}})
    assert ("S", 1, 2) not in c.unresolved  # mixed branch types are fine since Plan 2b
    assert c.unresolved[("S", 2, 2)] == "IF needs 2-3 arguments, got 1"
