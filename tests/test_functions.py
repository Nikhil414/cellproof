import pytest
from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute


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


def test_row_column_code_unicode_n_parity():
    from cellproof.compiler import compile_book
    from cellproof.verifier import execute
    from xlsxkit import book_of
    s = {"A1": "Apple", "A2": "€ uro", "A3": "Ā", "A4": True, "A5": 45000.0, "A6": "#DIV/0!",
         "B1": ("=ROW()", 1.0), "B2": ("=COLUMN()", 2.0), "B3": ("=ROW(D7)", 7.0), "B4": ("=COLUMN(F1)", 6.0),
         "B5": ("=ROW(A1:A3)", 1.0),                                   # a range outside an array formula
         "C1": ("=CODE(A1)", 65.0), "C2": ("=CODE(A2)", 128.0), "C3": ('=CODE("")', "#VALUE!"),
         "C4": ("=CODE(A3)", 0.0),                                     # not in Windows-1252: not guessed
         "C5": ("=_xlfn.UNICODE(A2)", 8364.0), "C6": ("=_xlfn.UNICHAR(8364)", "€"),
         "C7": ("=_xlfn.UNICHAR(0)", "#VALUE!"), "C8": ("=_xlfn.UNICHAR(55296)", "#N/A"),
         "D1": ("=N(A5)", 45000.0), "D2": ("=N(A4)", 1.0), "D3": ("=N(A1)", 0.0), "D4": ("=N(A6)", "#DIV/0!"),
         "D5": ("=ISODD(-3.7)", True), "D6": ("=ISEVEN(0)", True), "D7": ("=ISEVEN(A4)", "#VALUE!"),
         "D8": ("=ISODD(B3)", True)}
    book = book_of({"S": s})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert set(compiled.unresolved) == {("S", 5, 2)}
    want = {(1, 2): 1.0, (2, 2): 2.0, (3, 2): 7.0, (4, 2): 6.0, (1, 3): 65.0, (2, 3): 128.0, (5, 3): 8364.0,
            (1, 4): 45000.0, (2, 4): 1.0, (3, 4): 0.0}
    for (r, col), v in want.items():
        assert got[("S", r, col)] == ("n", v, None), (r, col)
    assert got[("S", 3, 3)] == ("e", None, "#VALUE!") and got[("S", 4, 3)] == ("e", None, "#UNCHECKED")
    assert got[("S", 6, 3)] == ("s", None, "€") and got[("S", 7, 3)] == ("e", None, "#VALUE!")
    assert got[("S", 8, 3)] == ("e", None, "#UNCHECKED") and got[("S", 4, 4)] == ("e", None, "#DIV/0!")
    assert [got[("S", r, 4)] for r in (5, 6, 7, 8)] == [("b", 1.0, None), ("b", 1.0, None),
                                                         ("e", None, "#VALUE!"), ("b", 1.0, None)]
