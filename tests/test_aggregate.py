from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute

DATA = {"A1": 4, "A2": "x", "A3": True, "A4": 8, "A5": "15,000", "B1": 2, "B2": 3, "B3": "", "C1": "#N/A"}


def run(formulas):
    book = book_of({"S": DATA, "F": formulas})
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def N(v):
    return ("n", float(v), None)


def test_aggregates_skip_text_and_logicals_in_ranges():
    got, c = run({
        "A1": ("=AVERAGE(S!A1:A9)", 6), "A2": ('=AVERAGE(S!A1:A2, "3", TRUE)', 2.6667), "A3": ("=MIN(S!A1:A9)", 4),
        "A4": ("=MAX(S!A1:A9, 20)", 20), "A5": ("=MAX(S!Z1:Z9)", 0), "A6": ("=COUNT(S!A1:B9)", 4),
        "A7": ("=COUNTA(S!A1:B9)", 8), "A8": ("=COUNTBLANK(S!B1:B4)", 2), "A9": ("=PRODUCT(S!A1:B9)", 192),
        "A10": ("=MEDIAN(S!A1:B9)", 3.5), "A11": ("=SUMPRODUCT(S!A1:A2, S!B1:B2)", 8),
        "A12": ("=AVERAGE(S!Z1:Z9)", "#DIV/0!"), "A13": ("=MAX(S!A1:A9, S!C1)", "#N/A"), "A14": ("=PRODUCT(S!Z1:Z2)", 0),
    })
    assert c.unresolved == {}
    assert got[("F", 2, 1)][1] == 8 / 3
    assert [got[("F", r, 1)] for r in (1, 3, 4, 5, 6, 7, 8, 9, 10, 11)] == [
        N(6), N(4), N(20), N(0), N(4), N(8), N(2), N(192), N(3.5), N(8)]
    assert got[("F", 12, 1)] == ("e", None, "#DIV/0!") and got[("F", 13, 1)] == ("e", None, "#N/A")
    assert got[("F", 14, 1)] == N(0)


def test_subtotal_leaves_out_other_subtotals():
    book = book_of({"S": {"A1": 1, "A2": 2, "A3": ("=SUBTOTAL(9, A1:A2)", 3), "A4": 10,
                          "A5": ("=SUBTOTAL(9, A1:A4)", 13), "A6": ("=SUBTOTAL(1, A1:A4)", 13 / 3),
                          "A7": ("=SUBTOTAL(109, A1:A4)", 13)}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert got[("S", 5, 1)] == N(13) and got[("S", 6, 1)][1] == 13 / 3
    assert compiled.unresolved[("S", 7, 1)] == "SUBTOTAL that ignores hidden rows not supported yet"
