from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute, verify


def run(sheets):
    book = book_of(sheets)
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def test_if_and_iferror_may_return_text_or_a_number():  # the most common mixed-type patterns
    got, c = run({"S": {
        "A1": 5, "A2": -1, "A3": 0,
        "B1": ('=IF(A1>0, "yes", 0)', "yes"), "B2": ('=IF(A2>0, "yes", 0)', 0),
        "B3": ('=IFERROR(1/A3, "")', ""), "B4": ('=IFERROR(1/A1, "")', 0.2),
        "C1": ('=B1&"!"', "yes!"), "C2": ("=B2+1", 1), "C4": ("=SUM(B1:B4)", 0.2),
    }})
    assert c.unresolved == {}
    assert got[("S", 1, 2)] == ("s", None, "yes") and got[("S", 2, 2)] == ("n", 0.0, None)
    assert got[("S", 3, 2)] == ("s", None, "") and got[("S", 4, 2)] == ("n", 0.2, None)
    assert got[("S", 1, 3)] == ("s", None, "yes!") and got[("S", 2, 3)] == ("n", 1.0, None)
    assert got[("S", 4, 3)] == ("n", 0.2, None)  # SUM skips the text results


def test_text_to_number_and_number_to_text():
    got, c = run({"S": {
        "A1": 1e18,
        "B1": ('="3"+1', 4), "B2": ('=" 3 "*2', 6), "B3": ('="abc"+1', "#VALUE!"), "B4": ('="1,000"+1', "#VALUE!"),
        "B5": ('="x"&1/3', "x0.333333333333333"), "B6": ('="x"&0.1+0.2', "x0.3"), "B7": ('="x"&TRUE', "xTRUE"),
        "B8": ('=A1&""', "1000000000000000000"), "B9": ('=IF("TRUE", 1, 2)', 1), "B10": ('=IF("x", 1, 2)', "#VALUE!"),
    }})
    assert c.unresolved == {}
    assert [got[("S", r, 2)] for r in range(1, 11)] == [
        ("n", 4.0, None), ("n", 6.0, None), ("e", None, "#VALUE!"), ("e", None, "#VALUE!"),
        ("s", None, "x0.333333333333333"), ("s", None, "x0.3"), ("s", None, "xTRUE"), ("s", None, "1000000000000000000"),  # real Excel writes 1E+18 as plain digits
        ("n", 1.0, None), ("e", None, "#VALUE!")]


def test_comparisons_across_types_follow_excel_order():
    got, c = run({"S": {
        "A1": 5, "A2": "5", "B1": ('=1<"a"', True), "B2": ('="a"<TRUE', True), "B3": ("=A1=A2", False),
        "B4": ('=Z9=""', True), "B5": ("=Z9=0", True), "B6": ('=IF(A1>0,"x",1)>2', True),  # text > any number
    }})
    assert c.unresolved == {}
    assert [got[("S", r, 2)] for r in range(1, 7)] == [
        ("b", 1.0, None), ("b", 1.0, None), ("b", 0.0, None), ("b", 1.0, None), ("b", 1.0, None), ("b", 1.0, None)]


def test_mixed_vlookup_column_and_verdicts_on_mixed_cells():
    book = book_of({"R": {"A1": "a", "B1": 1, "A2": "b", "B2": "two"}, "S": {
        "A1": ("=VLOOKUP(\"a\", R!A1:B2, 2, FALSE)", 1), "A2": ("=VLOOKUP(\"b\", R!A1:B2, 2, FALSE)", "two"),
        "A3": ('=IF(A1=1, "", 5)', "")}})
    assert {v.status for v in verify(book, compile_book(book))} == {"PASS"}
