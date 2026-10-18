from xlsxkit import book_of

from unexcel.compiler import DEEP, TOO_BIG, compile_book
from unexcel.verifier import verify


def test_very_long_formula_is_unresolved_not_a_crash():
    terms = "+".join(["A1"] * 1500)
    book = book_of({"S": {"A1": 1, "B1": ("=" + terms, 1500), "C1": ("=A1*2", 2)}})
    compiled = compile_book(book)
    assert compiled.unresolved == {("S", 1, 2): DEEP}
    assert [v.status for v in verify(book, compiled)] == ["UNRESOLVED", "PASS"]  # B1, C1


def test_deeply_nested_parentheses_are_unresolved():
    book = book_of({"S": {"A1": ("=" + "(" * 600 + "1" + ")" * 600, 1)}})
    assert compile_book(book).unresolved == {("S", 1, 1): DEEP}


def test_formula_whose_sql_doubles_per_level_is_unresolved_not_out_of_memory():
    f = "A1"
    for _ in range(30):  # each level repeats its argument in the value and the error SQL: 2^30 growth
        f = f"IFERROR(1/{f},0)"
    book = book_of({"S": {"A1": 2, "B1": ("=" + f, 0), "C1": ("=A1*2", 4)}})
    compiled = compile_book(book)
    assert compiled.unresolved == {("S", 1, 2): TOO_BIG}
    assert [v.status for v in verify(book, compiled)] == ["UNRESOLVED", "PASS"]
