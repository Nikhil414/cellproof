"""Array constants ({1,2;3,4}) behave like a range of those values wherever a range is taken."""
from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.parser import parse, refs
from cellproof.verifier import execute


def run(formulas, data=None):
    book = book_of({"S": formulas} | ({"D": data} if data else {}))
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def N(v):
    return ("n", float(v), None)


def test_array_constants_where_a_range_is_taken():
    got, c = run({
        "A1": ("=SUM({1,2;3,-4})", 2),
        "A2": ('=SUM({1,"2",TRUE})', 1),                    # like a range: text and TRUE skipped
        "A3": ('=MATCH("b",{"a","b","c"},0)', 2),
        "A4": ('=INDEX({10,20;30,40},2,1)', 30),
        "A5": ('=LOOKUP(75,{0,50,100},{"F","C","A"})', "C"),
        "A6": ('=VLOOKUP(2,{1,"x";2,"y"},2,FALSE)', "y"),
        "A7": ("=COUNTA({1,#N/A,\"x\"})", 3),
        "A8": ("=MAX({3,9,1})+1", 10),
        "A9": ("=SUMPRODUCT({1,2,3},{4,5,6})", 32),
        "A10": ("={7}*2", 14),                              # one value: just that value
    })
    assert c.unresolved == {}
    assert [got[("S", r, 1)] for r in range(1, 11)] == [
        N(2), N(1), N(2), N(30), ("s", None, "C"), ("s", None, "y"), N(3), N(10), N(32), N(14)]


def test_array_constants_where_one_value_is_expected_are_not_guessed():
    _, c = run({"A1": ("=A2*{1,2}", 0), "A2": 3, "A3": ('=IF({1,0},"a","b")', "a"), "A4": ("={1,2;3}", 0)})
    assert "array constant" in c.unresolved[("S", 1, 1)]
    assert "array constant" in c.unresolved[("S", 3, 1)]
    assert "rows of different lengths" in c.unresolved[("S", 4, 1)]


def test_array_constants_are_not_dependencies():
    assert list(refs(parse('SUM(A1,{1,2})'))) == ["A1"]
