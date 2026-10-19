from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.model import Cell, Verdict
from cellproof.verifier import matches, verify


def test_pass_fail_and_unresolved():
    book = book_of({"S": {
        "A1": 2,
        "A2": ("=A1*2", 4),
        "A3": ("=A1*2", 5),        # Excel says 5: someone typed over the result
        "A4": ("=1/0", "#DIV/0!"),
        "A5": ("=A1", "#N/A"),
        "A6": ("=NOW()", 1),
    }})
    v = {x.key: x for x in verify(book, compile_book(book))}
    assert v[("S", 2, 1)].status == "PASS"
    assert (v[("S", 3, 1)].status, v[("S", 3, 1)].expected, v[("S", 3, 1)].actual) == ("FAIL", 5.0, 4.0)
    assert v[("S", 4, 1)].status == "PASS"
    assert (v[("S", 5, 1)].status, v[("S", 5, 1)].actual) == ("FAIL", 2.0)
    assert v[("S", 6, 1)] == Verdict(("S", 6, 1), "UNRESOLVED", 1.0, None, "volatile function NOW")


def test_matching_rules():
    n = Cell("S", 1, 1, "x", 1.0, "n")
    assert matches(n, "n", 1.0 + 1e-12, None)
    assert not matches(n, "n", 1.0001, None)
    assert not matches(n, "s", None, "1")
    assert not matches(n, "n", None, None)
    assert matches(Cell("S", 1, 1, "x", "Hi", "s"), "s", None, "Hi")
    assert not matches(Cell("S", 1, 1, "x", "Hi", "s"), "s", None, "hi")
    assert matches(Cell("S", 1, 1, "x", True, "b"), "b", 1.0, None)
    assert matches(Cell("S", 1, 1, "x", "#N/A", "e"), "e", None, "#N/A")
    assert not matches(Cell("S", 1, 1, "x", "#N/A", "e"), "n", 0.0, None)
