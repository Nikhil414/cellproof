from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.model import Cell, Verdict
from cellproof.verifier import matches, same, verify


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


def test_same_agrees_with_matches_on_every_row_of_the_table_above():
    """`same` is `matches` with the cell pulled apart, and `prove --check` compares two recorded answers with it. If
    the two ever disagree, a bundle could pass the check and fail the run."""
    rows = [
        (Cell("S", 1, 1, "x", 1.0, "n"), "n", 1.0 + 1e-12, None),
        (Cell("S", 1, 1, "x", 1.0, "n"), "n", 1.0001, None),
        (Cell("S", 1, 1, "x", 1.0, "n"), "s", None, "1"),
        (Cell("S", 1, 1, "x", 1.0, "n"), "n", None, None),
        (Cell("S", 1, 1, "x", "Hi", "s"), "s", None, "Hi"),
        (Cell("S", 1, 1, "x", "Hi", "s"), "s", None, "hi"),
        (Cell("S", 1, 1, "x", True, "b"), "b", 1.0, None),
        (Cell("S", 1, 1, "x", "#N/A", "e"), "e", None, "#N/A"),
        (Cell("S", 1, 1, "x", "#N/A", "e"), "n", 0.0, None),
        (Cell("S", 1, 1, "x", 0.0, "n"), "n", -0.0, None),
        (Cell("S", 1, 1, "x", False, "b"), "b", 0.0, None),
        (Cell("S", 1, 1, "x", None, "blank"), "n", None, None),
    ]
    for cell, kind, num, txt in rows:
        assert same(kind, num, txt, cell.kind, cell.value) == matches(cell, kind, num, txt)
