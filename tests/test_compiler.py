import pytest
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute


def run(sheets, names=None):
    book = book_of(sheets, names)
    compiled = compile_book(book)
    return execute(book, compiled), compiled.unresolved


def test_arithmetic_precedence_percent_and_divide_by_zero():
    got, bad = run({"S": {"A1": 2, "A2": ("=-A1^2+10%*A1", 4.2), "A3": ("=A1/0", "#DIV/0!")}})
    assert bad == {}
    kind, num, txt = got[("S", 2, 1)]
    assert kind == "n" and num == pytest.approx(4.2) and txt is None
    assert got[("S", 3, 1)] == ("e", None, "#DIV/0!")  # an error, not inf


def test_text_concat_and_case_insensitive_compare():
    got, _ = run({"S": {"A1": "Hi", "A2": ('=A1&" there"', "Hi there"), "A3": ('=A1="hi"', True)}})
    assert got[("S", 2, 1)] == ("s", None, "Hi there")
    assert got[("S", 3, 1)] == ("b", 1.0, None)


def test_formula_chains_and_blank_cells():
    got, _ = run({"S": {"A1": 3, "A2": ("=A1*2", 6), "A3": ("=A2+B9", 6), "A4": ("=B9", 0)}})
    assert got[("S", 3, 1)] == ("n", 6.0, None)
    assert got[("S", 4, 1)] == ("n", 0.0, None)


def test_named_ranges():
    got, bad = run({"S": {"A1": 0.1, "A2": ("=Rate*100", 10)}}, names={"Rate": ("S", "$A$1")})
    assert bad == {} and got[("S", 2, 1)][1] == pytest.approx(10.0)


def test_sheet_names_with_quotes_and_spaces():  # Review Focus 3
    got, bad = run({"Bob's Sheet": {"A1": 5}, "S": {"A1": ("='Bob''s Sheet'!A1*2", 10)}})
    assert bad == {} and got[("S", 1, 1)] == ("n", 10.0, None)


def test_unresolved_reasons():  # includes Review Focus 5 (no saved value)
    _, bad = run({"S": {
        "A1": "x",
        "A2": ("=A1+1", "#VALUE!"),
        "A3": ("=NOW()", 1),
        "A4": ("=A3+1", 2),
        "A5": ("=1", None),
        "A6": ("=FOO(1)", 1),
    }})
    assert ("S", 2, 1) not in bad  # text in arithmetic is Excel's #VALUE!, computed like any other error
    assert bad[("S", 3, 1)] == "volatile function NOW"
    assert bad[("S", 4, 1)] == "depends on unresolved cell S!A3"
    assert bad[("S", 5, 1)].startswith("no saved value")
    assert bad[("S", 6, 1)] == "function FOO not supported yet"


def test_workbook_without_formulas_gives_empty_result():
    got, bad = run({"S": {"A1": 1}})
    assert got == {} and bad == {}


from unexcel.compiler import Compiler, NewBatch  # noqa: E402
from unexcel.model import Unsupported  # noqa: E402
from unexcel.refs import Area  # noqa: E402


def test_ready_reports_open_unresolved_and_done_cells():  # Review Focus 4
    book = book_of({"S": {"A1": ("=1", 1), "A2": 5.0}})
    comp = Compiler(book)
    area = Area("S", 1, 1, 2, 1)
    with pytest.raises(Unsupported, match="depends on unresolved cell S!A1"):
        comp.ready(area)
    comp.open[("S", 1, 1)] = ("c1", "num", False)
    with pytest.raises(NewBatch):
        comp.ready(area)
    comp.open.clear()
    comp.done[("S", 1, 1)] = ("num", True)
    comp.index.complete(("S", 1, 1), "n", True)
    keys = comp.ready(area)
    assert len(keys) == 2 and keys.kinds == {"n"} and keys.may_err
    assert list(keys) == [("S", 1, 1), ("S", 2, 1)]


def test_keys_of_an_empty_range():
    comp = Compiler(book_of({"S": {"A1": 1.0}}))
    keys = comp.ready(Area("S", 5, 1, 9, 1))
    assert len(keys) == 0 and keys.kinds == frozenset() and not keys.may_err


from unexcel.compiler import cell_row  # noqa: E402
from unexcel.expr import Expr  # noqa: E402
from unexcel.parser import Binary, Num, RefNode  # noqa: E402


def test_cell_row_drops_the_error_cases_when_the_cell_cannot_fail():
    assert cell_row("num", "x", "NULL", "NULL", "NULL") == ("'n'", "CAST(x AS DOUBLE)", "NULL")
    assert cell_row("txt", "x", "NULL", "NULL", "NULL") == ("'s'", "CAST(NULL AS DOUBLE)", "CAST(x AS VARCHAR)")
    assert cell_row("bool", "x", "NULL", "NULL", "NULL") == ("'b'", "CAST(x AS DOUBLE)", "NULL")
    kind, num_sql, txt_sql = cell_row("num", "x", "e", "NULL", "NULL")
    assert "CASE WHEN e IS NULL" in kind and "CASE WHEN e IS NULL" in num_sql and txt_sql == "e"


def test_rowrefs_compile_placeholders_to_the_given_expressions():
    comp = Compiler(book_of({"S": {"A1": 1.0}}))
    comp.rowrefs = {"__ROWREF_1__": Expr("amount", "num")}
    e = comp.expr(Binary("*", RefNode("__ROWREF_1__"), Num(2.0)), "S")
    assert e.sql == "(amount * CAST(2.0 AS DOUBLE))" and e.type == "num"
    comp.rowrefs = None
    with pytest.raises(Unsupported):
        comp.expr(RefNode("__ROWREF_1__"), "S")  # without the hook the placeholder is an unknown name


def test_a_number_beyond_excels_range_is_unresolved_not_a_crash():
    from unexcel.verifier import verify
    book = book_of({"S": {"A1": 2.0, "B1": ("=1E+400*A1", "#NUM!"), "C1": ("=A1+1", 3.0)}})
    compiled = compile_book(book)
    assert "beyond Excel's range" in compiled.unresolved[("S", 1, 2)]
    assert {v.status for v in verify(book, compiled)} == {"PASS", "UNRESOLVED"}
