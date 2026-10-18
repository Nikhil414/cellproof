"""Plan 2b review regressions; synthetic expectations, not Excel-saved fixtures."""
import pytest
from xlsxkit import book_of

import unexcel.compiler as compiler
from unexcel.compiler import compile_book
from unexcel.verifier import execute


def result(formula, inputs, expected):
    book = book_of({"S": {**inputs, "D1": (formula, expected), "E1": ("=1+1", 2)}})
    compiled = compile_book(book)
    assert not compiled.unresolved
    got = execute(book, compiled)
    assert got[("S", 1, 5)] == ("n", 2.0, None)
    kind, number, text = got[("S", 1, 4)]
    expected_kind = "n"
    if isinstance(expected, str):
        expected_kind = "e" if expected.startswith("#") else "s"
    assert kind == expected_kind
    return text if kind in ("s", "e") else number


@pytest.mark.parametrize("formula,inputs", [
    ("=INDEX(A1:A2,1E30)", {"A1": 1, "A2": 2}),
    ("=CHOOSE(1E30,1,2)", {}),
    ('=LEFT("abc",1E30)', {}),
    ("=EDATE(46023,1E30)", {}),
    ("=ROUNDUP(1,1000)", {}),
])
def test_supported_numeric_inputs_do_not_abort_batch(formula, inputs):
    # A bad argument must never prevent the companion cell from completing.
    book = book_of({"S": {**inputs, "D1": (formula, 1), "E1": ("=1+1", 2)}})
    got = execute(book, compile_book(book))
    assert got[("S", 1, 5)] == ("n", 2.0, None)


@pytest.mark.parametrize("formula", ["=EDATE(2958465,1)", "=EDATE(1,-1)", "=EOMONTH(2958465,1)"])
def test_shifted_date_stays_in_excel_domain(formula):
    assert result(formula, {}, "#NUM!") == "#NUM!"


def test_rept_respects_excel_character_limit():
    assert result('=LEN(REPT("a",32768))', {}, "#VALUE!") == "#VALUE!"


@pytest.mark.parametrize("mode,direction,expected", [(-1, 1, 10), (1, -1, 20), (0, 1, 10), (0, -1, 20)])
def test_xlookup_duplicate_follows_search_direction(mode, direction, expected):
    inputs = {"A1": 5, "A2": 5, "B1": 10, "B2": 20}
    formula = f'=XLOOKUP(5,A1:A2,B1:B2,"none",{mode},{direction})'
    assert result(formula, inputs, expected) == expected


def test_median_keeps_each_range_argument():
    assert result("=MEDIAN(A1:A2,A1:A1)", {"A1": 0, "A2": 10}, 0) == 0


def test_lookup_preserves_empty_text():
    assert result('=VLOOKUP("x",A1:B1,2,FALSE)', {"A1": "x", "B1": ""}, "") == ""


def test_aggregate_keeps_argument_error_order():
    assert result("=MAX(#DIV/0!,A1:A1)", {"A1": "#N/A"}, "#DIV/0!") == "#DIV/0!"


def test_sql_guard_counts_text_and_kind(monkeypatch):
    monkeypatch.setattr(compiler, "MAX_SQL", 1000)
    book = book_of({"S": {"D1": ('=IF(TRUE,"' + 'x' * 940 + '",0)', 'x' * 940)}})
    compiled = compile_book(book)
    assert compiled.unresolved == {("S", 1, 4): compiler.TOO_BIG}


@pytest.mark.parametrize("formula,expected", [
    ('=INDEX(A1:A2,1E30)', '#REF!'),
    ('=CHOOSE(1E30,1,2)', '#VALUE!'),
    ('=LEFT("abc",1E30)', 'abc'),
    ('=RIGHT("abc",1E30)', 'abc'),
    ('=MID("abc",1E30,1)', ''),
    ('=MID("abc",1,1E30)', 'abc'),
    ('=SUBSTITUTE("aba","a","x",1E30)', 'aba'),
    ('=REPLACE("abc",1E30,1E30,"x")', 'abcx'),
    ('=FIND("a","abc",1E30)', '#VALUE!'),
    ('=SEARCH("a","abc",1E30)', '#VALUE!'),
    ('=EDATE(46023,1E30)', '#NUM!'),
    ('=EOMONTH(46023,1E30)', '#NUM!'),
    ('=WEEKDAY(1E30)', '#NUM!'),
    ('=ROUNDUP(1,1000)', 1),
    ('=ROUNDDOWN(1,1000)', 1),
    ('=TRUNC(1,1000)', 1),
    ('=LEN(REPT("a",32767))', 32767),
    ('=LEN(REPT("ab",16384))', '#VALUE!'),
    ('=REPT("",1E30)', ''),
    ('=REPT("a",-1)', '#VALUE!'),
])
def test_numeric_bounds_and_text_limits(formula, expected):
    assert result(formula, {"A1": 1, "A2": 2}, expected) == expected

@pytest.mark.parametrize("name", ["AVERAGE", "MIN", "MAX", "PRODUCT"])
@pytest.mark.parametrize("first", [True, False])
def test_aggregate_error_order_in_both_directions(name, first):
    arguments = '#DIV/0!,A1:A1' if first else 'A1:A1,#DIV/0!'
    # Real Excel: MAX(H3,1/0) with H3 = #N/A is #DIV/0! - typed-in errors come before errors inside cells.
    # Only MAX was seen in Excel; the others follow the same rule until examples/functions.xlsx says otherwise.
    expected = '#DIV/0!'
    assert result(f'={name}({arguments})', {"A1": "#N/A"}, expected) == expected

@pytest.mark.parametrize("formula", [
    '=HLOOKUP("x",A1:A2,2,FALSE)', '=INDEX(A2:A2,1)',
    '=XLOOKUP("x",A1:A1,A2:A2)',
])
def test_shared_lookup_reader_keeps_empty_text(formula):
    assert result(formula, {"A1": "x", "A2": ""}, "") == ""


def test_index_rejects_large_offsets_before_address_arithmetic():
    assert result('=INDEX(A1000000:A1000001,9.2233720368547748E18)',
                  {"A1000000": 1, "A1000001": 2}, "#REF!") == "#REF!"
