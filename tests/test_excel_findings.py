"""Regression tests for what real Microsoft Excel answered in examples/functions.xlsx (Windows run, 2026-10-14).

Each expected value below is Excel's own answer, not ours. If one fails, fix the compiler, never the number.
"""
import pytest
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute

# S!A1:A4 - 5, "5", apple, a formula that returns ""; S!H1 is an error; S!D1:D3 is sorted ascending
S = {
    "A1": 5, "A2": "5", "A3": "apple", "A4": ('=IF(1>0,"","x")', ""), "H1": "#N/A",
    "D1": 0, "D2": 100, "D3": 500, "E1": 500, "E2": 100, "E3": 0,
}


def run(formulas):
    book = book_of({"S": S, "F": {f"A{i}": (f, v) for i, (f, v) in enumerate(formulas, 1)}})
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def N(v):
    return ("n", float(v), None)


def S_(v):
    return ("s", None, v)


def test_countif_not_equal_counts_text_that_looks_like_the_number():
    # Excel: <>5 skips only the number 5; the text "5" is "not equal". =5 still matches both.
    got, c = run([('=COUNTIF(S!A1:A3,"<>5")', 2), ("=COUNTIF(S!A1:A3,5)", 2)])
    assert c.unresolved == {} and got[("F", 1, 1)] == N(2) and got[("F", 2, 1)] == N(2)


def test_countif_not_empty_counts_a_formula_that_returns_empty_text():
    got, _ = run([('=COUNTIF(S!A1:A4,"<>")', 4), ('=COUNTIF(S!A1:A5,"<>")', 4)])  # A5 is truly empty
    assert got[("F", 1, 1)] == N(4) and got[("F", 2, 1)] == N(4)


def test_days360_end_of_february_is_only_adjusted_when_the_start_is_too():
    got, _ = run([("=DAYS360(DATE(2026,1,30),DATE(2026,2,28))", 28),
                  ("=DAYS360(DATE(2026,2,28),DATE(2026,3,31))", 30),
                  ("=DAYS360(DATE(2026,2,28),DATE(2026,2,28))", 0), ("=DAYS360(DATE(2026,1,15),DATE(2026,4,30))", 105)])
    assert [got[("F", i, 1)] for i in (1, 2, 3, 4)] == [N(28), N(30), N(0), N(105)]


def test_text_of_a_tiny_negative_number_has_no_minus_sign():
    got, _ = run([('=TEXT(-0.001,"0.00")', "0.00"), ('=TEXT(-0.004,"0.0%")', "-0.4%"), ('=TEXT(-1.5,"0.00")', "-1.50")])
    assert [got[("F", i, 1)] for i in (1, 2, 3)] == [S_("0.00"), S_("-0.4%"), S_("-1.50")]


def test_mod_is_the_exact_remainder_with_the_divisors_sign():
    got, _ = run([("=MOD(1,0.1)", 0.09999999999999995), ("=MOD(-3,2)", 1), ("=MOD(3,-2)", -1), ("=MOD(5.5,2)", 1.5),
                  ("=MOD(6,3)", 0)])
    assert got[("F", 1, 1)][1] == 0.09999999999999995
    assert [got[("F", i, 1)] for i in (2, 3, 4, 5)] == [N(1), N(-1), N(1.5), N(0)]


def test_date_1900_counts_a_29_february_that_never_existed():
    got, _ = run([("=DATE(1900,2,29)", 60), ("=DATE(1900,3,1)", 61), ("=DATE(1900,2,28)", 59), ("=DATE(1900,1,31)", 31),
                  ("=DATE(1900,2,30)", 61), ("=DATE(1900,1,60)", 60), ("=DATE(1900,3,0)", 60)])
    assert [got[("F", i, 1)] for i in range(1, 8)] == [N(60), N(61), N(59), N(31), N(61), N(60), N(60)]


@pytest.mark.parametrize("formula, want", [
    ('="x"&0.00001', "x0.00001"), ('="x"&1000000000000000', "x1000000000000000"),
    ('="x"&123456789012345000', "x123456789012345000"), ('="x"&-0.5', "x-0.5"), ('="x"&100', "x100"),
    ('="x"&123456789.123456789', "x123456789.123457"), ('="x"&(0.1+0.2)', "x0.3"), ('="x"&0', "x0"),
    ('="x"&0.000001', "x0.000001"), ('="x"&0.00000000001', "x0.00000000001"),
    ('="x"&1E+18', "x1000000000000000000"), ('="x"&1E+19', "x10000000000000000000"),
    ('="x"&1E+20', "x1E+20"), ('="x"&1E+22', "x1E+22"),
])
def test_number_to_text_is_plain_digits_in_the_range_excel_showed(formula, want):
    got, _ = run([(formula, want)])
    assert got[("F", 1, 1)] == S_(want)


def test_a_typed_in_error_is_reported_before_an_error_inside_a_referenced_cell():
    # Excel: MAX(H3, 1/0) with H3 = #N/A answers #DIV/0!. The expression 1/0 is evaluated before MAX reads H3.
    # Seen in real Excel for MAX, MIN, AVERAGE, PRODUCT and SUM (cells G35, G57-G62 of examples/functions.xlsx).
    got, _ = run([("=MAX(S!H1,1/0)", "#DIV/0!"), ("=MAX(1/0,S!H1)", "#DIV/0!"), ("=MAX(S!H1)", "#N/A"),
                  ("=SUM(S!H1,1/0)", "#DIV/0!"), ("=SUM(S!H1:H2,1/0)", "#DIV/0!"), ("=MIN(S!H1,1/0)", "#DIV/0!"),
                  ("=AVERAGE(S!H1,1/0)", "#DIV/0!"), ("=PRODUCT(S!H1,1/0)", "#DIV/0!")])
    assert [got[("F", i, 1)] for i in range(1, 9)] == [("e", None, "#DIV/0!")] * 2 + [("e", None, "#N/A")] +         [("e", None, "#DIV/0!")] * 5


def test_approximate_match_on_unsorted_data_is_unresolved_not_guessed():
    # Excel's MATCH(150, {0,100,500}, -1) is #N/A: it binary-searches and assumes descending data.
    _, c = run([("=MATCH(150,S!D1:D3,-1)", "#N/A"), ("=MATCH(150,S!E1:E3,-1)", 1), ("=MATCH(150,S!D1:D3,1)", 2),
                ("=MATCH(150,S!E1:E3,1)", 3)])
    assert "sorted" in c.unresolved[("F", 1, 1)] and "sorted" in c.unresolved[("F", 4, 1)]
    assert ("F", 2, 1) not in c.unresolved and ("F", 3, 1) not in c.unresolved


def test_a_negative_number_to_the_power_of_one_over_an_odd_number_is_the_real_root():
    # Real Excel: (-8)^(1/3) is -1.9999999999999998. Other fractional powers of a negative number are #NUM!.
    got, _ = run([("=(-8)^(1/3)", -1.9999999999999998), ("=(-8)^0.5", "#NUM!"), ("=(-32)^(1/5)", -2)])
    assert got[("F", 1, 1)][1] == pytest.approx(-2.0, rel=1e-9) and got[("F", 3, 1)][1] == pytest.approx(-2.0, rel=1e-9)
    assert got[("F", 2, 1)] == ("e", None, "#NUM!")
