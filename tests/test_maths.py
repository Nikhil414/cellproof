import math

from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute

CASES = [  # formula, Excel's answer (from Microsoft's function reference)
    ("=ABS(-2.5)", 2.5), ("=INT(-2.5)", -3), ("=SIGN(-0.1)", -1), ("=PI()", math.pi), ("=SQRT(16)", 4),
    ("=SQRT(-1)", "#NUM!"), ("=EXP(1)", math.e), ("=EXP(1000)", "#NUM!"), ("=LN(0)", "#NUM!"), ("=LOG10(1000)", 3),
    ("=LOG(8, 2)", 3), ("=LOG(100)", 2), ("=LOG(8, 1)", "#DIV/0!"), ("=POWER(2, 10)", 1024), ("=MOD(-7, 3)", 2),
    ("=MOD(7, -3)", -2), ("=MOD(5, 0)", "#DIV/0!"), ("=TRUNC(-8.9)", -8), ("=TRUNC(3.14159, 2)", 3.14),
    ("=ROUNDUP(0.1*3, 1)", 0.3), ("=ROUNDUP(-3.14159, 1)", -3.2), ("=ROUNDDOWN(3.14159, 3)", 3.141),
    ("=ROUNDUP(31415.9, -2)", 31500), ("=CEILING(2.5, 1)", 3), ("=CEILING(-2.5, -2)", -4), ("=CEILING(-2.5, 2)", -2),
    ("=CEILING(1.5, -1)", "#NUM!"), ("=FLOOR(-2.5, -2)", -2), ("=FLOOR(-2.5, 2)", -4), ("=FLOOR(1.5, 0)", "#DIV/0!"),
    ("=MROUND(10, 3)", 9), ("=MROUND(-10, -3)", -9), ("=MROUND(5, -2)", "#NUM!"), ("=MROUND(1.3, 0.2)", 1.4),
]


def test_math_functions_and_their_errors():
    book = book_of({"S": {f"A{i}": (f, v) for i, (f, v) in enumerate(CASES, 1)}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert compiled.unresolved == {}
    for i, (formula, want) in enumerate(CASES, 1):
        kind, n, t = got[("S", i, 1)]
        if isinstance(want, str):
            assert (kind, t) == ("e", want), formula
        else:
            assert kind == "n" and math.isclose(n, want, rel_tol=1e-12), (formula, n)


def test_round_rounds_the_15_digit_value_excel_sees():
    # Microsoft's ROUND examples, and ROUND(1.005, 2): Excel 1.01 (the double is 1.00499999..., which binary rounding
    # took to 1.0 before 0.6.1); the digits may come from a cell
    cases = [("=ROUND(2.15,1)", 2.2), ("=ROUND(2.149,1)", 2.1), ("=ROUND(-1.475,2)", -1.48), ("=ROUND(21.5,-1)", 20),
             ("=ROUND(626.3,-3)", 1000), ("=ROUND(1.98,-1)", 0), ("=ROUND(-50.55,-2)", -100), ("=ROUND(1.005,2)", 1.01),
             ("=ROUND(2.675,2)", 2.68), ("=ROUND(2.5,0)", 3), ("=ROUND(-2.5,0)", -3), ("=ROUND(1E+300,10)", 1e300),
             ("=ROUND(5,-400)", 0), ("=ROUND(1.005,B1)", 1.01), ("=ROUND(0.1+0.2,15)", 0.3)]
    book = book_of({"S": {"B1": 2.0, **{f"A{i}": (f, v) for i, (f, v) in enumerate(cases, 1)}}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert compiled.unresolved == {}
    for i, (formula, want) in enumerate(cases, 1):
        kind, n, _ = got[("S", i, 1)]
        assert kind == "n" and math.isclose(n, want, rel_tol=1e-12, abs_tol=1e-300), (formula, n)
