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
