from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute

CASES = [  # formula, Excel's answer (serial numbers: 46023 = 2026-01-01)
    ("=DATE(2026, 1, 1)", 46023), ("=DATE(2026, 14, 31)", 46449), ("=DATE(26, 1, 1)", 9498), ("=DATE(2026, 2, 29)", 46082),
    ("=DATE(-1, 1, 1)", "#NUM!"), ("=YEAR(46023)", 2026), ("=MONTH(46023.75)", 1), ("=DAY(46081)", 28),
    ("=MONTH(60)", 2), ("=DAY(60)", 29), ("=DAY(0)", 0), ("=YEAR(-1)", "#NUM!"), ("=WEEKDAY(46023)", 5),
    ("=WEEKDAY(46023, 2)", 4), ("=WEEKDAY(46023, 3)", 3), ("=WEEKDAY(1)", 1), ("=DAYS(46081, 46023)", 58),
    ("=EDATE(46053, 1)", 46081), ("=EOMONTH(46023, 1)", 46081), ("=EOMONTH(46023, -1)", 46022),
    ("=DAYS360(DATE(2011,1,30), DATE(2011,12,31))", 330), ("=DAYS360(DATE(2011,1,31), DATE(2011,3,31), TRUE)", 60),
    ("=DAYS360(DATE(2026,1,15), DATE(2026,3,15))", 60), ('=DATEDIF(DATE(2020,5,10), DATE(2026,5,9), "Y")', 5),
    ('=DATEDIF(DATE(2026,1,31), DATE(2026,2,28), "M")', 0), ('=DATEDIF(46023, 46081, "D")', 58),
    ('=DATEDIF(46081, 46023, "D")', "#NUM!"), ('=TEXT(46023, "yyyy-mm-dd")', "2026-01-01"),
    ('=TEXT(46023, "dddd d mmmm")', None), ('=TEXT(46087, "mmm yyyy")', "Mar 2026"), ('=TEXT(46087, "dd/mm/yyyy")', "06/03/2026"),
]


def test_date_functions():
    cases = [(f, v) for f, v in CASES if v is not None]
    book = book_of({"S": {f"A{i}": (f, v) for i, (f, v) in enumerate(cases, 1)}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert compiled.unresolved == {}
    for i, (formula, want) in enumerate(cases, 1):
        kind, n, t = got[("S", i, 1)]
        if isinstance(want, str) and want.startswith("#"):
            assert (kind, t) == ("e", want), formula
        elif isinstance(want, str):
            assert (kind, t) == ("s", want), (formula, t)
        else:
            assert (kind, n) == ("n", float(want)), (formula, n)


def test_1904_workbooks_and_odd_forms_say_why():
    book = book_of({"S": {"A1": ("=YEAR(1)", 1904), "A2": ('=TEXT(1, "yyyy")', "1904")}})
    book.date1904 = True
    assert set(compile_book(book).unresolved.values()) == {"1904 date system not supported yet"}
    other = book_of({"S": {"A1": ('=DATEDIF(1, 2, "MD")', 1), "A2": ("=WEEKDAY(1, 11)", 7)}})
    assert compile_book(other).unresolved == {
        ("S", 1, 1): 'DATEDIF unit must be "Y", "M" or "D" written in the formula',
        ("S", 2, 1): "WEEKDAY return type must be 1, 2 or 3"}
