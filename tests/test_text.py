from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute

CASES = [  # formula, Excel's answer (Microsoft's function reference)
    ('=CONCATENATE("a", 1, TRUE)', "a1TRUE"), ("=CONCAT(S!B1:B3)", "x2.5TRUE"), ('=TEXTJOIN("-", TRUE, S!C1:C3, "z")', "p-q-z"),
    ('=LEFT("Excel", 2)', "Ex"), ('=LEFT("Excel")', "E"), ('=RIGHT("Excel", 3)', "cel"), ('=MID("Excel", 2, 3)', "xce"),
    ('=MID("Excel", 0, 1)', "#VALUE!"), ('=LEN("héllo")', 5), ('=TRIM("  a   b  ")', "a b"), ('=UPPER("aB")', "AB"),
    ('=PROPER("hELLO o\'neil 2nd")', "Hello O'Neil 2Nd"), ('=SUBSTITUTE("a-b-c", "-", "+")', "a+b+c"),
    ('=SUBSTITUTE("a-b-c", "-", "+", 2)', "a-b+c"), ('=SUBSTITUTE("a-b", "-", "+", 5)', "a-b"),
    ('=REPLACE("abcdef", 2, 3, "XY")', "aXYef"), ('=FIND("c", "abcabc")', 3), ('=FIND("C", "abcabc")', "#VALUE!"),
    ('=FIND("c", "abcabc", 4)', 6), ('=SEARCH("C", "abcabc")', 3), ('=SEARCH("b?c", "xaBxCd")', 3),
    ('=SEARCH("~?", "a?b")', 2), ('=REPT("ab", 3)', "ababab"), ('=VALUE(" 12.5 ")', 12.5), ('=VALUE("x")', "#VALUE!"),
    ('=EXACT("a", "A")', False), ("=CHAR(65)", "A"), ("=CHAR(128)", "€"), ("=CHAR(0)", "#VALUE!"),
    ('=CLEAN("a"&CHAR(9)&"b")', "ab"), ('=TEXT(1234.567, "#,##0.00")', "1,234.57"), ('=TEXT(2.675, "0.00")', "2.68"),
    ('=TEXT(0.256, "0.0%")', "25.6%"), ('=TEXT(-1234.5, "#,##0")', "-1,235"), ('=TEXT("abc", "0.00")', "abc"),
]


def test_text_functions():
    book = book_of({"S": {"B1": "x", "B2": 2.5, "B3": True, "C1": "p", "C3": "q"},
                    "F": {f"A{i}": (f, v) for i, (f, v) in enumerate(CASES, 1)}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert compiled.unresolved == {}
    for i, (formula, want) in enumerate(CASES, 1):
        kind, n, t = got[("F", i, 1)]
        if isinstance(want, bool):
            assert (kind, n) == ("b", float(want)), formula
        elif isinstance(want, (int, float)):
            assert (kind, n) == ("n", float(want)), formula
        elif want.startswith("#"):
            assert (kind, t) == ("e", want), formula
        else:
            assert (kind, t) == ("s", want), (formula, t)


def test_unsupported_text_forms_say_why():
    book = book_of({"S": {"A1": "a", "A3": "c", "B1": ('=TEXTJOIN(",", FALSE, A1:A3)', "a,,c"),
                          "B2": ('=TEXT(1, "0.0E+00")', "1.0E+00")}})
    unresolved = compile_book(book).unresolved
    assert unresolved[("S", 1, 2)] == "TEXTJOIN keeping the empty cells of a range not supported yet"
    assert unresolved[("S", 2, 2)] == "TEXT format '0.0E+00' not supported yet"
