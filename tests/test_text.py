from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute

CASES = [  # formula, Excel's answer (Microsoft's function reference)
    ('=CONCATENATE("a", 1, TRUE)', "a1TRUE"), ("=CONCAT(S!B1:B3)", "x2.5TRUE"), ('=TEXTJOIN("-", TRUE, S!C1:C3, "z")', "p-q-z"),
    ('=LEFT("Excel", 2)', "Ex"), ('=LEFT("Excel")', "E"), ('=RIGHT("Excel", 3)', "cel"), ('=MID("Excel", 2, 3)', "xce"),
    ('=MID("Excel", 0, 1)', "#VALUE!"), ('=LEN("héllo")', 5), ('=TRIM("  a   b  ")', "a b"), ('=UPPER("aB")', "AB"),
    ('=PROPER("hELLO o\'neil 2nd")', "Hello O'Neil 2Nd"), ('=PROPER(S!B1&" éa ÉB")', "X Éa Éb"),
    ('=SUBSTITUTE("a-b-c", "-", "+")', "a+b+c"),
    ('=SUBSTITUTE("a-b-c", "-", "+", 2)', "a-b+c"), ('=SUBSTITUTE("a-b", "-", "+", 5)', "a-b"),
    ('=REPLACE("abcdef", 2, 3, "XY")', "aXYef"), ('=FIND("c", "abcabc")', 3), ('=FIND("C", "abcabc")', "#VALUE!"),
    ('=FIND("c", "abcabc", 4)', 6), ('=SEARCH("C", "abcabc")', 3), ('=SEARCH("b?c", "xaBxCd")', 3),
    ('=SEARCH("~?", "a?b")', 2), ('=REPT("ab", 3)', "ababab"), ('=VALUE(" 12.5 ")', 12.5), ('=VALUE("x")', "#VALUE!"),
    ('=EXACT("a", "A")', False), ("=CHAR(65)", "A"), ("=CHAR(128)", "€"), ("=CHAR(0)", "#VALUE!"),
    ('=CLEAN("a"&CHAR(9)&"b")', "ab"), ('=TEXT(1234.567, "#,##0.00")', "1,234.57"), ('=TEXT(2.675, "0.00")', "2.68"),
    ('=TEXT(0.256, "0.0%")', "25.6%"), ('=TEXT(-1234.5, "#,##0")', "-1,235"), ('=TEXT("abc", "0.00")', "abc"),
    ('=DEC2HEX(100, 4)', "0064"), ('=DEC2HEX(-54)', "FFFFFFFFCA"), ('=DEC2HEX(100)', "64"),
    ('=DEC2HEX(-1)', "FFFFFFFFFF"), ('=DEC2HEX(-549755813888)', "8000000000"),
    ('=DEC2HEX(549755813887)', "7FFFFFFFFF"), ('=DEC2HEX(549755813888)', "#NUM!"),
    ('=DEC2HEX(-549755813889)', "#NUM!"), ('=DEC2HEX(100, 1)', "#NUM!"), ('=DEC2HEX(100, 0)', "#NUM!"),
    ('=DEC2HEX(100, 11)', "#NUM!"), ('=DEC2HEX("abc")', "#VALUE!"), ('=DEC2HEX(100, "x")', "#VALUE!"),
    ('=DEC2HEX(-1, 4)', "FFFFFFFFFF"),
    ('=DEC2BIN(9)', "1001"), ('=DEC2BIN(9, 6)', "001001"), ('=DEC2BIN(-100)', "1110011100"),
    ('=DEC2BIN(-1)', "1111111111"), ('=DEC2BIN(-512)', "1000000000"), ('=DEC2BIN(511)', "111111111"),
    ('=DEC2BIN(512)', "#NUM!"), ('=DEC2BIN(-513)', "#NUM!"), ('=DEC2BIN(9, 2)', "#NUM!"),
    ('=DEC2BIN("abc")', "#VALUE!"),
    ('=DEC2OCT(58)', "72"), ('=DEC2OCT(58, 4)', "0072"), ('=DEC2OCT(-100)', "7777777634"),
    ('=DEC2OCT(-1)', "7777777777"), ('=DEC2OCT(-536870912)', "4000000000"),
    ('=DEC2OCT(536870911)', "3777777777"), ('=DEC2OCT(536870912)', "#NUM!"),
    ('=DEC2OCT(-536870913)', "#NUM!"), ('=DEC2OCT(58, 1)', "#NUM!"), ('=DEC2OCT("abc")', "#VALUE!"),
    ('=HEX2DEC("A5")', 165), ('=HEX2DEC("a5")', 165), ('=HEX2DEC("FFFFFFFFFF")', -1),
    ('=HEX2DEC("8000000000")', -549755813888), ('=HEX2DEC("7FFFFFFFFF")', 549755813887),
    ('=HEX2DEC("10000000000")', "#NUM!"), ('=HEX2DEC("XYZ")', "#NUM!"), ('=HEX2DEC(11)', 17),
    ('=BIN2DEC("1001")', 9), ('=BIN2DEC("1111111111")', -1), ('=BIN2DEC("1000000000")', -512),
    ('=BIN2DEC("0111111111")', 511), ('=BIN2DEC("11111111111")', "#NUM!"), ('=BIN2DEC("102")', "#NUM!"),
    ('=BIN2DEC(1001)', 9),
    ('=OCT2DEC("72")', 58), ('=OCT2DEC("7777777777")', -1), ('=OCT2DEC("4000000000")', -536870912),
    ('=OCT2DEC("3777777777")', 536870911), ('=OCT2DEC("77777777777")', "#NUM!"),
    ('=OCT2DEC("78")', "#NUM!"), ('=OCT2DEC(72)', 58),
    ('=HEX2DEC("")', "#UNCHECKED"), ('=BIN2DEC("")', "#UNCHECKED"), ('=OCT2DEC("")', "#UNCHECKED"),
    ('=DEC2HEX(-1,0)', "#UNCHECKED"), ('=DEC2BIN(-1,11)', "#UNCHECKED"), ('=DEC2OCT(-1,"x")', "#UNCHECKED"),
    ('=HEX2DEC(" A5")', "#NUM!"), ('=HEX2DEC("+A5")', "#NUM!"),
    ('=HEX2DEC(S!Z1)', "#UNCHECKED"), ('=DEC2HEX(-1,1/0)', "#DIV/0!"),
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


def test_substitute_with_an_instance_reads_cells():
    # DuckDB 2.0 refused list slicing over cell subqueries (Serialization Error), stopping the whole run (exit 3)
    from cellproof.compiler import compile_book
    from cellproof.verifier import execute
    from xlsxkit import book_of
    book = book_of({"S": {"A1": "a-b-c", "B1": "-", "C1": "+", "D1": 2.0, "D2": ("=D1+1", 3.0),
                          "E1": ("=SUBSTITUTE(A1,B1,C1,D1)", "a-b+c"), "E2": ("=SUBSTITUTE(A1,B1,C1,D2)", "a-b-c")}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert got[("S", 1, 5)] == ("s", None, "a-b+c") and got[("S", 2, 5)] == ("s", None, "a-b-c")


def test_base_conversions_with_cell_references_and_truncation():
    book = book_of({"S": {
        "A1": 10.9,
        "A2": -1.9,
        "A3": "FF",
        "A4": "1001",
        "A5": "72",
        "B1": ("=DEC2HEX(A1, 2)", "0A"),
        "B2": ("=DEC2HEX(A2)", "FFFFFFFFFF"),
        "B3": ("=HEX2DEC(A3)", 255.0),
        "B4": ("=BIN2DEC(A4)", 9.0),
        "B5": ("=OCT2DEC(A5)", 58.0),
        "B6": ("=DEC2BIN(9.8, 4.2)", "1001"),
        "B7": ("=DEC2OCT(58.7, 3.9)", "072"),
    }})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert got[("S", 1, 2)] == ("s", None, "0A")
    assert got[("S", 2, 2)] == ("s", None, "FFFFFFFFFF")
    assert got[("S", 3, 2)] == ("n", 255.0, None)
    assert got[("S", 4, 2)] == ("n", 9.0, None)
    assert got[("S", 5, 2)] == ("n", 58.0, None)
    assert got[("S", 6, 2)] == ("s", None, "1001")
    assert got[("S", 7, 2)] == ("s", None, "072")

