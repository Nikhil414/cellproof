from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute

CASES = [  # formula, Excel's answer (Microsoft's function reference)
    ('=CONCATENATE("a", 1, TRUE)', "a1TRUE"), ("=CONCAT(S!B1:B3)", "x2.5TRUE"), ('=TEXTJOIN("-", TRUE, S!C1:C3, "z")', "p-q-z"),
    ('=LEFT("Excel", 2)', "Ex"), ('=LEFT("Excel")', "E"), ('=RIGHT("Excel", 3)', "cel"), ('=MID("Excel", 2, 3)', "xce"),
    ('=MID("Excel", 0, 1)', "#VALUE!"), ('=LEN("héllo")', 5), ('=TRIM("  a   b  ")', "a b"), ('=UPPER("aB")', "AB"),
    ('=PROPER("hELLO o\'neil 2nd")', "Hello O'Neil 2Nd"), ('=PROPER(S!B1&" éa ÉB")', "X Éa Éb"),
    ('=PROPER("")', ""),  # DuckDB 2.0.1: string_split('', '') is [] not [''], so an empty text became NULL
    ('=SUBSTITUTE("a-b-c", "-", "+")', "a+b+c"),
    ('=SUBSTITUTE("a-b-c", "-", "+", 2)', "a-b+c"), ('=SUBSTITUTE("a-b", "-", "+", 5)', "a-b"),
    ('=REPLACE("abcdef", 2, 3, "XY")', "aXYef"), ('=FIND("c", "abcabc")', 3), ('=FIND("C", "abcabc")', "#VALUE!"),
    ('=FIND("c", "abcabc", 4)', 6), ('=SEARCH("C", "abcabc")', 3), ('=SEARCH("b?c", "xaBxCd")', 3),
    ('=SEARCH("~?", "a?b")', 2), ('=REPT("ab", 3)', "ababab"), ('=VALUE(" 12.5 ")', 12.5), ('=VALUE("x")', "#VALUE!"),
    ('=EXACT("a", "A")', False), ("=CHAR(65)", "A"), ("=CHAR(128)", "€"), ("=CHAR(0)", "#VALUE!"),
    ('=CLEAN("a"&CHAR(9)&"b")', "ab"), ('=TEXT(1234.567, "#,##0.00")', "1,234.57"), ('=TEXT(2.675, "0.00")', "2.68"),
    ('=TEXT(0, "0.00")', "0.00"),  # the literal-zero shape: the compile-time constant must not raise
    ('=TEXT(1.005, "0.00")', "1.01"),  # 15 digits then half away from zero, as ROUND(1.005,2); round(x,2) gave 1.00
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
    # A bad places fires #NUM! even when the number itself is text (BASES.xlsx places!G38:G49, Excel-saved):
    ('=DEC2BIN("E", 0)', "#NUM!"), ('=DEC2OCT("E", 0)', "#NUM!"), ('=DEC2HEX("E", 0)', "#NUM!"),
    ('=DEC2BIN("E", 11)', "#NUM!"), ('=DEC2OCT("E", 11)', "#NUM!"), ('=DEC2HEX("E", 11)', "#NUM!"),
    # TRUE/FALSE is #VALUE! as the number (test.xlsx J7, J17:J19, J27, J55, Excel-saved):
    ('=DEC2BIN(TRUE)', "#VALUE!"), ('=DEC2HEX(TRUE)', "#VALUE!"), ('=DEC2OCT(TRUE)', "#VALUE!"),
    ('=BIN2DEC(TRUE)', "#VALUE!"), ('=HEX2DEC(TRUE)', "#VALUE!"), ('=OCT2DEC(TRUE)', "#VALUE!"),
    ('=DEC2BIN(S!B3)', "#VALUE!"), ('=DEC2HEX(S!B3)', "#VALUE!"), ('=BIN2DEC(S!B3)', "#VALUE!"),
    # The cross-pairs (test.xlsx ENGINEERING!R26:AF26 etc., Excel-saved): the X2DEC parse feeds the DEC2* format,
    # and the target base's own two's-complement range still gates it (HEX2BIN("400") is #NUM! at 1024 > 511):
    ('=HEX2BIN("B")', "1011"), ('=HEX2BIN("E", 4)', "1110"), ('=HEX2BIN("0")', "0"), ('=HEX2BIN(11)', "10001"),
    ('=HEX2BIN(2.3)', "#NUM!"), ('=HEX2BIN(-1)', "#NUM!"), ('=HEX2BIN("400")', "#NUM!"),
    ('=HEX2BIN("FFFFFFFFFF")', "1111111111"), ('=HEX2BIN("FFFFFFFFFE")', "1111111110"),
    ('=HEX2BIN("FFFFFFFFFF", 3)', "1111111111"), ('=HEX2BIN(9999999999)', "#NUM!"),
    ('=HEX2BIN("11111111111")', "#NUM!"), ('=HEX2BIN("GGG")', "#NUM!"), ('=HEX2BIN("B", 1)', "#NUM!"),
    ('=HEX2OCT("F", 3)', "017"), ('=HEX2OCT("FFFFFFFFFE")', "7777777776"), ('=HEX2OCT("FFFFFFFFFF")', "7777777777"),
    ('=HEX2OCT("2000000000")', "#NUM!"), ('=HEX2OCT(9999999999)', "#NUM!"),
    ('=BIN2HEX("111111111")', "1FF"), ('=BIN2HEX("1000000000")', "FFFFFFFE00"), ('=BIN2HEX(111111111)', "1FF"),
    ('=BIN2HEX(0, 0)', "#NUM!"), ('=BIN2HEX(2.3)', "#NUM!"), ('=BIN2HEX("FFFFFFFFFE")', "#NUM!"),
    ('=BIN2HEX("11111111111")', "#NUM!"),
    ('=BIN2OCT("111111111")', "777"), ('=BIN2OCT("1000000000")', "7777777000"), ('=BIN2OCT("11", 3)', "003"),
    ('=BIN2OCT("2")', "#NUM!"),
    ('=OCT2HEX("7777777777")', "FFFFFFFFFF"), ('=OCT2HEX("54")', "2C"), ('=OCT2HEX(54)', "2C"),
    ('=OCT2HEX("4000000000")', "FFE0000000"), ('=OCT2HEX("77777777770")', "#NUM!"), ('=OCT2HEX("a")', "#NUM!"),
    ('=OCT2BIN("7777777777")', "1111111111"), ('=OCT2BIN("777")', "111111111"), ('=OCT2BIN("1001")', "#NUM!"),
    ('=OCT2BIN("a")', "#NUM!"), ('=OCT2BIN("7", 6)', "000111"),
    # TRUE is #VALUE! in the cross-pairs too (test.xlsx J26:J28, J8:J9, J54:J56, Excel-saved):
    ('=HEX2BIN(TRUE)', "#VALUE!"), ('=BIN2HEX(S!B3)', "#VALUE!"),
    ('=HEX2BIN(1/0)', "#DIV/0!"), ('=HEX2BIN(" 1")', "#NUM!"),
    # an empty text stays UNCHECKED, as in the X2DEC family (never verified in real Excel):
    ('=HEX2BIN("")', "#UNCHECKED"),
]


def test_a_text_format_rounds_like_round_does():
    # both must say 1.01 and 0.29: TEXT rounds the 15-significant-digit value, as ROUND does, not the raw double
    book = book_of({"S": {"A1": 1.005, "A2": 0.285, "A3": 2.675,
                          "B1": ('=TEXT(A1,"0.00")', "1.01"), "B2": ('=TEXT(A2,"0.00")', "0.29"),
                          "B3": ('=TEXT(A3,"0.00")', "2.68"),
                          "C1": ("=ROUND(A1,2)", 1.01), "C2": ("=ROUND(A2,2)", 0.29), "C3": ("=ROUND(A3,2)", 2.68)}})
    got = execute(book, compile_book(book))
    for row in (1, 2, 3):
        assert got[("S", row, 2)][2] == f"{got[('S', row, 3)][1]:.2f}", row


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


def test_a_huge_value_keeps_its_digits_and_does_not_lose_the_batch():
    # TEXT scales the value by 10^decimals before rounding, and Postgres raises on float8 overflow where DuckDB
    # returns inf: one such cell used to make Postgres refuse the whole statement, turning every cell in the batch
    # UNRESOLVED. A value too large to scale keeps the digits a plain format showed before (its own FAIL, not a
    # dead batch); a percent has no answer without scaling, so it is NULL and FAILs visibly.
    book = book_of({"S": {"A1": 1e308, "A2": 1e-320, "A3": 2.675,
                          "B1": ('=TEXT(A1,"0.00")', "1e308"), "B2": ('=TEXT(A2,"0.00")', "0.00"),
                          "B3": ('=TEXT(A3,"0.00")', "2.68"), "B4": ('=TEXT(A1,"0%")', "1e308%")}})
    got = execute(book, compile_book(book))
    assert str(got[("S", 1, 2)][2]).startswith("10000000000000000")   # unscaled digits, as 1.1.1 had them
    assert got[("S", 2, 2)][2] == "0.00"
    assert got[("S", 3, 2)][2] == "2.68"
    assert got[("S", 4, 2)][2] is None                                # no percent of a value that cannot be scaled


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


def test_base_functions_of_a_range_in_a_cse_block():
    # Real Excel saved test.xlsx (an ATP-era build): {=HEX2DEC(C27:J27)} is #VALUE! in every cell of the block -
    # the function does not broadcast over a range there. DEC2* of a range was never checked in real Excel, so it
    # stays honest instead of broadcasting like modern Excel would.
    from cellproof.reader import load
    from xlsxkit import write_xlsx
    import tempfile
    from pathlib import Path
    with tempfile.TemporaryDirectory() as tmp:
        wb = Path(tmp) / "b.xlsx"
        write_xlsx(wb, {"S": {"C1": 0, "D1": 9999999999, "E1": 2.3, "F1": 11111111111,
                              "G1": -1, "H1": "f000000000", "I1": True,          # J1 is left empty, as in test.xlsx
                              "A3": ("{=HEX2DEC(C1:J1)}", "#VALUE!", "A3:H3"),
                              "B3": "#VALUE!", "C3": "#VALUE!", "D3": "#VALUE!", "E3": "#VALUE!",
                              "F3": "#VALUE!", "G3": "#VALUE!", "H3": "#VALUE!",
                              "J3": ("{=DEC2HEX(C1:J1, 4)}", "#VALUE!", "J3:Q3"),  # saved value is a placeholder: the real
                                                                                  # Excel answer was never recorded, the
                                                                                  # cell is UNRESOLVED either way,
                              "K3": "", "L3": "", "M3": "", "N3": "", "O3": "", "P3": "", "Q3": ""}})
        book = load(wb)
        compiled = compile_book(book)
        got = execute(book, compiled)
        assert [got[("S", 3, c)] for c in range(1, 9)] == [("e", None, "#VALUE!")] * 8
        assert "DEC2HEX of a range in an array formula not supported yet" in compiled.unresolved[("S", 3, 10)]


def test_base_cross_pairs_of_a_blank_cell():
    # Real Excel saved test.xlsx (H26, H28, H54, H56): HEX2BIN of a blank cell is HEX2BIN(0) -> '0', and the
    # places check still gates it (W8, W9: BIN2HEX of a blank cell with places 0 is #NUM!).
    book = book_of({"S": {"A1": None, "B1": ("=HEX2BIN(A1)", "0"), "B2": ("=OCT2HEX(A1)", "0"),
                          "B3": ("=BIN2HEX(A1, 0)", "#NUM!"), "B4": ("=HEX2BIN(A1, 4)", "0000")}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert compiled.unresolved == {}
    assert got[("S", 1, 2)] == ("s", None, "0") and got[("S", 2, 2)] == ("s", None, "0")
    assert got[("S", 3, 2)] == ("e", None, "#NUM!") and got[("S", 4, 2)] == ("s", None, "0000")


def test_base_cross_pairs_of_a_range_in_a_cse_block():
    # Real Excel refused a range in a CSE block with #VALUE! for the X2DEC family (test.xlsx, proven in PR #12) -
    # but no real-Excel answer was ever recorded for the cross-pairs, so they stay honest UNRESOLVED instead of
    # claiming either the #VALUE! or modern Excel's broadcast.
    from cellproof.reader import load
    from xlsxkit import write_xlsx
    import tempfile
    from pathlib import Path
    with tempfile.TemporaryDirectory() as tmp:
        wb = Path(tmp) / "b.xlsx"
        write_xlsx(wb, {"S": {"C1": "F", "D1": "7", "E1": "1001",                      # J1 left empty, as in test.xlsx
                              "A3": ("{=HEX2BIN(C1:J1)}", "x", "A3:H3"),              # saved value is a placeholder
                              "B3": "", "C3": "", "D3": "", "E3": "", "F3": "", "G3": "", "H3": ""}})
        book = load(wb)
        compiled = compile_book(book)
        assert "HEX2BIN of a range in an array formula not supported yet" in compiled.unresolved[("S", 3, 1)]


