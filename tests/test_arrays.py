"""Array constants ({1,2;3,4}) behave like a range of those values wherever a range is taken."""
from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.parser import parse, refs
from cellproof.verifier import execute


def run(formulas, data=None):
    book = book_of({"S": formulas} | ({"D": data} if data else {}))
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def N(v):
    return ("n", float(v), None)


def test_array_constants_where_a_range_is_taken():
    got, c = run({
        "A1": ("=SUM({1,2;3,-4})", 2),
        "A2": ('=SUM({1,"2",TRUE})', 1),                    # like a range: text and TRUE skipped
        "A3": ('=MATCH("b",{"a","b","c"},0)', 2),
        "A4": ('=INDEX({10,20;30,40},2,1)', 30),
        "A5": ('=LOOKUP(75,{0,50,100},{"F","C","A"})', "C"),
        "A6": ('=VLOOKUP(2,{1,"x";2,"y"},2,FALSE)', "y"),
        "A7": ("=COUNTA({1,#N/A,\"x\"})", 3),
        "A8": ("=MAX({3,9,1})+1", 10),
        "A9": ("=SUMPRODUCT({1,2,3},{4,5,6})", 32),
        "A10": ("={7}*2", 14),                              # one value: just that value
    })
    assert c.unresolved == {}
    assert [got[("S", r, 1)] for r in range(1, 11)] == [
        N(2), N(1), N(2), N(30), ("s", None, "C"), ("s", None, "y"), N(3), N(10), N(32), N(14)]


def test_array_constants_where_one_value_is_expected_are_not_guessed():
    _, c = run({"A1": ("=A2*{1,2}", 0), "A2": 3, "A3": ('=IF({1,0},"a","b")', "a"), "A4": ("={1,2;3}", 0)})
    assert "array constant" in c.unresolved[("S", 1, 1)]
    assert "array constant" in c.unresolved[("S", 3, 1)]
    assert "rows of different lengths" in c.unresolved[("S", 4, 1)]


def test_array_constants_are_not_dependencies():
    assert list(refs(parse('SUM(A1,{1,2})'))) == ["A1"]


def test_element_wise_arrays_in_aggregates():
    # Modern Excel: a range arithmetic as an aggregate's argument evaluates element by element
    # (SUM(D!A1:A3*D!B1:B3) is the sum of the products, not the intersected first row).
    data = {"A1": 1, "B1": 10, "A2": 2, "B2": 20, "A3": 3, "B3": 30, "A4": 4, "B4": 40, "A5": 5, "B5": 50,
            "C1": 6, "C2": 1, "C3": 8, "C4": 2, "C5": 1, "D1": 3, "D3": 7}
    got, c = run({
        "E1": ("=SUM(D!A1:A5*D!B1:B5)", 550),
        "E2": ("=SUM(D!A1:A5*D!C1)", 90),               # a scalar is one element for every row: (1+..+5)*6
        "E3": ("=MAX(D!A1:A5*D!B1:B5)", 250),
        "E4": ("=MIN(D!A1:A5*D!B1:B5)", 10),
        "E5": ("=AVERAGE(D!A1:A3*D!B1:B3)", 140 / 3),
        "E6": ("=COUNT(D!A1:A5*D!B1:B5)", 5),
        "E7": ("=PRODUCT(D!A1:A3*D!B1:B3)", 36000),
        "E8": ("=SUMPRODUCT((D!C1:C5>5)*D!A1:A5)", 4),  # TRUE,_,TRUE,_,_ times 1..5
        "E9": ("=SUM((D!C1:C5>2)*1)", 2),
        "E10": ("=COUNT(D!C1:C5>5)", 0),              # COUNT skips TRUE/FALSE in an array, like in a range
        "E11": ("=COUNTA(D!C1:C5>5)", 5),
        "E12": ("=SUM(D!A1:A5*D!D1:D5)", 24),           # empty cells are 0 elements: 3 + 0 + 21 + 0 + 0
        "E13": ("=SUM(D!A1:A3+D!B1:B3)", 66),
    }, data)
    assert c.unresolved == {}
    assert [got[("S", r, 5)] for r in range(1, 14)] == [
        N(550), N(90), N(250), N(10), N(140 / 3), N(5), N(36000), N(4), N(2), N(0), N(5), N(24), N(66)]


def test_array_elements_carry_excel_s_errors():
    data = {"A1": 1, "B1": 2, "A2": "x", "B2": 3, "A3": 3, "B3": 4, "H1": "#N/A", "I1": 2, "I2": 4}
    got, c = run({
        "E1": ("=SUM(D!A1:A3*D!B1:B3)", "#VALUE!"),     # text in the arithmetic: that element is #VALUE!
        "E2": ("=SUM(D!H1:H2*D!I1:I2)", "#N/A"),        # an error element is the result
        "E3": ("=SUM(D!A1:A4*D!B1:B3)", "#VALUE!"),     # arrays of different sizes
        "E4": ("=SUM(D!A1:A3/0)", "#DIV/0!"),         # the guard applies per element
    }, data)
    assert c.unresolved == {}
    assert [got[("S", r, 5)] for r in range(1, 5)] == [
        ("e", None, "#VALUE!"), ("e", None, "#N/A"), ("e", None, "#VALUE!"), ("e", None, "#DIV/0!")]


def test_single_cell_and_block_array_formulas_compile():
    from cellproof.reader import load
    from xlsxkit import write_xlsx
    import tempfile
    from pathlib import Path
    with tempfile.TemporaryDirectory() as tmp:
        wb = Path(tmp) / "b.xlsx"
        write_xlsx(wb, {"S": {"A1": 1, "B1": 10, "A2": 2, "B2": 20, "A3": 3, "B3": 30,
                              "F1": ("{=SUM(A1:A3*B1:B3)}", 140),                 # one cell: evaluated as a whole
                              "G1": ("{=A1:A3*B1:B3}", 10, "G1:G3"), "G2": 40, "G3": 90}})  # a block, per cell
        book = load(wb)
        compiled = compile_book(book)
        assert compiled.unresolved == {}
        got = execute(book, compiled)
        assert got[("S", 1, 6)] == N(140)
        assert [got[("S", r, 7)] for r in (1, 2, 3)] == [N(10), N(40), N(90)]


def test_if_over_ranges_in_aggregates():
    # The classic conditional aggregate: MAX(IF(A="x", B)) entered as an array formula.
    data = {"A1": "x", "B1": 10, "A2": "y", "B2": 20, "A3": "x", "B3": 30, "A4": "y", "B4": 40, "A5": "x", "B5": 50,
            "C1": 6, "C2": 1, "C3": 8, "C4": 2, "C5": 1, "H1": "#N/A", "H2": 9}
    got, c = run({
        "E1": ('=MAX(IF(D!A1:A5="x", D!B1:B5))', 50),
        "E2": ('=SUM(IF(D!C1:C5>2, D!B1:B5, 0))', 40),      # 10 + 0 + 30 + 0 + 0
        "E3": ('=COUNT(IF(D!A1:A5="x", D!B1:B5))', 3),      # FALSE elements are not numbers
        "E4": ('=MAX(IF(D!A1:A5="y", D!B1:B5))', 40),
        "E5": ('=MAX(IF(D!A1:A5="x", "a", "b"))', 0),       # MAX skips text, like a range of text
        "E6": ('=SUM(IF(D!B1:B5, 1, 0))', 5),               # numbers as conditions: every one is TRUE
        "E7": ('=MAX(IF(D!A1:A5="X", D!B1:B5))', 50),       # text compares case-insensitively
        "E8": ('=SUM(IF(D!A1:A3="x", D!H1:H3, D!C1:C3))', "#N/A"),  # taken branch's error wins: #N/A, 1, blank
    }, data)
    assert c.unresolved == {}
    assert [got[("S", r, 5)] for r in range(1, 9)] == [
        N(50), N(40), N(3), N(40), N(0), N(5), N(50), ("e", None, "#N/A")]


def test_if_arrays_refuse_what_excel_answers_differently():
    _, c = run({"E1": ('=MAX(IF(D!A1:A5="x", D!B1:B4))', 0)},
               {"A1": "x", "A2": "y", "B1": 1, "B2": 2, "B3": 3, "B4": 4})
    assert "different sizes" in c.unresolved[("S", 1, 5)]


def test_text_comparisons_work_element_wise():
    data = {"A1": "a", "A2": "b", "A3": 3}
    got, c = run({
        "E1": ('=SUMPRODUCT((D!A1:A3="a")*1)', 1),
        "E2": ('=SUMPRODUCT((D!A1:A3="A")*1)', 1),          # case-insensitive
        "E3": ('=SUMPRODUCT((D!A1:A3>1)*1)', 3),            # text outranks every number: all three are TRUE
        "E4": ('=SUM(D!A1:A3*2)', "#VALUE!"),               # text that is not a number, in arithmetic
        "E5": ('=SUM(IF(D!A1:A3="b", 1, 0))', 1),
    }, data)
    assert c.unresolved == {}
    assert [got[("S", r, 5)] for r in range(1, 6)] == [N(1), N(1), N(3), ("e", None, "#VALUE!"), N(1)]


def test_iferror_and_ifna_over_arrays():
    data = {"A1": 10, "B1": 2, "A2": 20, "B2": 0, "A3": 30, "B3": 5, "H1": "#N/A", "H2": 2, "H3": 3}
    got, c = run({
        "E1": ("=SUM(IFERROR(D!A1:A3/D!B1:B3, 0))", 11),    # 5 + (guarded) + 6
        "E2": ('=COUNTA(IFERROR(D!A1:A3/D!B1:B3, "none"))', 3),
        "E3": ("=SUM(IFNA(D!H1:H3, 0))", 5),                # only #N/A is replaced: 0 + 2 + 3
        "E4": ("=SUM(IFERROR(D!H1:H3, 0))", 5),
        "E5": ("=SUM(IFERROR(D!H1:H3/0, -1))", -3),         # every element errors, every one replaced
    }, data)
    assert c.unresolved == {}
    assert [got[("S", r, 5)] for r in range(1, 6)] == [N(11), N(3), N(5), N(5), N(-3)]


def test_concat_and_power_over_arrays():
    data = {"A1": "a", "B1": 10, "A2": "b", "B2": 20, "A3": "c", "B3": 30, "C1": 8, "C2": 27, "C3": 64,
            "D1": -4}
    got, c = run({
        "E1": ('=SUMPRODUCT((D!A1:A3&D!B1:B3="a10")*1)', 1),     # numbers become text the Excel way
        "E2": ("=COUNTA(D!A1:A3&D!B1:B3)", 3),
        "E3": ("=SUM(D!C1:C3^2)", 4889),                          # 64 + 729 + 4096
        "E4": ("=SUM((-D!C1:C3)^(1/3))", -9),                     # Excel's real cube roots: -2 - 3 - 4
        "E5": ("=SUM(D!D1:D1^0.5)", "#NUM!"),                     # a negative base to a fractional power
    }, data)
    assert c.unresolved == {}
    assert [got[("S", r, 5)] for r in range(1, 6)] == [
        N(1), N(3), N(4889), N(-9), ("e", None, "#NUM!")]
