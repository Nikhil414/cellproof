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


def test_array_comparisons_of_text_stay_unresolved():
    # Excel compares text by its collation per element; until that is compiled, refuse rather than guess.
    _, c = run({"E1": ('=SUMPRODUCT((D!A1:A2="a")*1)', 0)}, {"A1": "a", "A2": "b"})
    assert "not supported yet" in c.unresolved[("S", 1, 5)]


def test_single_cell_array_formulas_compile_and_multi_cell_ones_do_not():
    from cellproof.reader import load
    from xlsxkit import write_xlsx
    import tempfile
    from pathlib import Path
    with tempfile.TemporaryDirectory() as tmp:
        wb = Path(tmp) / "b.xlsx"
        write_xlsx(wb, {"S": {"A1": 1, "B1": 10, "A2": 2, "B2": 20, "A3": 3, "B3": 30,
                              "F1": ("{=SUM(A1:A3*B1:B3)}", 140),                 # one cell: evaluated as a whole
                              "G1": ("{=A1:A3*B1:B3}", 10, "G1:G3")}})            # a block: needs spill support
        book = load(wb)
        compiled = compile_book(book)
        assert compiled.unresolved[("S", 1, 7)] == "array formula over several cells"
        got = execute(book, compiled)
        assert got[("S", 1, 6)] == N(140)
