"""Implicit intersection, INDIRECT of fixed text, MODE: the first features a real user's workbook needed."""
from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute


def run(sheets):
    book = book_of(sheets)
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def N(v):
    return ("n", float(v), None)


def test_a_range_used_as_one_value_takes_the_formulas_own_row_or_column():
    s = {"B1": "x"}
    for r in range(2, 6):
        s[f"A{r}"] = float(r * 10)
        s[f"B{r}"] = "F" if r % 2 else "M"
        s[f"C{r}"] = (f'=IF(B$2:B$5="F",A$2:A$5,0)', 0.0)  # Excel: this row's B and A
    s["A7"] = ("=A$2:A$5", 0.0)  # row 7 is outside rows 2-5: #VALUE!
    s["E1"], s["F1"], s["G1"] = 1.0, 2.0, 3.0
    s["F3"] = ("=E1:G1*2", 4.0)  # one row: this column's cell
    got, c = run({"S": s})
    assert c.unresolved == {}
    assert [got[("S", r, 3)] for r in range(2, 6)] == [N(0), N(30), N(0), N(50)]
    assert got[("S", 7, 1)] == ("e", None, "#VALUE!") and got[("S", 3, 6)] == N(4)


def test_a_range_on_another_sheet_used_as_one_value_stays_unresolved():
    _, c = run({"S": {"A2": ("=T!A1:A5", 0.0)}, "T": {"A2": 1.0}})
    assert ("S", 2, 1) in c.unresolved


def test_indirect_of_fixed_text_is_the_reference_it_names():
    book = book_of({"S": {"A1": 3.0, "A2": 5.0, "A3": 5.0, "B1": "M", "B2": "F", "B3": "F",
                          "C1": ('=AVERAGE(INDIRECT("Scores"))', 0.0), "C2": ('=COUNTIF(INDIRECT("Gender"),"F")', 0.0),
                          "C3": ('=INDIRECT("A2")+1', 0.0), "C4": ('=MEDIAN(INDIRECT("S!A1:A3"))', 0.0),
                          "C5": ('=INDIRECT("A"&1)', 0.0)}})
    book.names.update({"SCORES": ("S", "A1:A3"), "GENDER": ("S", "B1:B3")})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert [got[("S", r, 3)] for r in range(1, 5)] == [N(13 / 3), N(2), N(6), N(5)]
    assert compiled.unresolved[("S", 5, 3)] == "volatile function INDIRECT"  # built from text at run time


def test_indirect_to_formula_cells_waits_for_them():
    got, c = run({"S": {"A1": 2.0, "B1": ("=A1*10", 20.0), "C1": ('=INDIRECT("B1")+1', 21.0)}})
    assert c.unresolved == {} and got[("S", 1, 3)] == N(21)


def test_mode_is_the_most_frequent_number_first_seen_on_ties():
    got, c = run({"S": {"A1": 1.0, "A2": 2.0, "A3": 2.0, "A4": 3.0, "A5": 3.0, "A6": "x", "B1": 7.0, "B2": 8.0,
                        "C1": ("=MODE(A1:A6)", 2.0), "C2": ("=_xlfn.MODE.SNGL(A4:A5,A2:A3)", 3.0),
                        "C3": ("=MODE(B1:B2)", "#N/A")}})
    assert c.unresolved == {}
    assert [got[("S", r, 3)] for r in (1, 2, 3)] == [N(2), N(3), ("e", None, "#N/A")]


def test_a_range_that_if_passes_into_another_function_is_not_intersected():
    # real Excel (a user's workbook): MAX(IF(E3:E41="F",math,0)) intersects the condition (this row is "F") and then
    # gives MAX the whole math range; cellproof cannot pass ranges through IF yet, so it must not answer
    s = {"A2": 10.0, "A3": 99.0, "A4": 5.0, "B2": "F", "B3": "M", "B4": "F",
         "C2": ('=MAX(IF(B2:B4="F",A2:A4,0))', 99.0), "D2": ('=IF(B2:B4="F",A2:A4)+1', 11.0),
         "E2": ('=ABS(A2:A4)', 10.0)}
    got, c = run({"S": s})
    assert ("S", 2, 3) in c.unresolved
    assert got[("S", 2, 4)] == N(11) and got[("S", 2, 5)] == N(10)


def test_a_bare_word_is_a_name_never_a_whole_column():
    # =i with no name i defined (or one cellproof skips) is not column I: Excel writes a column as I:I
    _, c = run({"S": {"I4": 5.0, "R4": ("=i", 11.0), "R5": ("=SUM(I:I)", 5.0)}})
    assert "unknown name" in c.unresolved[("S", 4, 18)] and ("S", 5, 18) not in c.unresolved
