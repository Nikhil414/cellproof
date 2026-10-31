"""Array formulas over a block of cells (legacy Ctrl+Shift+Enter and dynamic spills): each cell of the block is
computed as Excel evaluates it, element by element. A range where one value is expected takes this cell's element;
a range one row (or column) high repeats down (or across); beyond the array's size a cell is #N/A."""
from xlsxkit import write_xlsx

from cellproof.compiler import compile_book
from cellproof.reader import load
from cellproof.verifier import verify

INPUTS = {"A1": 1.0, "A2": 2.0, "A3": 3.0, "B1": 10.0, "E1": 1.0, "F1": 2.0, "G1": 3.0, "H1": "x", "H2": "#DIV/0!"}


def check(tmp_path, cells, others=None):
    """cells: {ref: (formula, Excel's value)} or {ref: Excel's value} for the non-anchor cells of a block."""
    path = tmp_path / "s.xlsx"
    write_xlsx(path, {"S": {**INPUTS, **cells}, **(others or {})})
    book = load(path)
    compiled = compile_book(book)
    verdicts = {v.key: v for v in verify(book, compiled)}
    bad = {k: (v.status, v.expected, v.actual, v.reason) for k, v in verdicts.items() if v.status != "PASS"}
    return verdicts, bad


def test_a_block_is_computed_element_by_element(tmp_path):
    verdicts, bad = check(tmp_path, {
        "C1": ("{=A1:A3*2}", 2.0, "C1:C3"), "C2": 4.0, "C3": 6.0,
        "D1": ('{=IF(A1:A3>1,"big","small")}', "small", "D1:D3"), "D2": "big", "D3": "big",
        "I1": ("{=A1:A3+B1}", 11.0, "I1:I3"), "I2": 12.0, "I3": 13.0,  # a single cell repeats
    })
    assert bad == {}
    assert len(verdicts) >= 9  # every cell of every block is verified, not only the anchors


def test_a_row_and_a_column_broadcast_to_a_grid(tmp_path):
    cells = {"J1": ("{=E1:G1*A1:A3}", 1.0, "J1:L3")}
    for i, a in enumerate((1, 2, 3)):
        for j, e in enumerate((1, 2, 3)):
            if (i, j) != (0, 0):
                cells[f"{'JKL'[j]}{i + 1}"] = float(a * e)
    _, bad = check(tmp_path, cells)
    assert bad == {}


def test_beyond_the_array_a_cell_is_na_but_is_functions_see_padding(tmp_path):
    # formulas' test.xlsx (Excel-saved): past the array a cell is #N/A whatever the formula does with errors
    # ({=IFERROR(G9:Q9,R9)} one cell past Q9 is #N/A), but an IS function's own range is padded with #N/A elements
    # instead ({=ISERROR(C8:Q8)} past Q8 is TRUE, {=ISNUMBER(..)} FALSE)
    _, bad = check(tmp_path, {
        "C1": ("{=A1:A3*2}", 2.0, "C1:C4"), "C2": 4.0, "C3": 6.0, "C4": "#N/A",
        "D1": ('{=IFERROR(H1:H2,"none")}', "x", "D1:D3"), "D2": "none", "D3": "#N/A",
        "K5": ("{=ISERROR(A1:A3)}", False, "K5:K8"), "K6": False, "K7": False, "K8": True,
        "L5": ("{=ISBLANK(G1:G3)}", False, "L5:L8"), "L6": True, "L7": True, "L8": False,
        "M5": ("{=EDATE(A1:A3,1)}", "#VALUE!", "M5:M6"), "M6": "#VALUE!",  # Analysis ToolPak: no arrays
        "E5": ("{=SUM(A1:A3)}", 6.0, "E5:F5"), "F5": 6.0,  # one value fills the block
    })
    assert bad == {}


def test_a_single_cell_array_formula_shows_the_first_element(tmp_path):
    # formulas' test.xlsx (Excel-saved): {=MONTH(I6:J6)} in one cell is MONTH(I6)
    _, bad = check(tmp_path, {"J5": ("{=A1:A3*100}", 100.0), "J6": ("{=SUM(A1:A3*B1)}", 60.0)})
    assert bad == {}


def test_references_to_a_spill_and_cells_inside_it(tmp_path):
    _, bad = check(tmp_path, {
        "C1": ("{=A1:A3*2}", 2.0, "C1:C3"), "C2": 4.0, "C3": 6.0,
        "M1": ("=SUM(_xlfn.ANCHORARRAY(C1))", 12.0),       # C1# in Excel
        "N1": ("{=_xlfn.ANCHORARRAY(C1)+1}", 3.0, "N1:N3"), "N2": 5.0, "N3": 7.0,
        "O1": ("=C2+1", 5.0),                             # a cell inside a block waits for it
        "O2": ("=_xlfn.SINGLE(A1:A3)", 2.0),              # @A1:A3 in row 2
        "O3": ("=_xlfn.ANCHORARRAY(A1)", 0.0),            # A1 is not a spill
    })
    assert set(bad) == {("S", 3, 15)} and "not a spilled array" in bad[("S", 3, 15)][3]


def test_index_of_a_whole_row_spills_it(tmp_path):
    # pycel lookup.xlsx (Excel-saved): {=INDEX($B$2:$F$6,$B12,J$10)} with J10 = 0 spills row B12 of the table
    verdicts, bad = check(tmp_path, {"P1": 0.0, "P2": 2.0,
                                     "Q1": ("{=INDEX(A1:B3,P2,P1)}", 2.0, "Q1:R1"), "R1": 0.0,
                                     "Q3": ("{=INDEX(A1:A3,0)*10}", 10.0, "Q3:Q5"), "Q4": 20.0, "Q5": 30.0,
                                     "S1": ("=INDEX(A1:B3,P2,P1)", 0.0)})
    assert set(bad) == {("S", 1, 19)} and "INDEX of a whole row" in bad[("S", 1, 19)][3]


def test_a_block_over_other_sheets(tmp_path):
    # xlcalculator's use_case_01.xlsm (Excel-saved): {=First!A2:C7+Second!A2:C7}
    _, bad = check(tmp_path, {"C1": ("{=T!A1:A2+U!A1:A2}", 11.0, "C1:C2"), "C2": 22.0},
                   {"T": {"A1": 1.0, "A2": 2.0}, "U": {"A1": 10.0, "A2": 20.0}})
    assert bad == {}
