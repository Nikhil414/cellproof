from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.model import Cell
from cellproof.verifier import execute, matches


def run(sheets):
    book = book_of(sheets)
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def E(code):
    return ("e", None, code)


def test_operators_produce_excel_error_codes():
    got, c = run({"S": {
        "A1": 0, "A2": -8,
        "B1": ("=1/A1", "#DIV/0!"),
        "B2": ("=A1^-1", "#DIV/0!"),
        "B3": ("=A1^0", "#NUM!"),
        "B4": ("=A2^(1/3)", -1.9999999999999998),  # real Excel: the real cube root (the verifier allows 1e-9)
        "B5": ("=#N/A", "#N/A"),
        "B6": ("=B5+1", "#N/A"),
        "B7": ("=B1+B5", "#DIV/0!"),  # the first error, left to right
        "B8": ('=B5&"x"', "#N/A"),
        "B9": ("=-B1", "#DIV/0!"),
    }})
    assert c.unresolved == {}
    assert [got[("S", r, 2)] for r in range(1, 10)] == [
        E("#DIV/0!"), E("#DIV/0!"), E("#NUM!"), ("n", -2.0, None), E("#N/A"), E("#N/A"), E("#DIV/0!"), E("#N/A"),
        E("#DIV/0!")]


def test_input_error_cells_flow_through():
    got, c = run({"S": {
        "A1": "#N/A", "A2": 5,
        "B1": ("=A1*2", "#N/A"), "B2": ("=IFERROR(A1, A2)", 5), "B3": ("=SUM(A1:A2)", "#N/A"),
    }})
    assert c.unresolved == {}
    assert got[("S", 1, 2)] == E("#N/A")
    assert got[("S", 2, 2)] == ("n", 5.0, None)
    assert got[("S", 3, 2)] == E("#N/A")


def test_sum_returns_the_first_error_row_by_row():
    got, _ = run({"S": {"A1": 1, "B1": ("=1/0", "#DIV/0!"), "A2": "#N/A", "B2": 2,
                        "C1": ("=SUM(A1:B2)", "#DIV/0!")}})
    assert got[("S", 1, 3)] == E("#DIV/0!")  # B1 (row 1) before A2 (row 2)


def test_if_iferror_and_or_round_carry_errors():
    got, c = run({"S": {
        "A1": 0,
        "B1": ("=IF(1/A1>1, 1, 2)", "#DIV/0!"),
        "B2": ("=IF(A1=0, 1, 1/A1)", 1),  # the error branch is not taken
        "B3": ("=IFERROR(1/A1, -1)", -1),
        "B4": ("=AND(TRUE, 1/A1)", "#DIV/0!"),
        "B5": ("=ROUND(1/A1, 2)", "#DIV/0!"),
        "B6": ("=IFERROR(1/A1, 1/A1)", "#DIV/0!"),
    }})
    assert c.unresolved == {}
    assert [got[("S", r, 2)] for r in range(1, 7)] == [
        E("#DIV/0!"), ("n", 1.0, None), ("n", -1.0, None), E("#DIV/0!"), E("#DIV/0!"), E("#DIV/0!")]


def test_vlookup_missing_key_and_found_error():
    rates = {"A1": "West", "B1": 0.05, "A2": "South", "B2": "#DIV/0!"}
    got, c = run({"R": rates, "X": {
        "A1": "North", "B1": ("=VLOOKUP(A1, R!A1:B2, 2, FALSE)", "#N/A"),
        "A2": "South", "B2": ("=VLOOKUP(A2, R!A1:B2, 2, FALSE)", "#DIV/0!"),
        "A3": "West", "B3": ("=VLOOKUP(A3, R!A1:B2, 2, FALSE)", 0.05),
    }})
    assert c.unresolved == {}
    assert got[("X", 1, 2)] == E("#N/A")
    assert got[("X", 2, 2)] == E("#DIV/0!")
    assert got[("X", 3, 2)] == ("n", 0.05, None)


def test_sumifs_over_a_matched_error_cell_returns_that_error():
    got, c = run({"S": {"A1": "x", "B1": "#N/A", "C1": ('=SUMIFS(B1:B1, A1:A1, "x")', "#N/A")}})
    assert c.unresolved == {} and got[("S", 1, 3)] == E("#N/A")


def test_error_codes_must_match_exactly():
    cell = Cell("S", 1, 1, "x", "#N/A", "e")
    assert matches(cell, "e", None, "#N/A")
    assert not matches(cell, "e", None, "#DIV/0!")
    assert not matches(cell, "n", None, None)
    assert not matches(Cell("S", 1, 1, "x", 1.0, "n"), "e", None, "#N/A")


def test_arithmetic_overflow_in_cells_and_regions():
    from cellproof.verifier import verify

    # The synthetic inputs are below Excel's entry limit. The expected
    # overflow errors follow Microsoft's documented limits and #NUM! rule.
    sheet = {}
    for row in range(1, 5):
        sheet[f'A{row}'] = 9e307
        for col, formula in zip('BCDEF', [f'A{row}+A{row}', f'-A{row}-A{row}',
                                         f'A{row}*10', f'A{row}/0.1', f'A{row}^2']):
            sheet[f'{col}{row}'] = ('=' + formula, '#NUM!')
        sheet[f'G{row}'] = (f'=IFERROR(B{row},7)', 7)
        sheet[f'H{row}'] = (f'=A{row}/0', '#DIV/0!')
        sheet[f'I{row}'] = (f'=A{row}/10', 9e306)
    book = book_of({'S': sheet})
    for regions in (False, True):
        compiled = compile_book(book, regions=regions)
        assert not compiled.unresolved
        verdicts = list(verify(book, compiled))
        assert len(verdicts) == 32
        assert all(v.status == 'PASS' for v in verdicts)


def test_arithmetic_overflow_in_arrays():
    from cellproof.verifier import verify

    sheet = {'A1': 9e307, 'A2': 9e307}
    for row, formula in enumerate(['A1:A2+A1:A2', '-A1:A2-A1:A2', 'A1:A2*10',
                                   'A1:A2/0.1', 'A1:A2^2'], 1):
        sheet[f'B{row}'] = (f'=SUM({formula})', '#NUM!')
    book = book_of({'S': sheet})
    compiled = compile_book(book)
    assert not compiled.unresolved
    verdicts = list(verify(book, compiled))
    assert len(verdicts) == 5
    assert all(v.status == 'PASS' for v in verdicts)


def test_long_arithmetic_retains_intermediate_errors_without_expanding_sql():
    from cellproof.compiler import batch_rows

    sheet = {f'A{row}': 1 for row in range(1, 201)}
    total = '+'.join(sheet)
    sheet['B1'] = ('=' + total, 200)
    sheet['B2'] = ('=Z1*10+' + total + '+1/0', '#NUM!')
    sheet['Z1'] = 9e307
    book = book_of({'S': sheet})
    compiled = compile_book(book, regions=False)
    assert not compiled.unresolved
    assert max(len(batch_rows(b, 'inputs')) for b in compiled.batches) < 125_000
    got = execute(book, compiled)
    assert got[('S', 1, 2)] == ('n', 200, None)
    assert got[('S', 2, 2)] == E('#NUM!')


def test_overflowing_criteria_are_unresolved_without_guessing_excel_behavior():
    from cellproof.verifier import verify

    book = book_of({'S': {
        'A1': 9e307, 'A2': 1,
        'B1': ('=COUNTIF(A1:A2,A1*10)', 0),
        'B2': ('=SUMIF(A1:A2,A1*10)', 0),
        'B3': ('=IFERROR(COUNTIF(A1:A2,A1*10),7)', 7),
        'B4': ('=COUNTIF(A1:A2,A2*1)', 1),
    }})
    got = {v.key: v.status for v in verify(book, compile_book(book))}
    assert [got[('S', row, 2)] for row in range(1, 5)] == ['UNRESOLVED'] * 3 + ['PASS']
