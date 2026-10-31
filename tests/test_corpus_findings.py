"""Regression tests for what real Microsoft Excel saved in public formula-engine test workbooks (pycel's
tests/fixtures/logical.xlsx and excelcompiler.xlsx, formulas' test/test_files/test.xlsx; checked 2026-10-19 that
docProps/app.xml names Microsoft Excel). Each expected value is Excel's own answer. If one fails, fix the compiler,
never the number.
"""
import datetime

import openpyxl
from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.reader import load
from cellproof.verifier import execute

S = {"A1": 2, "A2": "1", "A3": 36526, "A5": -1, "A6": "ciao", "A7": True, "A8": False, "A9": 0, "A10": 366,
     "A11": "1987-08-26", "A12": "1/1/1987 05:00 AM"}  # A4 is empty


def run(formulas):
    book = book_of({"S": S, "F": {f"A{i}": (f, v) for i, (f, v) in enumerate(formulas, 1)}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert compiled.unresolved == {}, compiled.unresolved
    return [got[("F", i, 1)] for i in range(1, len(formulas) + 1)]


def B(v):
    return ("b", 1.0 if v else 0.0, None)


def N(v):
    return ("n", float(v), None)


def E(code):
    return ("e", None, code)


def test_and_or_xor_skip_text_and_empty_cells_in_references():
    assert run([("=AND(S!A1,S!A2)", True), ("=AND(S!A3,S!A4)", True), ("=OR(S!A6,FALSE)", False),
                ("=XOR(S!A6,FALSE)", False), ("=OR(S!A2,S!A4)", "#VALUE!"), ('=AND(TRUE,"0",#REF!)', "#REF!"),
                ('=OR("TRUE",FALSE)', True), ('=AND(TRUE,"0")', True)]) == [
        B(True), B(True), B(False), B(False), E("#VALUE!"), E("#REF!"), B(True), B(True)]


def test_an_omitted_or_empty_result_is_zero():
    assert run([('=IF(S!A5,,"Is False")', 0), ("=IF(1,,S!A6)", 0), ("=IFERROR(S!A4,S!A6)", 0),
                ("=IFNA(S!A4,S!A6)", 0)]) == [N(0), N(0), N(0), N(0)]


def _saved_with_date_style(tmp_path, values):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for i, v in enumerate(values, 1):
        ws.cell(row=i, column=1, value=v).number_format = "yyyy-mm-dd"
    path = tmp_path / "d.xlsx"
    wb.save(path)
    return path


def test_date_cells_keep_excels_serial_number(tmp_path):
    # Excel's 1900 calendar has a 29 February 1900 (serial 60); a date-formatted -1 is still the number -1
    path = _saved_with_date_style(tmp_path, [59, 60, 61, -1, datetime.datetime(2024, 3, 1, 12)])
    book = load(path)
    assert [book.cells[("S", r, 1)].value for r in range(1, 6)] == [59.0, 60.0, 61.0, -1.0, 45352.5]
    assert all(book.cells[("S", r, 1)].kind == "n" for r in range(1, 6))


def test_escaped_control_characters_in_text_are_decoded(tmp_path):
    # the file format writes CHAR(1) as _x0001_ and a literal "_x0001_" as _x005F_x0001_; Excel shows them decoded
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"], ws["A2"], ws["A3"] = "a_x0001_b", "_x005F_x0001_", "plain_x"
    path = tmp_path / "x.xlsx"
    wb.save(path)
    book = load(path)
    assert [book.cells[("S", r, 1)].value for r in (1, 2, 3)] == ["a\x01b", "_x0001_", "plain_x"]


def test_edate_eomonth_mround_value_refuse_true_false():
    assert run([("=EDATE(S!A10,S!A8)", "#VALUE!"), ("=EOMONTH(S!A10,S!A7)", "#VALUE!"), ("=EDATE(S!A8,S!A4)", "#VALUE!"),
                ("=MROUND(S!A7,0)", "#VALUE!"), ("=VALUE(S!A7)", "#VALUE!"), ("=VALUE(S!A10)", 366)]) == [
        E("#VALUE!"), E("#VALUE!"), E("#VALUE!"), E("#VALUE!"), E("#VALUE!"), N(366)]


def test_eomonth_of_serial_zero_is_the_end_of_january_1900():
    assert run([("=EOMONTH(S!A9,S!A4)", 31)]) == [N(31)]


def test_text_that_reads_as_a_date_depends_on_the_locale_and_is_unresolved():
    from cellproof.verifier import verify
    formulas = [("=VALUE(S!A11)", 32015), ("=WEEKDAY(S!A12)", 5), ('=AVERAGE(1,"29/02/1900")', 30.5),
                ("=IFERROR(VALUE(S!A12),0)", 31778.2), ("=ISERROR(VALUE(S!A11))", False), ('=VALUE("1,000")', 1000),
                ('=VALUE("Jan 5")', 5), ('=VALUE("x1")', "#VALUE!"), ('=VALUE("Marketing")', "#VALUE!")]
    book = book_of({"S": S, "F": {f"A{i}": (f, v) for i, (f, v) in enumerate(formulas, 1)}})
    status = {v.key: (v.status, v.reason) for v in verify(book, compile_book(book))}
    for i in range(1, 8):
        assert status[("F", i, 1)][0] == "UNRESOLVED" and "language" in status[("F", i, 1)][1], formulas[i - 1]
    assert status[("F", 8, 1)][0] == status[("F", 9, 1)][0] == "PASS"


def _with_app(tmp_path, app):
    import re
    import zipfile
    src = tmp_path / "src.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.active["A1"] = 200
    wb.save(src)
    out = tmp_path / "app.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item)
            if item.filename == "docProps/app.xml":
                data = re.sub(rb"<Application>[^<]*</Application>", f"<Application>{app}</Application>".encode(), data)
            zout.writestr(item, data)
    return out


def test_the_saving_program_is_read_and_a_non_excel_one_is_reported(tmp_path):
    assert load(_with_app(tmp_path, "Microsoft Excel")).app == "Microsoft Excel"
    assert not compile_book(load(_with_app(tmp_path, "Microsoft Excel"))).findings
    book = load(_with_app(tmp_path, "LibreOffice/6.4.5.2$Linux_X86_64"))
    assert book.app.startswith("LibreOffice")
    assert any("LibreOffice" in f and "not Microsoft Excel" in f for f in compile_book(book).findings)


def test_char_above_127_is_unresolved_when_excel_for_mac_calculated_the_file():
    from cellproof.verifier import verify
    book = book_of({"S": {"A1": 200, "A2": 65, "B1": ("=CHAR(A1)", "\u00c8"), "B2": ("=CHAR(A2)", "A")}})
    book.app = "Microsoft Macintosh Excel"
    status = {v.key: v.status for v in verify(book, compile_book(book))}
    assert status[("S", 1, 2)] == "UNRESOLVED" and status[("S", 2, 2)] == "PASS"


def test_sumproduct_of_different_sizes_returns_an_error_in_its_arrays_first():
    book = book_of({"S": {"A1": ("=1/0", "#DIV/0!"), "B1": 1, "C1": 2, "C2": 3, "A2": 4,
                          "D1": ("=SUMPRODUCT(A1:B1,C1:C3)", "#DIV/0!"), "D2": ("=SUMPRODUCT(A2:B2,C1:C3)", "#VALUE!")}})
    got = execute(book, compile_book(book))
    assert got[("S", 1, 4)] == E("#DIV/0!") and got[("S", 2, 4)] == E("#VALUE!")


def _cse(book, *keys):
    """Mark cells as single-cell array (Ctrl+Shift+Enter) formulas, as the reader does."""
    import dataclasses
    from cellproof.model import a1
    for k in keys:
        book.cells[k] = dataclasses.replace(book.cells[k], array=True, ref=a1(k).split("!")[-1])
    return book


def test_index_with_a_negative_row_or_column_is_value_error():
    # pycel lookup.xlsx (Excel-saved): INDEX($B$2:$F$6, 1, -1) and INDEX({...}, -1, 2) are #VALUE!, not #REF!
    s = {"A1": 1.0, "A2": 2.0, "B1": 3.0, "B2": 4.0}
    got = run([("=INDEX(S!A1:B2,1,-1)", "#VALUE!"), ("=INDEX(S!A1:B2,-1,2)", "#VALUE!"),
               ("=INDEX(S!A1:B2,1,3)", "#REF!"), ("=INDEX(S!A1:B2,-0.5,1)", "#UNCHECKED")])
    assert got == [E("#VALUE!"), E("#VALUE!"), E("#REF!"), E("#UNCHECKED")]  # -0.5 truncates to 0: a whole column


def test_switch_reads_an_empty_cell_as_zero():
    # pycel logical.xlsx (Excel-saved): SWITCH(FALSE, <empty>, "match") and SWITCH(<empty>, "", "match") are #N/A,
    # SWITCH(0, <empty>, "match") and SWITCH(<empty>, 0, "match") are "match" - unlike =, where empty is like FALSE
    got = run([('=SWITCH(S!A8,S!A4,"match")', "#N/A"), ('=SWITCH(S!A4,"","match")', "#N/A"),
               ('=SWITCH(S!A9,S!A4,"match")', "match"), ('=SWITCH(S!A4,0,"match")', "match")])
    assert got == [E("#N/A"), E("#N/A"), ("s", None, "match"), ("s", None, "match")]


def test_a_range_where_one_value_is_expected_in_an_array_formula_is_not_intersected():
    # formulas test.xlsx (Excel-saved): {=MONTH(I6:J6)} in one cell is MONTH of the first element (array
    # evaluation), not implicit intersection's #VALUE!
    book = _cse(book_of({"S": {"A1": 31000.0, "B1": 32000.0, "C5": ("=MONTH(A1:B1)", 11.0),
                               "C6": ("=SUM(A1:B1*2)", 126000.0)}}), ("S", 5, 3), ("S", 6, 3))
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert compiled.unresolved == {}
    assert got[("S", 5, 3)] == N(11) and got[("S", 6, 3)] == N(126000)
