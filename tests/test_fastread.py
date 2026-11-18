"""The fast one-pass loader must read exactly what openpyxl's two-pass load reads: same Book, same errors."""
from __future__ import annotations

import os
import zipfile
from pathlib import Path

import pytest

import cellproof.reader as reader
from cellproof import fastread
from xlsxkit import MAIN, PKG, REL, XML_HEAD, SALES, write_xlsx

CORPUS = Path("/tmp/cellbench/corpus")

# Mixed kinds: numbers, strings, booleans, an Excel error, formulas with and without saved values,
# Excel's cached empty string, a number stored as text, and an escaped control character.
MIXED = {
    "A1": 2, "A2": ("=A1*2", 4), "A3": "text", "A4": True, "A5": "#DIV/0!",
    "A6": ("=A3", "text"), "A7": ("=A1", ""), "A8": ("=A9", None), "A9": "15,000",
    "B1": "_x000D_x", "B2": ("=B1", "_x000D_x"),
}

ESCAPES = {
    "A1": "_x005F_x000D_",  # a literal "_x" written as text
    "A2": ("=A1", "_x005F_x000D_"),
}

TABLES = {"Data": [("Sales", "A1:C3", ["Rep", "Region", "Amount"], 0)]}

# An array formula over a block, and a dynamic one (Excel 365's cm mark).
ARRAYS = {"A1": 1, "A2": 2, "B1": ("{=A1:A2*2}", "1", "B1:B2", "dynamic"),
          "B2": ("{=A1:A2*2}", "2", "B1:B2", "dynamic")}


def same(path) -> object:
    """Both readers must return the same Book (or raise the same kind of error)."""
    fast = fastread.load(path)
    slow = reader.load(path)
    assert fast == slow
    return fast


def test_matches_reader_on_the_battery(tmp_path):
    wb = tmp_path / "mixed.xlsx"
    write_xlsx(wb, {"S": MIXED, "Other": {"A1": "hi"}})
    book = same(wb)
    assert book.sheets == ["S", "Other"]
    assert book.max_row == {"S": 9, "Other": 1}


def test_matches_reader_on_names_tables_arrays_escapes(tmp_path):
    wb = tmp_path / "rich.xlsx"
    write_xlsx(wb, {"Data": {**SALES, "E2": ("=Sales[Amount]", 40000)}},
               names={"Rate": "'Data'!$C$2"}, tables=TABLES)
    same(wb)


def test_matches_reader_on_array_and_dynamic_formulas(tmp_path):
    wb = tmp_path / "arrays.xlsx"
    write_xlsx(wb, {"S": ARRAYS})
    same(wb)


def test_matches_reader_on_the_commission_workbook(tmp_path):
    from test_cli import COMMISSION
    wb = tmp_path / "commission.xlsx"
    write_xlsx(wb, {"Sales": SALES, "Commission": COMMISSION})
    same(wb)


# --- raw XML fixtures: the parts openpyxl reads specially ---------------------

WORKBOOK_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
TABLE_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
SHARED_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
APP_CT = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
STYLE_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"

CONTENT_TYPES = {
    "xl/workbook.xml": WORKBOOK_CT,
    "xl/sharedStrings.xml": SHARED_CT,
    "docProps/app.xml": APP_CT,
    "xl/styles.xml": STYLE_CT,
}

# Four identical cell formats, so a cell can carry s="3": openpyxl indexes _cell_styles by that number, and a
# workbook with no styles part at all has exactly one (so any s >= 1 is an IndexError there too).
STYLES = (f'{XML_HEAD}<styleSheet xmlns="{MAIN}">'
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill>'
          '<fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          '<cellXfs count="4">'
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
          '</cellXfs></styleSheet>')


def raw_xlsx(path, parts):
    """Write a zip of raw XML parts; [Content_Types].xml and _rels/.rels are generated."""
    overrides = []
    for name in parts:
        if name.endswith(".rels"):
            continue
        ct = CONTENT_TYPES.get(name)
        if ct is None:
            if name.startswith("xl/worksheets/"):
                ct = WORKSHEET_CT
            elif name.startswith("xl/tables/"):
                ct = TABLE_CT
            else:
                ct = "application/xml"
        overrides.append(f'<Override PartName="/{name}" ContentType="{ct}"/>')
    content = (f'{XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
               '<Default Extension="xml" ContentType="application/xml"/>'
               + "".join(overrides) + "</Types>")
    root = (f'{XML_HEAD}<Relationships xmlns="{PKG}"><Relationship Id="rId1" '
            f'Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", content)
        z.writestr("_rels/.rels", root)
        for name, xml in parts.items():
            z.writestr(name, xml)


def wb_xml(sheet_tags, defined="", props=""):
    return (f'{XML_HEAD}<workbook xmlns="{MAIN}" xmlns:r="{REL}">{props}'
            f"<sheets>{sheet_tags}</sheets>{defined}</workbook>")


def rels_xml(entries):
    body = "".join(f'<Relationship Id="{i}" Type="{REL}/{kind}" Target="{target}"/>'
                   for i, kind, target in entries)
    return f'{XML_HEAD}<Relationships xmlns="{PKG}">{body}</Relationships>'


def sheet_xml(body):
    return f'{XML_HEAD}<worksheet xmlns="{MAIN}" xmlns:r="{REL}">{body}</worksheet>'


def test_matches_reader_on_shared_strings_and_rich_text(tmp_path):
    wb = tmp_path / "strings.xlsx"
    strings = ('<si><t>plain</t></si><si><t/></si>'
               "<si><r><t>base</t></r><r><t>run</t></r></si>"
               '<si><t xml:space="preserve">  padded  </t></si>'
               "<si><t>a_x000D_b</t></si>")
    cells = "".join(f'<c r="A{i}" t="s"><v>{i - 1}</v></c>' for i in range(1, 6))
    raw_xlsx(wb, {
        "xl/workbook.xml": wb_xml('<sheet name="S" sheetId="1" r:id="rId1"/>'),
        "xl/_rels/workbook.xml.rels": rels_xml([("rId1", "worksheet", "worksheets/sheet1.xml"),
                                                ("rId2", "sharedStrings", "sharedStrings.xml")]),
        "xl/worksheets/sheet1.xml": sheet_xml(f'<sheetData><row r="1">{cells}</row></sheetData>'),
        "xl/sharedStrings.xml": f'{XML_HEAD}<sst xmlns="{MAIN}" count="5">{strings}</sst>',
    })
    same(wb)


def test_matches_reader_on_shared_formulas(tmp_path):
    wb = tmp_path / "shared.xlsx"
    raw_xlsx(wb, {
        "xl/workbook.xml": wb_xml('<sheet name="S" sheetId="1" r:id="rId1"/>'),
        "xl/_rels/workbook.xml.rels": rels_xml([("rId1", "worksheet", "worksheets/sheet1.xml")]),
        "xl/worksheets/sheet1.xml": sheet_xml(
            '<sheetData><row r="1"><c r="A1"><v>5</v></c></row>'
            '<row r="2"><c r="A2"><f t="shared" si="0" ref="A2:A4">A1*2</f><v>10</v></c></row>'
            '<row r="3"><c r="A3"><f t="shared" si="0"/><v>10</v></c></row>'
            '<row r="4"><c r="A4"><f t="shared" si="0"/><v>10</v></c></row></sheetData>'),
    })
    book = same(wb)
    assert book.cells[("S", 2, 1)].formula == "A1*2"
    assert book.cells[("S", 3, 1)].formula == "A2*2"
    assert book.cells[("S", 4, 1)].formula == "A3*2"


def test_matches_reader_on_1904_epoch_and_app(tmp_path):
    wb = tmp_path / "old.xlsx"
    raw_xlsx(wb, {
        "xl/workbook.xml": wb_xml('<sheet name="S" sheetId="1" r:id="rId1"/>', props='<workbookPr date1904="1"/>'),
        "xl/_rels/workbook.xml.rels": rels_xml([("rId1", "worksheet", "worksheets/sheet1.xml")]),
        "xl/worksheets/sheet1.xml": sheet_xml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'),
        "docProps/app.xml": f'{XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
                            'extended-properties"><Application>Microsoft Excel</Application></Properties>',
    })
    book = same(wb)
    assert book.date1904 is True
    assert book.app == "Microsoft Excel"


def test_a_styled_empty_cell_counts_for_max_row_only(tmp_path):
    wb = tmp_path / "styled.xlsx"
    raw_xlsx(wb, {
        "xl/workbook.xml": wb_xml('<sheet name="S" sheetId="1" r:id="rId1"/>'),
        "xl/_rels/workbook.xml.rels": rels_xml([("rId1", "worksheet", "worksheets/sheet1.xml")]),
        "xl/worksheets/sheet1.xml": sheet_xml(
            '<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="5"><c r="A5" s="3"/></row></sheetData>'),
        "xl/styles.xml": STYLES,
    })
    book = same(wb)
    assert ("S", 5, 1) not in book.cells
    assert book.max_row["S"] == 5


def test_rows_and_cells_without_refs_use_the_counters(tmp_path):
    wb = tmp_path / "sparse.xlsx"
    raw_xlsx(wb, {
        "xl/workbook.xml": wb_xml('<sheet name="S" sheetId="1" r:id="rId1"/>'),
        "xl/_rels/workbook.xml.rels": rels_xml([("rId1", "worksheet", "worksheets/sheet1.xml")]),
        "xl/worksheets/sheet1.xml": sheet_xml(
            '<sheetData><row><c><v>10</v></c><c><v>20</v></c></row>'
            '<row r="3"><c r="B3"><v>7</v></c></row></sheetData>'),
    })
    book = same(wb)
    assert book.cells[("S", 1, 1)].value == 10.0
    assert book.cells[("S", 1, 2)].value == 20.0
    assert book.cells[("S", 3, 2)].value == 7.0


def test_matches_reader_on_an_iso_date_cell(tmp_path):
    wb = tmp_path / "iso.xlsx"
    raw_xlsx(wb, {
        "xl/workbook.xml": wb_xml('<sheet name="S" sheetId="1" r:id="rId1"/>'),
        "xl/_rels/workbook.xml.rels": rels_xml([("rId1", "worksheet", "worksheets/sheet1.xml")]),
        "xl/worksheets/sheet1.xml": sheet_xml(
            '<sheetData><row r="1"><c r="A1" t="d"><v>2026-01-01T00:00:00</v></c></row></sheetData>'),
    })
    book = same(wb)
    assert book.cells[("S", 1, 1)].kind == "n"


def test_matches_reader_when_the_sheet_xml_is_corrupt(tmp_path):
    wb = tmp_path / "bad.xlsx"
    raw_xlsx(wb, {
        "xl/workbook.xml": wb_xml('<sheet name="S" sheetId="1" r:id="rId1"/>'),
        "xl/_rels/workbook.xml.rels": rels_xml([("rId1", "worksheet", "worksheets/sheet1.xml")]),
        "xl/worksheets/sheet1.xml": '<worksheet><sheetData><row r="1"><c r="A1">',
    })
    for load in (fastread.load, reader.load):
        with pytest.raises(reader.UnreadableWorkbook, match="cannot open bad.xlsx"):
            load(wb)


def test_a_phonetic_shared_string_breaks_both_readers(tmp_path):
    wb = tmp_path / "phonetic.xlsx"
    raw_xlsx(wb, {
        "xl/workbook.xml": wb_xml('<sheet name="S" sheetId="1" r:id="rId1"/>'),
        "xl/_rels/workbook.xml.rels": rels_xml([("rId1", "worksheet", "worksheets/sheet1.xml"),
                                                ("rId2", "sharedStrings", "sharedStrings.xml")]),
        "xl/worksheets/sheet1.xml": sheet_xml('<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>'),
        "xl/sharedStrings.xml": f'{XML_HEAD}<sst xmlns="{MAIN}" count="1">'
                                '<si><rPh><t>ph</t></rPh><t>base</t></si></sst>',
    })
    for load in (fastread.load, reader.load):
        with pytest.raises(reader.UnreadableWorkbook):
            load(wb)


def test_sheet_titles_are_validated_and_uniquified_like_openpyxl(tmp_path):
    wb = tmp_path / "dup.xlsx"
    raw_xlsx(wb, {
        "xl/workbook.xml": wb_xml('<sheet name="S" sheetId="1" r:id="rId1"/><sheet name="S" sheetId="2" r:id="rId2"/>'),
        "xl/_rels/workbook.xml.rels": rels_xml([("rId1", "worksheet", "worksheets/sheet1.xml"),
                                                ("rId2", "worksheet", "worksheets/sheet2.xml")]),
        "xl/worksheets/sheet1.xml": sheet_xml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'),
        "xl/worksheets/sheet2.xml": sheet_xml('<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>'),
    })
    book = same(wb)
    assert book.sheets == ["S", "S1"]
    assert book.cells[("S", 1, 1)].value == 1.0
    assert book.cells[("S1", 1, 1)].value == 2.0


def test_an_invalid_sheet_title_breaks_both_readers(tmp_path):
    wb = tmp_path / "title.xlsx"
    raw_xlsx(wb, {
        "xl/workbook.xml": wb_xml('<sheet name="A[B" sheetId="1" r:id="rId1"/>'),
        "xl/_rels/workbook.xml.rels": rels_xml([("rId1", "worksheet", "worksheets/sheet1.xml")]),
        "xl/worksheets/sheet1.xml": sheet_xml('<sheetData/>'),
    })
    for load in (fastread.load, reader.load):
        with pytest.raises(reader.UnreadableWorkbook):
            load(wb)


def test_a_table_whose_relationship_is_missing_breaks_both_readers(tmp_path):
    wb = tmp_path / "norel.xlsx"
    raw_xlsx(wb, {
        "xl/workbook.xml": wb_xml('<sheet name="S" sheetId="1" r:id="rId1"/>'),
        "xl/_rels/workbook.xml.rels": rels_xml([("rId1", "worksheet", "worksheets/sheet1.xml")]),
        "xl/worksheets/sheet1.xml": sheet_xml(
            '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><tableParts count="1">'
            '<tablePart r:id="rId9"/></tableParts>'),
        "xl/tables/table1.xml": f'{XML_HEAD}<table xmlns="{MAIN}" id="1" name="T" displayName="T" ref="A1:B2">'
                                '<tableColumns count="1"><tableColumn id="1" name="x"/></tableColumns></table>',
    })
    for load in (fastread.load, reader.load):
        with pytest.raises(reader.UnreadableWorkbook):
            load(wb)


def test_duplicate_table_names_break_both_readers(tmp_path):
    wb = tmp_path / "duptable.xlsx"
    write_xlsx(wb, {"S": {"A1": "x", "B1": "y"}},
               tables={"S": [("Sales", "A1:A1", ["x"], 0), ("Sales", "B1:B1", ["y"], 0)]})
    for load in (fastread.load, reader.load):
        with pytest.raises(reader.UnreadableWorkbook):
            load(wb)


# --- the whole-file checks both readers share ---------------------------------

def test_the_same_suffix_and_zip_errors_as_the_reader(tmp_path):
    xls = tmp_path / "old.xls"
    xls.write_bytes(b"x")
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a zip")
    for path in (xls, bad):
        fast_err = slow_err = None
        try:
            fastread.load(path)
        except reader.UnreadableWorkbook as err:
            fast_err = str(err)
        try:
            reader.load(path)
        except reader.UnreadableWorkbook as err:
            slow_err = str(err)
        assert fast_err == slow_err
        assert "save the workbook as .xlsx" in fast_err or "cannot open bad.xlsx" in fast_err


def test_the_zip_bomb_limit_is_live_for_both_readers(tmp_path, monkeypatch):
    wb = tmp_path / "big.xlsx"
    write_xlsx(wb, {"S": {"A1": "x" * 5000}})
    monkeypatch.setattr(reader, "MAX_UNZIPPED", 1000)
    for load in (fastread.load, reader.load):
        with pytest.raises(reader.UnreadableWorkbook, match="GB limit"):
            load(wb)


# --- the downloaded corpus, when present --------------------------------------

def corpus_files():
    if not CORPUS.is_dir():
        return []
    return sorted(p for d in CORPUS.iterdir() if d.is_dir()
                  for p in d.iterdir() if p.suffix.lower() in reader.SUPPORTED)


@pytest.mark.skipif(not corpus_files(), reason="the corpus is not downloaded")
@pytest.mark.parametrize("path", corpus_files(), ids=lambda p: f"{p.parent.name}/{p.name}")
def test_matches_reader_on_the_corpus(path):
    same(path)


# --- the CLI wiring -----------------------------------------------------------

def test_convert_fast_matches_the_default(tmp_path, capsys):
    from cellproof.cli import main
    wb = tmp_path / "commission.xlsx"
    write_xlsx(wb, {"Sales": SALES, "Commission": {
        "A1": "Rep", "B1": "Total", "C1": "Bonus",
        "A2": "Asha", "B2": ("=SUMIFS(Sales!C:C, Sales!A:A, A2)", 120000),
        "C2": ("=IF(B2>100000, B2*5%, B2*2%)", 6000)}})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out"), "--fast"]) == 0
    assert "2 pass, 0 fail, 0 unresolved" in capsys.readouterr().out
    assert (tmp_path / "out" / "REPORT.md").exists()


def test_sql_fast_prints_sql(tmp_path, capsys):
    from cellproof.cli import main
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2, "A2": ("=A1*2", 4)}})
    assert main(["sql", str(wb), "--fast"]) == 0
    assert "SELECT" in capsys.readouterr().out.upper()


def test_watch_fast_verifies(tmp_path):
    from cellproof.cli import main
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2, "A2": ("=A1*2", 4)}})
    assert main(["watch", str(wb), "--runs", "0", "--fast"]) == 0


def test_an_unreadable_file_falls_back_to_openpyxl(tmp_path, capsys, monkeypatch):
    from cellproof.cli import main
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2, "A2": ("=A1*2", 4)}})
    import cellproof.fastread as f

    def boom(path):
        raise reader.UnreadableWorkbook("synthetic test failure")

    monkeypatch.setattr(f, "load", boom)
    assert main(["convert", str(wb), "--out", str(tmp_path / "out"), "--fast"]) == 0
    err = capsys.readouterr().err
    assert "openpyxl" in err and "b.xlsx" in err
    assert (tmp_path / "out" / "REPORT.md").exists()
