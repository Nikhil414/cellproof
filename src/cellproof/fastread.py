"""Read a workbook in one pass, with the standard library's XML parser.

openpyxl reads the file twice (formulas, then the values Excel saved) and builds a Cell object for every cell.
This module reads each worksheet's XML once and keeps only what cellproof needs. openpyxl's own parsers still do
everything except the cell loop -- the manifest, shared strings, styles, workbook, merges, hyperlinks, comments,
tables and names -- so the Book is exactly the one reader.load builds. Anything unexpected here raises
UnreadableWorkbook and the CLI reads the file with reader.load instead.
"""
from __future__ import annotations

import datetime
import zipfile
from pathlib import Path

from openpyxl.cell.cell import MergedCell
from openpyxl.chartsheet.chartsheet import Chartsheet
from openpyxl.comments.comment_sheet import CommentSheet
from openpyxl.packaging.manifest import Manifest
from openpyxl.packaging.relationship import RelationshipList, get_dependents, get_rels_path
from openpyxl.reader.excel import _find_workbook_part
from openpyxl.reader.workbook import WorkbookParser
from openpyxl.styles.borders import Border
from openpyxl.styles.stylesheet import apply_stylesheet
from openpyxl.utils import range_boundaries
from openpyxl.utils.datetime import CALENDAR_MAC_1904, to_excel
from openpyxl.worksheet._reader import FORMULA_TAG, WorkSheetParser, WorksheetReader
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
from openpyxl.worksheet.table import Table as XlTable
from openpyxl.xml.constants import ARC_CONTENT_TYPES, COMMENTS_NS, SHARED_STRINGS
from openpyxl.xml.functions import fromstring

from cellproof import reader
from cellproof.model import Book, Calc, Cell
from cellproof.reader import UnreadableWorkbook
from cellproof.tables import Table

# Merged-cell bookkeeping asks every cell for its border; a value cell without one shares this.
_NO_BORDER = Border()


class _Plain:
    """What openpyxl's Cell does for us during binding: merging, hyperlinks and comments change cells in place.

    We keep coordinates, the formula-pass value and the saved value; nothing ever reads a cell's style or text
    from these objects, only merges move borders between them."""
    __slots__ = ("row", "column", "value", "data_type", "saved", "border", "protection",
                 "hyperlink", "comment")

    def __init__(self, row: int, column: int, value, data_type: str, saved) -> None:
        self.row, self.column = row, column
        self.value, self.data_type, self.saved = value, data_type, saved
        self.border, self.protection = _NO_BORDER, None
        self.hyperlink = self.comment = None


class _SheetParser(WorkSheetParser):
    """openpyxl's worksheet parser, plus the value Excel saved, taken from the same <c> element.

    A cell with no <f> reads the same either way, so only formula cells are converted a second time."""

    def parse_cell(self, element):
        counter = self.col_counter
        formula = super().parse_cell(element)
        if isinstance(formula["value"], DataTableFormula):
            raise ValueError("data table formulas are not supported")
        if element.find(FORMULA_TAG) is None:
            formula["saved"] = (formula["value"], formula["data_type"])
            return formula
        self.col_counter, self.data_only = counter, True  # the second pass must not move a missing ref's counter
        try:
            saved = super().parse_cell(element)
        finally:
            self.data_only = False
        formula["saved"] = (saved["value"], saved["data_type"])
        return formula


class _Binder(WorksheetReader):
    """WorksheetReader with openpyxl's Cell objects replaced by _Plain ones; the rest of bind_all is reused."""

    def __init__(self, ws, source, strings) -> None:
        self.ws = ws
        self.parser = _SheetParser(source, strings, False, ws.parent.epoch, set(), set(), False)
        self.tables = []

    def bind_cells(self) -> None:
        for _number, row in self.parser.parse():
            for cell in row:
                self.ws.parent._cell_styles[cell["style_id"]]  # openpyxl raises here on a bad style index
                self.ws._cells[(cell["row"], cell["column"])] = _Plain(
                    cell["row"], cell["column"], cell["value"], cell["data_type"], cell["saved"])
        if self.ws._cells:
            self.ws._current_row = self.ws.max_row  # cells, not row dimensions, as openpyxl does


def _saved(value, data_type: str, epoch) -> tuple:
    """The saved value and its kind, exactly as reader._saved reads an openpyxl data_only cell."""
    if value is None:
        if data_type == "str":  # Excel's cached empty string is no value at all
            return "", "s"
        return None, "blank"
    if data_type == "e":
        return str(value), "e"
    if isinstance(value, bool):  # before int: bool is an int subclass
        return value, "b"
    if isinstance(value, (int, float)):
        return float(value), "n"
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):  # ISO dates (t="d")
        return float(to_excel(value, epoch)), "n"
    return reader._text(str(value)), "s"


def _read(path: Path, app: str, calc: Calc) -> Book:
    with zipfile.ZipFile(path) as z:
        valid = z.namelist()
        package = Manifest.from_tree(fromstring(z.read(ARC_CONTENT_TYPES)))
        strings = []
        part = package.find(SHARED_STRINGS)
        if part is not None:
            with z.open(part.PartName[1:]) as source:
                strings = reader._raw_strings(source)
        parser = WorkbookParser(z, _find_workbook_part(package).PartName[1:])
        parser.parse()
        wb = parser.wb
        wb._sheets = []
        apply_stylesheet(z, wb)
        for sheet, rel in parser.find_sheets():
            if rel.target not in valid:  # openpyxl drops a sheet whose part is missing
                continue
            if "chartsheet" in rel.Type:
                cs = Chartsheet.from_tree(fromstring(z.read(rel.target)))
                cs._parent = wb
                cs.title = sheet.name
                wb._add_sheet(cs)
                continue
            rels = RelationshipList()
            rels_path = get_rels_path(rel.target)
            if rels_path in valid:
                rels = get_dependents(z, rels_path)
            ws = wb.create_sheet(sheet.name)
            ws._rels = rels
            with z.open(rel.target) as source:
                binder = _Binder(ws, source, strings)
                binder.bind_all()
            for r in rels.find(COMMENTS_NS):
                comments = CommentSheet.from_tree(fromstring(z.read(r.target)))
                for ref, comment in comments.comments:
                    try:
                        ws[ref].comment = comment
                    except AttributeError:
                        if isinstance(ws[ref], MergedCell):  # a comment on a merged-away cell is dropped
                            continue
            for target in binder.tables:
                ws.add_table(XlTable.from_tree(fromstring(z.read(target))))
            ws.sheet_state = sheet.state

        parser.assign_names()
        cells, max_row = {}, {}
        for ws in wb.worksheets:
            max_row[ws.title] = ws.max_row
            for (r, c), cell in sorted(ws._cells.items()):
                if not isinstance(cell, _Plain) or cell.value is None:  # a merge replaced it, or it was empty
                    continue
                value, kind = _saved(*cell.saved, wb.epoch)
                formula, array, ref = None, False, None
                if isinstance(cell.value, ArrayFormula):
                    formula, array, ref = (cell.value.text or "").removeprefix("="), True, cell.value.ref
                elif cell.data_type == "f":
                    formula = cell.value.removeprefix("=")
                cells[(ws.title, r, c)] = Cell(ws.title, r, c, formula, value, kind, array, ref)

    names = {}
    for name, dn in wb.defined_names.items():
        try:
            dests = list(dn.destinations)
        except Exception:  # noqa: BLE001 - openpyxl crashes on a name that is a formula of names ("h+i")
            dests = []
        if len(dests) == 1:
            sheet, ref = dests[0]
            names[name.upper()] = (sheet, ref.replace("$", ""))
    tables = {}
    for ws in wb.worksheets:
        for t in ws.tables.values():
            c1, r1, c2, r2 = range_boundaries(t.ref)
            header = 1 if t.headerRowCount is None else int(t.headerRowCount)
            tables[(t.displayName or t.name).upper()] = Table(
                t.displayName or t.name, ws.title, r1, c1, r2, c2, header, int(t.totalsRowCount or 0),
                tuple(col.name for col in t.tableColumns))
    return Book(cells, [ws.title for ws in wb.worksheets], names, max_row, wb.epoch == CALENDAR_MAC_1904, app,
                tables=tables, calc=calc)


def load(path: str | Path) -> Book:
    path = Path(path)
    if path.suffix.lower() not in reader.SUPPORTED:
        raise UnreadableWorkbook(f"{path.suffix or 'these'} files are not supported in v1 - save the workbook as .xlsx")
    try:
        with zipfile.ZipFile(path) as z:
            size = sum(i.file_size for i in z.infolist())
            app = reader._app(z)
            calc = reader._calc(z)
    except (OSError, zipfile.BadZipFile) as err:
        raise UnreadableWorkbook(f"cannot open {path.name}: {err}") from err
    if size > reader.MAX_UNZIPPED:
        raise UnreadableWorkbook(f"cannot open {path.name}: it unzips to {size / 1024**3:.1f} GB, more than the "
                                 f"{reader.MAX_UNZIPPED // 1024**3} GB limit")
    try:
        book = _read(path, app, calc)
    except UnreadableWorkbook:
        raise
    except Exception as err:  # openpyxl raises many types for corrupt or encrypted files
        raise UnreadableWorkbook(f"cannot open {path.name}: {err}") from err

    reader._mark_dynamic(path, book)
    reader.spread_blocks(book)
    reader.rewrite_anchors(book)
    reader.rewrite_tables(book)
    return book
