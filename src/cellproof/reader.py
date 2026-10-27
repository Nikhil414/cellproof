"""Load a workbook twice: once for formulas, once for the values Excel saved."""
from __future__ import annotations

import datetime
import re
import zipfile
from contextlib import contextmanager
from pathlib import Path

import openpyxl
from openpyxl.utils.datetime import CALENDAR_MAC_1904, to_excel
from openpyxl.worksheet.formula import ArrayFormula

from cellproof.model import Book, Cell, Kind, Value

ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
SUPPORTED = {".xlsx", ".xlsm"}
# An .xlsx is a zip of XML. A small file that unzips to gigabytes (a "zip bomb") would exhaust memory in openpyxl.
# ponytail: one total-size limit; raise it if real workbooks ever need more
MAX_UNZIPPED = 2 * 1024**3


class UnreadableWorkbook(Exception):
    pass


ESCAPED = re.compile(r"_x([0-9A-Fa-f]{4})_")


def _text(s: str) -> str:
    """The file format writes a control character as _xHHHH_ (CHAR(13) is _x000D_) and a literal "_x" as _x005F_x:
    one left-to-right pass decodes both, as Excel does."""
    return ESCAPED.sub(lambda m: chr(int(m.group(1), 16)), s) if "_x" in s else s


def _raw_strings(xml_source):
    """openpyxl's read_string_table without its partial unescaping (it drops "x005F_" only, which makes an escaped
    control character and the same text typed literally look alike)."""
    from openpyxl.cell.text import Text
    from openpyxl.xml.constants import SHEET_MAIN_NS
    from openpyxl.xml.functions import iterparse
    strings, tag = [], "{%s}si" % SHEET_MAIN_NS
    for _, node in iterparse(xml_source):
        if node.tag == tag:
            strings.append(Text.from_tree(node).content)
            node.clear()
    return strings


@contextmanager
def _as_saved():
    """Load cells as the file stores them: a date-formatted number stays its serial number (openpyxl's dates lose
    Excel's 29 February 1900, serial 60, and turn a date-formatted -1 into #VALUE!), and text stays escaped for
    _text. Patches openpyxl only while our own load runs."""
    import openpyxl.reader.excel as excel
    import openpyxl.worksheet._reader as sheet_reader
    saved = excel.read_string_table, sheet_reader.from_excel
    excel.read_string_table = _raw_strings
    sheet_reader.from_excel = lambda value, *args, **kwargs: value
    try:
        yield
    finally:
        excel.read_string_table, sheet_reader.from_excel = saved


def _saved(cell, epoch) -> tuple[Value, Kind]:
    v = cell.value
    if v is None:
        if cell.data_type == "str":  # Excel's cached empty string is None in openpyxl.
            return "", "s"
        return None, "blank"
    if cell.data_type == "e":
        return str(v), "e"
    if isinstance(v, bool):  # before int: bool is an int subclass
        return v, "b"
    if isinstance(v, (int, float)):
        return float(v), "n"
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):  # ISO dates (t="d") from other writers
        return float(to_excel(v, epoch)), "n"  # handles Excel's 1900 leap-year bug and the 1904 system
    return _text(str(v)), "s"


def _app(z: zipfile.ZipFile) -> str:
    """The program that last saved the file, as docProps/app.xml names it (printable ASCII, at most 80 characters)."""
    try:
        info = z.getinfo("docProps/app.xml")
    except KeyError:
        return ""
    if info.file_size > 1024**2:
        return ""
    m = re.search(rb"<(?:\w+:)?Application>([^<]*)<", z.read(info))
    return re.sub(r"[^ -~]", "", m.group(1).decode("utf-8", "replace"))[:80].strip() if m else ""


def load(path: str | Path) -> Book:
    path = Path(path)
    if path.suffix.lower() not in SUPPORTED:
        raise UnreadableWorkbook(f"{path.suffix or 'these'} files are not supported in v1 - save the workbook as .xlsx")
    try:
        with zipfile.ZipFile(path) as z:
            size = sum(i.file_size for i in z.infolist())
            app = _app(z)
    except (OSError, zipfile.BadZipFile) as err:
        raise UnreadableWorkbook(f"cannot open {path.name}: {err}") from err
    if size > MAX_UNZIPPED:
        raise UnreadableWorkbook(f"cannot open {path.name}: it unzips to {size / 1024**3:.1f} GB, more than the "
                                 f"{MAX_UNZIPPED // 1024**3} GB limit")
    try:
        with _as_saved():
            fwb = openpyxl.load_workbook(path)
            vwb = openpyxl.load_workbook(path, data_only=True)
    except Exception as err:  # openpyxl raises many types for corrupt or encrypted files
        raise UnreadableWorkbook(f"cannot open {path.name}: {err}") from err

    names = {}
    for name, dn in fwb.defined_names.items():
        try:
            dests = list(dn.destinations)
        except Exception:  # noqa: BLE001 - openpyxl crashes on a name that is a formula of names ("h+i")
            dests = []  # not a plain range: formulas using it are UNRESOLVED ("unknown name")
        if len(dests) == 1:
            sheet, ref = dests[0]
            names[name.upper()] = (sheet, ref.replace("$", ""))

    cells, max_row = {}, {}
    for ws in fwb.worksheets:
        vws = vwb[ws.title]
        max_row[ws.title] = ws.max_row
        # ponytail: openpyxl's private _cells holds only the cells the file has; iter_rows() would build
        # every row x column up to max_row (one styled cell at G1048576 cost 6 s; at Z, 3.8 GB)
        for (r, col), c in sorted(ws._cells.items()):
            if c.value is None:
                continue
            value, kind = _saved(vws.cell(row=r, column=col), fwb.epoch)
            formula, array, ref = None, False, None
            if isinstance(c.value, ArrayFormula):
                formula, array, ref = (c.value.text or "").removeprefix("="), True, c.value.ref
            elif c.data_type == "f":
                formula = c.value.removeprefix("=")
            cells[(ws.title, c.row, c.column)] = Cell(ws.title, c.row, c.column, formula, value, kind, array, ref)
    return Book(cells, [ws.title for ws in fwb.worksheets], names, max_row, fwb.epoch == CALENDAR_MAC_1904, app)
