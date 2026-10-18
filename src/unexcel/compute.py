"""Answer key for formula cells the file has no saved value for (written by a script, never opened in Excel).

Computed by the `formulas` library: pip install unexcel[compute]. Only cells without a saved value are
filled; Excel's own answers always win. The report labels these answers 'computed'.
"""
from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from openpyxl.utils.cell import coordinate_to_tuple

from unexcel.model import Book, Key, Kind, Value


def _answer(x) -> tuple[Value, Kind]:
    if hasattr(x, "item"):  # numpy scalar -> plain Python value
        x = x.item()
    if isinstance(x, bool):
        return x, "b"
    if isinstance(x, (int, float)):
        return float(x), "n"
    if type(x).__name__ == "XlError":  # before str: formulas' tokens subclass str
        return str(x), "e"
    if type(x) is str:
        return x, "s"
    return None, "blank"  # formulas' empty token (or anything unknown): leave the cell unresolved


def fill_missing(book: Book, path: Path) -> tuple[Book, set[Key], str]:
    """Returns (book with computed values, keys computed, why not - empty when it ran)."""
    missing = {k for k, c in book.cells.items() if c.formula is not None and c.kind == "blank"}
    if not missing:
        return book, set(), ""
    if any("'" in s for s in book.sheets):  # formulas 1.3.4 answers #REF! for every formula on these sheets
        return book, set(), "cannot compute answers when a sheet name contains an apostrophe"
    try:
        import formulas
    except ImportError:
        return book, set(), "or install unexcel[compute] to compute it"
    try:
        solution = formulas.ExcelModel().loads(str(path)).finish().calculate()
    except Exception as err:  # formulas raises many types on workbooks it can't model
        return book, set(), f"the formulas library failed: {err}"
    sheets = {s.upper(): s for s in book.sheets}  # formulas upper-cases sheet names
    cells, computed = dict(book.cells), set()
    for ref, ranges in solution.items():
        where, _, addr = ref.rpartition("!")  # "'[book.xlsx]SHEET'!A1"
        sheet = sheets.get(where.split("]", 1)[-1].rstrip("'"))
        if sheet is None or ":" in addr:
            continue
        try:
            row, col = coordinate_to_tuple(addr.replace("$", ""))
        except ValueError:
            continue
        key = (sheet, row, col)
        if key in missing:
            value, kind = _answer(ranges.value[0][0])
            if kind != "blank":
                cells[key] = replace(cells[key], value=value, kind=kind)
                computed.add(key)
    return replace(book, cells=cells), computed, ""
