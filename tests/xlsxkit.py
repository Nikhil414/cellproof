"""Tiny .xlsx files WITH saved values, so tests control what 'Excel' answered."""
from __future__ import annotations

import zipfile
from collections import defaultdict
from xml.sax.saxutils import escape

from openpyxl.utils.cell import coordinate_to_tuple

from cellproof.model import Book, Cell
from cellproof.reader import ERRORS

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


def _kind(v):
    if isinstance(v, bool):
        return "b", "1" if v else "0"
    if isinstance(v, (int, float)):
        return "n", repr(v)
    if isinstance(v, str) and v in ERRORS:
        return "e", v
    return "s", str(v)


def _cell_xml(ref, spec):
    if isinstance(spec, tuple):
        formula, saved = spec[0], spec[1]
        dyn = ""
        if formula.startswith("{="):
            spill = spec[2] if len(spec) > 2 else ref  # a CSE formula's ref: one cell, or the block it fills
            f = f'<f t="array" ref="{spill}">{escape(formula[2:-1])}</f>'
            dyn = ' cm="1"' if len(spec) > 3 and spec[3] == "dynamic" else ""  # Excel 365's dynamic-array mark
        else:
            f = f"<f>{escape(formula[1:])}</f>"
        if saved is None:
            return f'<c r="{ref}"{dyn}>{f}</c>'
        kind, text = _kind(saved)
        t = {"n": "", "b": ' t="b"', "e": ' t="e"', "s": ' t="str"'}[kind]
        return f'<c r="{ref}"{t}{dyn}>{f}<v>{escape(text)}</v></c>'
    kind, text = _kind(spec)
    if kind == "s":
        return f'<c r="{ref}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'
    t = {"n": "", "b": ' t="b"', "e": ' t="e"'}[kind]
    return f'<c r="{ref}"{t}><v>{text}</v></c>'


def _sheet_xml(cells, parts=0):
    rows = defaultdict(list)
    for ref, spec in cells.items():
        r, c = coordinate_to_tuple(ref)
        rows[r].append((c, _cell_xml(ref, spec)))
    body = "".join(
        f'<row r="{r}">' + "".join(x for _, x in sorted(cs)) + "</row>" for r, cs in sorted(rows.items())
    )
    tail = ""
    if parts:
        tail = f'<tableParts count="{parts}">' + "".join(
            f'<tablePart r:id="rId{i}"/>' for i in range(1, parts + 1)) + "</tableParts>"
    return (f'{XML_HEAD}<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetData>{body}</sheetData>{tail}'
            "</worksheet>")


def _table_xml(tid, name, ref, columns, totals):
    cols = "".join(f'<tableColumn id="{i}" name="{escape(c, {chr(34): "&quot;"})}"/>'
                   for i, c in enumerate(columns, 1))
    return (f'{XML_HEAD}<table xmlns="{MAIN}" id="{tid}" name="{name}" displayName="{name}" ref="{ref}"'
            + (f' totalsRowCount="{totals}"' if totals else "")
            + f'><tableColumns count="{len(columns)}">{cols}</tableColumns></table>')


def write_xlsx(path, sheets, names=None, tables=None):
    """sheets: {"Sheet": {"A1": 2, "A2": ("=A1*2", 4)}}; names: {"Rate": "'My Sheet'!$A$1"};
    tables: {"Sheet": [("Sales", "A1:C5", ["Rep", "Region", "Amount"], totals_rows)]}."""
    n = len(sheets)
    tables = tables or {}
    table_parts, tid = [], 0  # (sheet number, table number, xml)
    for i, name in enumerate(sheets, 1):
        for spec in tables.get(name, []):
            tid += 1
            table_parts.append((i, tid, _table_xml(tid, *spec)))
    sheet_ct = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, n + 1)
    )
    content_types = (
        f'{XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        f"{sheet_ct}"
        + "".join(f'<Override PartName="/xl/tables/table{t}.xml" '
                  'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'
                  for _, t, _ in table_parts)
        + "</Types>"
    )
    root_rels = (
        f'{XML_HEAD}<Relationships xmlns="{PKG}"><Relationship Id="rId1" '
        f'Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    )
    sheet_tags = "".join(
        f'<sheet name="{escape(name, {chr(34): "&quot;"})}" sheetId="{i}" r:id="rId{i}"/>'
        for i, name in enumerate(sheets, 1)
    )
    defined = "".join(f'<definedName name="{k}">{escape(v)}</definedName>' for k, v in (names or {}).items())
    defined = f"<definedNames>{defined}</definedNames>" if defined else ""
    workbook = f'{XML_HEAD}<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>{sheet_tags}</sheets>{defined}</workbook>'
    wb_rels = (
        f'{XML_HEAD}<Relationships xmlns="{PKG}">'
        + "".join(
            f'<Relationship Id="rId{i}" Type="{REL}/worksheet" Target="worksheets/sheet{i}.xml"/>'
            for i in range(1, n + 1)
        )
        + "</Relationships>"
    )
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        for i, cells in enumerate(sheets.values(), 1):
            mine = [(t, xml) for s, t, xml in table_parts if s == i]
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(cells, len(mine)))
            if mine:
                z.writestr(f"xl/worksheets/_rels/sheet{i}.xml.rels", f'{XML_HEAD}<Relationships xmlns="{PKG}">' + "".join(
                    f'<Relationship Id="rId{k}" Type="{REL}/table" Target="../tables/table{t}.xml"/>'
                    for k, (t, _) in enumerate(mine, 1)) + "</Relationships>")
            for t, xml in mine:
                z.writestr(f"xl/tables/table{t}.xml", xml)


def book_of(sheets, names=None) -> Book:
    """Same spec as write_xlsx, built in memory (skips the reader). names: {"Rate": ("Sheet", "A1")}."""
    cells, max_row = {}, {}
    for sheet, specs in sheets.items():
        max_row[sheet] = max((coordinate_to_tuple(ref)[0] for ref in specs), default=1)
        for ref, spec in specs.items():
            r, c = coordinate_to_tuple(ref)
            formula, saved = (spec[0].removeprefix("="), spec[1]) if isinstance(spec, tuple) else (None, spec)
            kind = "blank" if saved is None else _kind(saved)[0]
            value = float(saved) if kind == "n" else saved
            cells[(sheet, r, c)] = Cell(sheet, r, c, formula, value, kind)
    up = {k.upper(): (s, ref.replace("$", "")) for k, (s, ref) in (names or {}).items()}
    return Book(cells, list(sheets), up, max_row)


# The commission example from the design: header row, mixed-case names, one number stored as text.
SALES = {
    "A1": "Rep", "B1": "Region", "C1": "Amount",
    "A2": "Asha", "B2": "West", "C2": 80000,
    "A3": "asha", "B3": "West", "C3": 40000,
    "A4": "Vikram", "B4": "South", "C4": 60000,
    "A5": "Vikram", "B5": "South", "C5": "15,000",
}
