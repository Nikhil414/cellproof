# unexcel — Plan 1: Deterministic Core (cell mode) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `unexcel convert book.xlsx` reads a workbook, compiles every formula cell to SQL, proves each result against Excel's saved value on DuckDB, and writes a dbt project + `REPORT.md`.

**Architecture:** Plain-code pipeline `reader → graph → compiler(+functions) → verifier → emitter → report → cli`. This plan builds **cell mode** (design §4 fallback "C"): each formula cell becomes one materialized CTE over a long-format `input_cells` seed. It's always exact and is the foundation the later plans build on. No AI, no network.

**Setup:** every command assumes the venv is active (`source .venv/bin/activate` on macOS/Linux, `.venv\Scripts\activate` on Windows).

**Tech Stack:** Python ≥3.10, openpyxl 3.1, DuckDB ≥1.1, pytest, dbt-duckdb (optional), Jinja2 (tests only).

**Spec:** `UNEXCEL/DESIGN.md` (accepted 2026-10-11)

### Roadmap (this is Plan 1 of 4)

| Plan | Scope | Weeks |
|---|---|---|
| **1 (this)** | Reader, parser, dependency graph, cell-mode compiler with 8 functions + operators, verifier, dbt emitter, report, CLI | 1–2 |
| 2 | Table mode (readable `SUM…GROUP BY` / join models) + region detection; ~80 functions; error *kinds*; computed answer key via `formulas` when the file has no saved values; 200k-formula perf check | 2–3 |
| 3 | Captain loop (≤3 rounds, no-progress stop), Decider adapter (Laya local default, Jev optional), Translator adapter (LLM, structure only), Hermes skill | 3 |
| 4 | Demo workbook, README + VHS gif, CI matrix (Win/macOS/Linux × py3.10–3.13), PyPI under the final name | 4 |

**Deliberate deviations from the spec in Plan 1** (each is picked up by a later plan):
- Cell mode only, with no table detection (Plan 2).
- Errors match only as "any error = any error" (SQL `NULL`); `#N/A` vs `#DIV/0!` is not yet distinguished (Plan 2).
- A formula with no saved value is `UNRESOLVED` and isn't computed (Plan 2).

## Global Constraints

- Python `>=3.10`. Runtime dependencies: **only** `openpyxl>=3.1`, `duckdb>=1.1`. Optional extras: `dbt` → `dbt-duckdb>=1.8`; `dev` → `pytest>=8`, `jinja2>=3.1`.
- Supported inputs: `.xlsx`, `.xlsm`. Anything else → `UnreadableWorkbook`, exit code 2.
- Match rule: numbers `math.isclose(rel_tol=1e-9, abs_tol=1e-12)`; text exact (case-sensitive); booleans exact; Excel error ⇔ SQL `NULL`.
- Never guess: any cell that can't be translated → `UNRESOLVED` with a human-readable reason. Never silently skipped.
- Cell values never leave the process (no network calls in Plan 1).
- Exit codes: 0 = no FAIL, 1 = at least one FAIL (or dbt build failed), 2 = unreadable workbook.
- Console output is ASCII-only, and stdout/stderr use `errors="replace"`.
- In SQL every numeric literal is `CAST(x AS DOUBLE)`, because DuckDB types bare literals as DECIMAL. Every division is `a / NULLIF(b, 0)`, because DuckDB returns `inf` for `x/0`, not an error. (Both verified 2026-10-11 on DuckDB 1.5.5.)
- Input table schema (seed and verifier): `sheet VARCHAR, row_num INTEGER, col_num INTEGER, kind VARCHAR, num DOUBLE, txt VARCHAR`. Model output has the same columns.

## Review Focus

1. **Windows console that can't print Unicode** (cp1252 + a file named `Ümsatz✓.xlsx`) → the CLI must not crash with `UnicodeEncodeError`. Test in Task 11.
2. **Workbook text containing `{{ }}` / `{% %}`** → the dbt model must still render under Jinja to exactly the verified SQL. Test in Task 9.
3. **Sheet names with spaces and apostrophes** (`'Bob''s Sheet'!A1`) → resolve and escape correctly in SQL. Test in Task 5.
4. **Whole-column references with a text header row** (`SUMIFS(Sales!C:C, …)`, `VLOOKUP(x, Rates!A:B, 2, FALSE)`) → the header must not break typing or totals. Tests in Task 7.
5. **Workbook written by a script (no saved values)** → those formula cells are `UNRESOLVED` with "no saved value…", not FAIL and not a crash. Test in Task 5.

---

## File structure

```
UNEXCEL/
  pyproject.toml          packaging, extras, pytest config
  README.md               one-paragraph stub (full README is Plan 4)
  .gitignore
  DESIGN.md / PLAN.md
  src/unexcel/
    __init__.py           version
    __main__.py           python -m unexcel
    model.py              Cell, Book, Verdict, Key, Unsupported, a1()
    reader.py             load(path) -> Book
    parser.py             parse(formula) -> AST, refs(node)
    refs.py               Area, resolve(text, sheet, book)
    graph.py              formula_order(book)
    expr.py               Expr, INPUT, coercions
    functions.py          FUNCS registry: SUM IF ROUND AND OR IFERROR SUMIFS VLOOKUP
    compiler.py           Compiler, compile_book, render_model, input_rows
    verifier.py           execute, matches, verify
    emitter.py            write_project, dbt_sql, project_name, run_dbt
    report.py             write_report
    cli.py                main(argv)
  tests/
    xlsxkit.py            write_xlsx (real .xlsx WITH saved values), book_of, SALES
    test_reader.py test_parser.py test_refs.py test_graph.py test_compiler.py
    test_functions.py test_lookup.py test_verifier.py test_emitter.py test_report.py
    test_cli.py test_excel_saved.py
  examples/
    make_commission.py    builds the workbook you save in real Excel
```

---

### Task 1: Project skeleton, data model, test workbook kit, reader

**Files:**
- Create: `pyproject.toml`, `README.md`, `.gitignore`, `src/unexcel/__init__.py`, `src/unexcel/model.py`, `src/unexcel/reader.py`, `tests/xlsxkit.py`
- Test: `tests/test_reader.py`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `model.Key = tuple[str, int, int]` (sheet, row, col)
  - `model.Cell(sheet, row, col, formula: str|None, value: float|str|bool|None, kind: "n"|"s"|"b"|"e"|"blank", array: bool=False)` with `.key`
  - `model.Book(cells: dict[Key, Cell], sheets: list[str], names: dict[str, tuple[str, str]], max_row: dict[str, int])`
  - `model.Verdict(key, status: "PASS"|"FAIL"|"UNRESOLVED", expected=None, actual=None, reason="")`
  - `model.Unsupported(Exception)`, `model.a1(key) -> "Sheet!B2"`
  - `reader.load(path) -> Book`, `reader.UnreadableWorkbook`, `reader.ERRORS`
  - `xlsxkit.write_xlsx(path, sheets, names=None)`, `xlsxkit.book_of(sheets, names=None) -> Book`

Cell spec used by `write_xlsx` and `book_of`: a plain value = input cell (`2`, `"hi"`, `True`); a tuple `("=FORMULA", saved)` = formula cell with Excel's saved value (`saved=None` means no saved value; a string in `ERRORS` means an Excel error). `write_xlsx` also accepts `("{=FORMULA}", saved)` for an array formula.

- [ ] **Step 1: Scaffold the repo**

```bash
# repo already exists: git clone https://github.com/Nikhil414/unexcel.git && cd unexcel
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
python -m pip install --upgrade pip
```

`pyproject.toml`:
```toml
[build-system]
requires = ["hatchling>=1.24"]
build-backend = "hatchling.build"

[project]
name = "unexcel"
version = "0.1.0.dev0"
description = "Turn an Excel workbook into a dbt project and prove every number matches."
readme = "README.md"
requires-python = ">=3.10"
dependencies = ["openpyxl>=3.1", "duckdb>=1.1"]

[project.optional-dependencies]
dev = ["pytest>=8", "jinja2>=3.1"]
dbt = ["dbt-duckdb>=1.8"]

[project.scripts]
unexcel = "unexcel.cli:main"

[tool.hatch.build.targets.wheel]
packages = ["src/unexcel"]

[tool.pytest.ini_options]
pythonpath = ["src", "tests"]
testpaths = ["tests"]
```

`README.md`:
```markdown
# unexcel (working name)

Turn an Excel workbook into a dbt project and prove every number matches. Work in progress, see DESIGN.md and PLAN.md.
```

`.gitignore`:
```
.venv/
__pycache__/
.pytest_cache/
*_dbt/
*.duckdb
target/
logs/
examples/commission_unsaved.xlsx
```

`src/unexcel/__init__.py`:
```python
__version__ = "0.1.0.dev0"
```

Install: `python -m pip install -e ".[dev]"`

- [ ] **Step 2: Write `src/unexcel/model.py`**

```python
"""Data types shared by every stage."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Union

from openpyxl.utils.cell import get_column_letter

Key = tuple[str, int, int]  # (sheet, row, col)
Kind = Literal["n", "s", "b", "e", "blank"]
Value = Union[float, str, bool, None]


class Unsupported(Exception):
    """A cell cannot be translated. The message becomes the UNRESOLVED reason shown to the user."""


def a1(key: Key) -> str:
    sheet, row, col = key
    return f"{sheet}!{get_column_letter(col)}{row}"


@dataclass(frozen=True)
class Cell:
    sheet: str
    row: int
    col: int
    formula: str | None  # formula text without the leading "=", None for input cells
    value: Value  # Excel-saved value; numbers as float, dates as Excel serial numbers
    kind: Kind  # kind of the saved value
    array: bool = False  # array (Ctrl+Shift+Enter) formula

    @property
    def key(self) -> Key:
        return (self.sheet, self.row, self.col)


@dataclass
class Book:
    cells: dict[Key, Cell]
    sheets: list[str]
    names: dict[str, tuple[str, str]] = field(default_factory=dict)  # NAME (upper) -> (sheet, ref)
    max_row: dict[str, int] = field(default_factory=dict)


@dataclass(frozen=True)
class Verdict:
    key: Key
    status: Literal["PASS", "FAIL", "UNRESOLVED"]
    expected: Value = None
    actual: Value = None
    reason: str = ""
```

- [ ] **Step 3: Write `tests/xlsxkit.py`** (openpyxl can't write saved values, so this writes the XML directly. The format was verified with openpyxl 3.1.5 on 2026-10-11.)

```python
"""Tiny .xlsx files WITH saved values, so tests control what 'Excel' answered."""
from __future__ import annotations

import zipfile
from collections import defaultdict
from xml.sax.saxutils import escape

from openpyxl.utils.cell import coordinate_to_tuple

from unexcel.model import Book, Cell
from unexcel.reader import ERRORS

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
        formula, saved = spec
        if formula.startswith("{="):
            f = f'<f t="array" ref="{ref}">{escape(formula[2:-1])}</f>'
        else:
            f = f"<f>{escape(formula[1:])}</f>"
        if saved is None:
            return f'<c r="{ref}">{f}</c>'
        kind, text = _kind(saved)
        t = {"n": "", "b": ' t="b"', "e": ' t="e"', "s": ' t="str"'}[kind]
        return f'<c r="{ref}"{t}>{f}<v>{escape(text)}</v></c>'
    kind, text = _kind(spec)
    if kind == "s":
        return f'<c r="{ref}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'
    t = {"n": "", "b": ' t="b"', "e": ' t="e"'}[kind]
    return f'<c r="{ref}"{t}><v>{text}</v></c>'


def _sheet_xml(cells):
    rows = defaultdict(list)
    for ref, spec in cells.items():
        r, c = coordinate_to_tuple(ref)
        rows[r].append((c, _cell_xml(ref, spec)))
    body = "".join(
        f'<row r="{r}">' + "".join(x for _, x in sorted(cs)) + "</row>" for r, cs in sorted(rows.items())
    )
    return f'{XML_HEAD}<worksheet xmlns="{MAIN}"><sheetData>{body}</sheetData></worksheet>'


def write_xlsx(path, sheets, names=None):
    """sheets: {"Sheet": {"A1": 2, "A2": ("=A1*2", 4)}}; names: {"Rate": "'My Sheet'!$A$1"}."""
    n = len(sheets)
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
        f"{sheet_ct}</Types>"
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
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(cells))


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
```

- [ ] **Step 4: Write the failing tests `tests/test_reader.py`**

```python
import datetime

import openpyxl
import pytest
from xlsxkit import write_xlsx

from unexcel.model import Cell
from unexcel.reader import UnreadableWorkbook, load


def test_reads_inputs_formulas_and_saved_values(tmp_path):
    p = tmp_path / "b.xlsx"
    write_xlsx(
        p,
        {"My Sheet": {
            "A1": 2, "B1": "hi", "C1": True,
            "A2": ("=A1*Rate", 4), "B2": ('=B1&"x"', "hix"), "C2": ("=1/0", "#DIV/0!"), "D2": ("=A1", None),
        }},
        names={"Rate": "'My Sheet'!$A$1"},
    )
    book = load(p)
    c = book.cells
    assert c[("My Sheet", 1, 1)] == Cell("My Sheet", 1, 1, None, 2.0, "n")
    assert (c[("My Sheet", 1, 2)].value, c[("My Sheet", 1, 2)].kind) == ("hi", "s")
    assert (c[("My Sheet", 1, 3)].value, c[("My Sheet", 1, 3)].kind) == (True, "b")
    assert c[("My Sheet", 2, 1)] == Cell("My Sheet", 2, 1, "A1*Rate", 4.0, "n")
    assert c[("My Sheet", 2, 2)] == Cell("My Sheet", 2, 2, 'B1&"x"', "hix", "s")
    assert c[("My Sheet", 2, 3)] == Cell("My Sheet", 2, 3, "1/0", "#DIV/0!", "e")
    assert c[("My Sheet", 2, 4)] == Cell("My Sheet", 2, 4, "A1", None, "blank")
    assert book.names == {"RATE": ("My Sheet", "A1")}
    assert book.sheets == ["My Sheet"] and book.max_row == {"My Sheet": 2}


def test_array_formula_is_flagged(tmp_path):
    p = tmp_path / "a.xlsx"
    write_xlsx(p, {"S": {"A1": 1, "B1": ("{=SUM(A1:A1*2)}", 2)}})
    cell = load(p).cells[("S", 1, 2)]
    assert cell.array and cell.formula == "SUM(A1:A1*2)"


def test_dates_become_excel_serials(tmp_path):
    p = tmp_path / "d.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.active["A1"] = datetime.date(2026, 1, 1)
    wb.save(p)
    assert load(p).cells[("S", 1, 1)] == Cell("S", 1, 1, None, 46023.0, "n")


def test_rejects_unsupported_format(tmp_path):
    p = tmp_path / "old.xls"
    p.write_bytes(b"x")
    with pytest.raises(UnreadableWorkbook, match="save the workbook as .xlsx"):
        load(p)


def test_rejects_corrupt_file(tmp_path):
    p = tmp_path / "bad.xlsx"
    p.write_bytes(b"not a zip")
    with pytest.raises(UnreadableWorkbook, match="cannot open bad.xlsx"):
        load(p)
```

- [ ] **Step 5: Run to verify it fails**

Run: `python -m pytest tests/test_reader.py -v`
Expected: collection error `ModuleNotFoundError: No module named 'unexcel.reader'` (xlsxkit imports it).

- [ ] **Step 6: Write `src/unexcel/reader.py`**

```python
"""Load a workbook twice: once for formulas, once for the values Excel saved."""
from __future__ import annotations

import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.formula import ArrayFormula

from unexcel.model import Book, Cell, Kind, Value

ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
SUPPORTED = {".xlsx", ".xlsm"}


class UnreadableWorkbook(Exception):
    pass


def _saved(cell, epoch) -> tuple[Value, Kind]:
    v = cell.value
    if v is None:
        return None, "blank"
    if cell.data_type == "e":
        return str(v), "e"
    if isinstance(v, bool):  # before int: bool is an int subclass
        return v, "b"
    if isinstance(v, (int, float)):
        return float(v), "n"
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        return float(to_excel(v, epoch)), "n"  # handles Excel's 1900 leap-year bug and the 1904 system
    return str(v), "s"


def load(path: str | Path) -> Book:
    path = Path(path)
    if path.suffix.lower() not in SUPPORTED:
        raise UnreadableWorkbook(f"{path.suffix or 'these'} files are not supported in v1 - save the workbook as .xlsx")
    try:
        fwb = openpyxl.load_workbook(path)
        vwb = openpyxl.load_workbook(path, data_only=True)
    except Exception as err:  # openpyxl raises many types for corrupt or encrypted files
        raise UnreadableWorkbook(f"cannot open {path.name}: {err}") from err

    names = {}
    for name, dn in fwb.defined_names.items():
        dests = list(dn.destinations)
        if len(dests) == 1:
            sheet, ref = dests[0]
            names[name.upper()] = (sheet, ref.replace("$", ""))

    cells, max_row = {}, {}
    for ws in fwb.worksheets:
        vws = vwb[ws.title]
        max_row[ws.title] = ws.max_row
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                value, kind = _saved(vws.cell(row=c.row, column=c.column), fwb.epoch)
                formula, array = None, False
                if isinstance(c.value, ArrayFormula):
                    formula, array = (c.value.text or "").removeprefix("="), True
                elif c.data_type == "f":
                    formula = c.value.removeprefix("=")
                cells[(ws.title, c.row, c.column)] = Cell(ws.title, c.row, c.column, formula, value, kind, array)
    return Book(cells, [ws.title for ws in fwb.worksheets], names, max_row)
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `python -m pytest tests/test_reader.py -v`
Expected: 5 passed.

- [ ] **Step 8: Commit**

```bash
git add .
git commit -m "feat: project skeleton, data model and workbook reader" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Formula parser

**Files:**
- Create: `src/unexcel/parser.py`
- Test: `tests/test_parser.py`

**Interfaces:**
- Consumes: `model.Unsupported`.
- Produces: AST dataclasses `Num(value)`, `Str(value)`, `Bool(value)`, `ErrLit(value)`, `RefNode(text)`, `Unary(op, operand)`, `Postfix(op, operand)`, `Binary(op, left, right)`, `Call(name, args: tuple)` (an arg is `None` when empty); `parse(formula: str) -> Node` (formula without `=`); `refs(node) -> Iterator[str]`.

- [ ] **Step 1: Write the failing tests `tests/test_parser.py`**

```python
import pytest

from unexcel.model import Unsupported
from unexcel.parser import Binary, Bool, Call, Num, Postfix, RefNode, Str, Unary, parse, refs


def test_precedence_matches_excel():
    assert parse("-2^2") == Binary("^", Unary("-", Num(2)), Num(2))  # Excel: 4
    assert parse("1+2*3") == Binary("+", Num(1), Binary("*", Num(2), Num(3)))
    assert parse("2^3^2") == Binary("^", Binary("^", Num(2), Num(3)), Num(2))  # left-assoc: 64
    assert parse('1+2&"a"="3a"') == Binary("=", Binary("&", Binary("+", Num(1), Num(2)), Str("a")), Str("3a"))


def test_percent_text_bool_and_refs():
    assert parse("5%") == Postfix("%", Num(5))
    assert parse('"say ""hi"""') == Str('say "hi"')
    assert parse("TRUE") == Bool(True)
    assert parse("'My Sheet'!$A$1") == RefNode("'My Sheet'!$A$1")


def test_functions_and_empty_args():
    assert parse("SUM(A1:A3, 2)") == Call("SUM", (RefNode("A1:A3"), Num(2)))
    assert parse("_xlfn.XLOOKUP(A1,B:B,C:C)").name == "XLOOKUP"
    assert parse("IF(A1,,1)") == Call("IF", (RefNode("A1"), None, Num(1)))
    assert parse("NOW()") == Call("NOW", ())


@pytest.mark.parametrize("bad", ["SUM(1,", "{1,2}", "A1:B2 B1:C3", "(1+2"])
def test_unparseable_raises_unsupported(bad):
    with pytest.raises(Unsupported):
        parse(bad)


def test_refs_walks_the_tree_in_order():
    assert list(refs(parse("SUM(A1:A3)+B1*IF(C1,1,2)"))) == ["A1:A3", "B1", "C1"]
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_parser.py -v`
Expected: FAIL, `ModuleNotFoundError: No module named 'unexcel.parser'`.

- [ ] **Step 3: Write `src/unexcel/parser.py`**

```python
"""Excel formula text -> small AST. Pratt parser over openpyxl's tokenizer."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator, Union

from openpyxl.formula.tokenizer import Token, Tokenizer, TokenizerError

from unexcel.model import Unsupported


@dataclass(frozen=True)
class Num:
    value: float


@dataclass(frozen=True)
class Str:
    value: str


@dataclass(frozen=True)
class Bool:
    value: bool


@dataclass(frozen=True)
class ErrLit:
    value: str


@dataclass(frozen=True)
class RefNode:
    text: str


@dataclass(frozen=True)
class Unary:
    op: str
    operand: "Node"


@dataclass(frozen=True)
class Postfix:
    op: str
    operand: "Node"


@dataclass(frozen=True)
class Binary:
    op: str
    left: "Node"
    right: "Node"


@dataclass(frozen=True)
class Call:
    name: str
    args: tuple  # Node or None (empty argument)


Node = Union[Num, Str, Bool, ErrLit, RefNode, Unary, Postfix, Binary, Call]

# Excel precedence, low -> high. Unary minus binds tighter than ^, so -2^2 = 4.
BINARY = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}
PREFIX = 6


def parse(formula: str) -> Node:
    """Parse formula text without the leading '='."""
    try:
        tokens = [t for t in Tokenizer("=" + formula).items if t.type != Token.WSPACE]
    except TokenizerError as err:
        raise Unsupported(f"cannot parse formula: {err}") from None
    p = _Parser(tokens)
    node = p.expr(0)
    if p.peek() is not None:
        raise Unsupported(f"unexpected {p.peek().value!r} in formula")
    return node


def refs(node) -> Iterator[str]:
    """Every reference text in the tree, left to right."""
    if isinstance(node, RefNode):
        yield node.text
    elif isinstance(node, (Unary, Postfix)):
        yield from refs(node.operand)
    elif isinstance(node, Binary):
        yield from refs(node.left)
        yield from refs(node.right)
    elif isinstance(node, Call):
        for arg in node.args:
            if arg is not None:
                yield from refs(arg)


class _Parser:
    def __init__(self, tokens):
        self.tokens, self.i = tokens, 0

    def peek(self):
        return self.tokens[self.i] if self.i < len(self.tokens) else None

    def take(self):
        tok = self.peek()
        if tok is None:
            raise Unsupported("formula ends unexpectedly")
        self.i += 1
        return tok

    def at(self, type_, subtype):
        tok = self.peek()
        return tok is not None and tok.type == type_ and tok.subtype == subtype

    def expr(self, min_bp):
        left = self.atom()
        while (tok := self.peek()) is not None:
            if tok.type == Token.OP_POST:
                self.take()
                left = Postfix(tok.value, left)
                continue
            if tok.type != Token.OP_IN:
                break
            bp = BINARY.get(tok.value)
            if bp is None:
                raise Unsupported(f"operator {tok.value!r} not supported")
            if bp <= min_bp:
                break
            self.take()
            left = Binary(tok.value, left, self.expr(bp))
        return left

    def atom(self):
        tok = self.take()
        if tok.type == Token.OP_PRE:
            operand = self.expr(PREFIX)
            return operand if tok.value == "+" else Unary("-", operand)
        if tok.type == Token.OPERAND:
            if tok.subtype == Token.NUMBER:
                return Num(float(tok.value))
            if tok.subtype == Token.TEXT:
                return Str(tok.value[1:-1].replace('""', '"'))
            if tok.subtype == Token.LOGICAL:
                return Bool(tok.value.upper() == "TRUE")
            if tok.subtype == Token.ERROR:
                return ErrLit(tok.value)
            return RefNode(tok.value)
        if tok.type == Token.PAREN and tok.subtype == Token.OPEN:
            inner = self.expr(0)
            if not self.at(Token.PAREN, Token.CLOSE):
                raise Unsupported("missing ')'")
            self.take()
            return inner
        if tok.type == Token.FUNC and tok.subtype == Token.OPEN:
            return self.call(tok.value[:-1].upper().removeprefix("_XLFN."))
        raise Unsupported(f"unexpected {tok.value!r} in formula")

    def call(self, name):
        if self.at(Token.FUNC, Token.CLOSE):
            self.take()
            return Call(name, ())
        args = []
        while True:
            if self.at(Token.SEP, Token.ARG) or self.at(Token.FUNC, Token.CLOSE):
                args.append(None)
            else:
                args.append(self.expr(0))
            tok = self.take()
            if tok.type == Token.FUNC and tok.subtype == Token.CLOSE:
                return Call(name, tuple(args))
            if not (tok.type == Token.SEP and tok.subtype == Token.ARG):
                raise Unsupported(f"unexpected {tok.value!r} in {name}()")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_parser.py -v`
Expected: 8 passed.

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/parser.py tests/test_parser.py
git commit -m "feat: Excel formula parser with Excel operator precedence" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Reference resolution

**Files:**
- Create: `src/unexcel/refs.py`
- Test: `tests/test_refs.py`

**Interfaces:**
- Consumes: `model.Book`, `model.Key`, `model.Unsupported`.
- Produces: `Area(sheet, r1, c1, r2, c2)` with `.key` (top-left `Key`), `.is_single`, `.cells() -> Iterator[Key]` (row-major); `resolve(text: str, sheet: str, book: Book) -> Area`.

- [ ] **Step 1: Write the failing tests `tests/test_refs.py`**

```python
import pytest

from unexcel.model import Book, Unsupported
from unexcel.refs import Area, resolve

BOOK = Book(cells={}, sheets=["S", "My Sheet", "Bob's"], names={"RATE": ("My Sheet", "B2")},
            max_row={"S": 10, "My Sheet": 5, "Bob's": 1})


def test_resolve_forms():
    assert resolve("A1", "S", BOOK) == Area("S", 1, 1, 1, 1)
    assert resolve("$B$2:C3", "S", BOOK) == Area("S", 2, 2, 3, 3)
    assert resolve("'My Sheet'!A1", "S", BOOK) == Area("My Sheet", 1, 1, 1, 1)
    assert resolve("'Bob''s'!A1", "S", BOOK) == Area("Bob's", 1, 1, 1, 1)
    assert resolve("C:C", "S", BOOK) == Area("S", 1, 3, 10, 3)
    assert resolve("rate", "S", BOOK) == Area("My Sheet", 2, 2, 2, 2)


@pytest.mark.parametrize("text, msg", [
    ("[Book2.xlsx]S!A1", "external link"),
    ("Nope!A1", "unknown sheet"),
    ("1:1", "whole-row"),
    ("Missing", "unknown name"),
])
def test_resolve_rejects(text, msg):
    with pytest.raises(Unsupported, match=msg):
        resolve(text, "S", BOOK)


def test_area_helpers():
    area = Area("S", 1, 1, 2, 2)
    assert list(area.cells()) == [("S", 1, 1), ("S", 1, 2), ("S", 2, 1), ("S", 2, 2)]
    assert area.key == ("S", 1, 1) and not area.is_single
    assert Area("S", 3, 3, 3, 3).is_single
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_refs.py -v`
Expected: FAIL, `ModuleNotFoundError: No module named 'unexcel.refs'`.

- [ ] **Step 3: Write `src/unexcel/refs.py`**

```python
"""Reference text ('Sales'!C:C, $A$1, Rate) -> a concrete rectangle of cells."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator

from openpyxl.utils.cell import range_boundaries

from unexcel.model import Book, Key, Unsupported


@dataclass(frozen=True)
class Area:
    sheet: str
    r1: int
    c1: int
    r2: int
    c2: int

    @property
    def key(self) -> Key:
        return (self.sheet, self.r1, self.c1)

    @property
    def is_single(self) -> bool:
        return self.r1 == self.r2 and self.c1 == self.c2

    def cells(self) -> Iterator[Key]:
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                yield (self.sheet, r, c)


def resolve(text: str, sheet: str, book: Book) -> Area:
    ref = text
    if "!" in text:
        sheet, ref = text.rsplit("!", 1)
        if sheet.startswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
        if "[" in sheet:
            raise Unsupported(f"external link {text}")
    elif ref.upper() in book.names:
        sheet, ref = book.names[ref.upper()]
    if sheet not in book.sheets:
        raise Unsupported(f"unknown sheet {sheet!r}")
    try:
        c1, r1, c2, r2 = range_boundaries(ref.replace("$", ""))
    except ValueError:
        raise Unsupported(f"unknown name {text!r}") from None
    if c1 is None:
        raise Unsupported(f"whole-row reference {text}")
    return Area(sheet, r1 or 1, c1, r2 or book.max_row.get(sheet, 1), c2)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_refs.py -v`
Expected: 6 passed.

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/refs.py tests/test_refs.py
git commit -m "feat: resolve cell, range, sheet and named references" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: Dependency graph

**Files:**
- Create: `src/unexcel/graph.py`
- Test: `tests/test_graph.py`

**Interfaces:**
- Consumes: `parser.parse`, `parser.refs`, `refs.resolve`, `Area.cells()`.
- Produces: `graph.CYCLE: str`; `formula_order(book) -> tuple[list[Key], dict[Key, str]]`. The first item is the formula cells in dependency order (ties broken by key order). The second is cells in or behind a cycle, mapped to `CYCLE`.

- [ ] **Step 1: Write the failing tests `tests/test_graph.py`**

```python
from xlsxkit import book_of

from unexcel.graph import CYCLE, formula_order


def test_dependencies_come_first():
    book = book_of({"S": {"A1": 1, "A2": ("=A3+A1", 3), "A3": ("=A1*2", 2)}})
    assert formula_order(book) == ([("S", 3, 1), ("S", 2, 1)], {})


def test_cycles_and_their_dependents_are_unresolved():
    book = book_of({"S": {"A1": ("=B1", 0), "B1": ("=A1", 0), "C1": ("=A1+1", 1), "D1": ("=2", 2)}})
    order, bad = formula_order(book)
    assert order == [("S", 1, 4)]
    assert bad == {("S", 1, 1): CYCLE, ("S", 1, 2): CYCLE, ("S", 1, 3): CYCLE}


def test_self_reference_is_a_cycle():
    assert formula_order(book_of({"S": {"A1": ("=A1+1", 0)}})) == ([], {("S", 1, 1): CYCLE})


def test_unparseable_formula_is_still_ordered():
    assert formula_order(book_of({"S": {"A1": ("=SUM(", 0)}})) == ([("S", 1, 1)], {})
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_graph.py -v`
Expected: FAIL, `ModuleNotFoundError: No module named 'unexcel.graph'`.

- [ ] **Step 3: Write `src/unexcel/graph.py`**

```python
"""Order formula cells so every cell comes after the cells it reads (Kahn's algorithm)."""
from __future__ import annotations

import heapq
from collections import defaultdict

from unexcel.model import Book, Key, Unsupported
from unexcel.parser import parse, refs
from unexcel.refs import resolve

CYCLE = "circular reference (or depends on one)"


def formula_order(book: Book) -> tuple[list[Key], dict[Key, str]]:
    formulas = {k for k, c in book.cells.items() if c.formula is not None}
    deps: dict[Key, set[Key]] = {}
    for key in formulas:
        cell = book.cells[key]
        found: set[Key] = set()
        try:
            for text in refs(parse(cell.formula)):
                found.update(resolve(text, cell.sheet, book).cells())
        except Unsupported:
            pass  # the compiler reports the real reason for this cell
        deps[key] = found & formulas

    users = defaultdict(set)
    for key, ds in deps.items():
        for d in ds:
            users[d].add(key)
    waiting = {k: len(ds) for k, ds in deps.items()}
    ready = [k for k, n in waiting.items() if n == 0]
    heapq.heapify(ready)
    order = []
    while ready:
        key = heapq.heappop(ready)
        order.append(key)
        for user in users[key]:
            waiting[user] -= 1
            if waiting[user] == 0:
                heapq.heappush(ready, user)
    return order, {k: CYCLE for k in sorted(formulas - set(order))}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_graph.py -v`
Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/graph.py tests/test_graph.py
git commit -m "feat: dependency order with cycle detection" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: Compiler core (operators, references, model SQL) + DuckDB execution

**Files:**
- Create: `src/unexcel/expr.py`, `src/unexcel/functions.py` (empty registry for now), `src/unexcel/compiler.py`, `src/unexcel/verifier.py` (`execute` only)
- Test: `tests/test_compiler.py`

**Interfaces:**
- Consumes: `parse`, `resolve`, `Area`, `formula_order`, `Book`, `Cell`, `Unsupported`, `a1`.
- Produces:
  - `expr.Expr(sql: str, type: "num"|"txt"|"bool"|"blank")`, `expr.INPUT = "__unexcel_input__"`, `expr.ZERO`, `expr.BLANK`, `expr.lit_num(v)`, `expr.lit_str(v)`, `expr.num(e)`, `expr.txt(e)`, `expr.as_bool(e)`, `expr.unify(a, b) -> (Expr, Expr)`, `expr.finalize(e)`
  - `functions.FUNCS: dict[str, Callable[[Compiler, tuple, str], Expr]]`, `functions.register(name)`
  - `compiler.Compiler(book)` with `.book`, `.findings: list[str]`, `.expr(node, sheet) -> Expr`, `.cell_value(key) -> Expr`
  - `compiler.Compiled(ctes: list[tuple[str, Key, Expr]], unresolved: dict[Key, str], findings: list[str])`
  - `compiler.compile_book(book) -> Compiled`, `compiler.render_model(compiled, input_table: str) -> str`, `compiler.input_rows(book) -> list[tuple]`, `compiler.INPUT_COLUMNS: str`, `compiler.VOLATILE: set[str]`
  - `verifier.execute(book, compiled) -> dict[Key, tuple[kind, num|None, txt|None]]`

- [ ] **Step 1: Write the failing tests `tests/test_compiler.py`**

```python
import pytest
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute


def run(sheets, names=None):
    book = book_of(sheets, names)
    compiled = compile_book(book)
    return execute(book, compiled), compiled.unresolved


def test_arithmetic_precedence_percent_and_divide_by_zero():
    got, bad = run({"S": {"A1": 2, "A2": ("=-A1^2+10%*A1", 4.2), "A3": ("=A1/0", "#DIV/0!")}})
    assert bad == {}
    kind, num, txt = got[("S", 2, 1)]
    assert kind == "n" and num == pytest.approx(4.2) and txt is None
    assert got[("S", 3, 1)] == ("n", None, None)  # error, not inf


def test_text_concat_and_case_insensitive_compare():
    got, _ = run({"S": {"A1": "Hi", "A2": ('=A1&" there"', "Hi there"), "A3": ('=A1="hi"', True)}})
    assert got[("S", 2, 1)] == ("s", None, "Hi there")
    assert got[("S", 3, 1)] == ("b", 1.0, None)


def test_formula_chains_and_blank_cells():
    got, _ = run({"S": {"A1": 3, "A2": ("=A1*2", 6), "A3": ("=A2+B9", 6), "A4": ("=B9", 0)}})
    assert got[("S", 3, 1)] == ("n", 6.0, None)
    assert got[("S", 4, 1)] == ("n", 0.0, None)


def test_named_ranges():
    got, bad = run({"S": {"A1": 0.1, "A2": ("=Rate*100", 10)}}, names={"Rate": ("S", "$A$1")})
    assert bad == {} and got[("S", 2, 1)][1] == pytest.approx(10.0)


def test_sheet_names_with_quotes_and_spaces():  # Review Focus 3
    got, bad = run({"Bob's Sheet": {"A1": 5}, "S": {"A1": ("='Bob''s Sheet'!A1*2", 10)}})
    assert bad == {} and got[("S", 1, 1)] == ("n", 10.0, None)


def test_unresolved_reasons():  # includes Review Focus 5 (no saved value)
    _, bad = run({"S": {
        "A1": "x",
        "A2": ("=A1+1", "#VALUE!"),
        "A3": ("=NOW()", 1),
        "A4": ("=A3+1", 2),
        "A5": ("=1", None),
        "A6": ("=FOO(1)", 1),
    }})
    assert bad[("S", 2, 1)] == "text used as a number"
    assert bad[("S", 3, 1)] == "volatile function NOW"
    assert bad[("S", 4, 1)] == "depends on unresolved cell S!A3"
    assert bad[("S", 5, 1)].startswith("no saved value")
    assert bad[("S", 6, 1)] == "function FOO not supported yet"


def test_workbook_without_formulas_gives_empty_result():
    got, bad = run({"S": {"A1": 1}})
    assert got == {} and bad == {}
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_compiler.py -v`
Expected: FAIL, `ModuleNotFoundError: No module named 'unexcel.compiler'`.

- [ ] **Step 3: Write `src/unexcel/expr.py`**

```python
"""Typed SQL expressions and Excel-style type coercion."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from unexcel.model import Unsupported

INPUT = "__unexcel_input__"  # replaced by the real input table when SQL is rendered
Type = Literal["num", "txt", "bool", "blank"]


@dataclass(frozen=True)
class Expr:
    sql: str
    type: Type


ZERO = Expr("CAST(0 AS DOUBLE)", "num")
BLANK = Expr("NULL", "blank")  # an empty cell; always coerced before use, never rendered raw


def lit_num(v: float) -> Expr:
    return Expr(f"CAST({v!r} AS DOUBLE)", "num")  # bare literals would be DECIMAL in DuckDB


def lit_str(v: str) -> Expr:
    return Expr("'" + v.replace("'", "''") + "'", "txt")


def num(e: Expr) -> Expr:
    if e.type == "num":
        return e
    if e.type == "blank":
        return ZERO
    if e.type == "bool":
        return Expr(f"CAST({e.sql} AS DOUBLE)", "num")
    raise Unsupported("text used as a number")


def txt(e: Expr) -> Expr:
    if e.type == "txt":
        return e
    if e.type == "blank":
        return Expr("''", "txt")
    raise Unsupported("number used as text")


def as_bool(e: Expr) -> Expr:
    if e.type == "bool":
        return e
    if e.type == "blank":
        return Expr("FALSE", "bool")
    if e.type == "num":
        return Expr(f"({e.sql} <> 0)", "bool")
    raise Unsupported("text used as TRUE/FALSE")


COERCE = {"num": num, "txt": txt, "bool": as_bool}


def unify(a: Expr, b: Expr) -> tuple[Expr, Expr]:
    """Give two expressions one type (IF branches, IFERROR, comparisons)."""
    if a.type == "blank" and b.type == "blank":
        return ZERO, ZERO
    if a.type == "blank":
        a = COERCE[b.type](a)
    if b.type == "blank":
        b = COERCE[a.type](b)
    if a.type != b.type:
        raise Unsupported(f"mixes {a.type} and {b.type}")
    return a, b


def finalize(e: Expr) -> Expr:
    return ZERO if e.type == "blank" else e  # =A1 on an empty cell shows 0 in Excel
```

- [ ] **Step 4: Write `src/unexcel/functions.py`** (registry only; Tasks 6–7 add functions)

```python
"""Excel function -> SQL. Each entry: fn(compiler, args, sheet) -> Expr."""
from __future__ import annotations

FUNCS = {}


def register(name):
    def deco(fn):
        FUNCS[name] = fn
        return fn

    return deco
```

- [ ] **Step 5: Write `src/unexcel/compiler.py`**

```python
"""Formula AST -> SQL. Cell mode: every formula cell becomes one materialized CTE."""
from __future__ import annotations

from dataclasses import dataclass

from unexcel.expr import BLANK, INPUT, Expr, finalize, lit_num, lit_str, num, txt, unify
from unexcel.functions import FUNCS
from unexcel.graph import formula_order
from unexcel.model import Book, Cell, Key, Unsupported, a1
from unexcel.parser import Binary, Bool, Call, ErrLit, Num, Postfix, RefNode, Str, Unary, parse
from unexcel.refs import resolve

INPUT_COLUMNS = "sheet VARCHAR, row_num INTEGER, col_num INTEGER, kind VARCHAR, num DOUBLE, txt VARCHAR"
VOLATILE = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT"}
ARITH = {"+", "-", "*"}
KIND = {"num": "n", "txt": "s", "bool": "b"}


@dataclass
class Compiled:
    ctes: list[tuple[str, Key, Expr]]
    unresolved: dict[Key, str]
    findings: list[str]


class Compiler:
    def __init__(self, book: Book):
        self.book = book
        self.done: dict[Key, tuple[str, str]] = {}  # formula cell -> (cte name, type)
        self.findings: list[str] = []

    def compile_cell(self, cell: Cell) -> Expr:
        if cell.array:
            raise Unsupported("array formula")
        return finalize(self.expr(parse(cell.formula), cell.sheet))

    def expr(self, node, sheet: str) -> Expr:
        if node is None:
            return BLANK
        if isinstance(node, Num):
            return lit_num(node.value)
        if isinstance(node, Str):
            return lit_str(node.value)
        if isinstance(node, Bool):
            return Expr("TRUE" if node.value else "FALSE", "bool")
        if isinstance(node, ErrLit):
            raise Unsupported(f"error literal {node.value}")
        if isinstance(node, RefNode):
            area = resolve(node.text, sheet, self.book)
            if not area.is_single:
                raise Unsupported(f"range {node.text} used as a single value")
            return self.cell_value(area.key)
        if isinstance(node, Unary):
            return Expr(f"(-{num(self.expr(node.operand, sheet)).sql})", "num")
        if isinstance(node, Postfix):
            return Expr(f"({num(self.expr(node.operand, sheet)).sql} / 100)", "num")
        if isinstance(node, Binary):
            return self.binary(node, sheet)
        if isinstance(node, Call):
            if node.name in VOLATILE:
                raise Unsupported(f"volatile function {node.name}")
            fn = FUNCS.get(node.name)
            if fn is None:
                raise Unsupported(f"function {node.name} not supported yet")
            return fn(self, node.args, sheet)
        raise Unsupported(f"cannot translate {node!r}")

    def binary(self, node: Binary, sheet: str) -> Expr:
        a, b, op = self.expr(node.left, sheet), self.expr(node.right, sheet), node.op
        if op in ARITH:
            return Expr(f"({num(a).sql} {op} {num(b).sql})", "num")
        if op == "/":
            return Expr(f"({num(a).sql} / NULLIF({num(b).sql}, 0))", "num")  # DuckDB gives inf, Excel #DIV/0!
        if op == "^":
            return Expr(f"POWER({num(a).sql}, {num(b).sql})", "num")
        if op == "&":
            return Expr(f"({txt(a).sql} || {txt(b).sql})", "txt")
        a, b = unify(a, b)
        if a.type == "txt":  # Excel text comparison ignores case
            return Expr(f"(LOWER({a.sql}) {op} LOWER({b.sql}))", "bool")
        return Expr(f"({a.sql} {op} {b.sql})", "bool")

    def cell_value(self, key: Key) -> Expr:
        cell = self.book.cells.get(key)
        if cell is None or (cell.formula is None and cell.kind == "blank"):
            return BLANK
        if cell.formula is not None:
            if key not in self.done:
                raise Unsupported(f"depends on unresolved cell {a1(key)}")
            name, typ = self.done[key]
            return Expr(f"(SELECT v FROM {name})", typ)
        sheet, row, col = key
        where = f"FROM {INPUT} WHERE sheet = {lit_str(sheet).sql} AND row_num = {row} AND col_num = {col}"
        if cell.kind == "n":
            return Expr(f"(SELECT num {where})", "num")
        if cell.kind == "s":
            return Expr(f"(SELECT txt {where})", "txt")
        if cell.kind == "b":
            return Expr(f"((SELECT num {where}) <> 0)", "bool")
        raise Unsupported(f"input cell {a1(key)} holds an error")


def compile_book(book: Book) -> Compiled:
    order, unresolved = formula_order(book)
    comp = Compiler(book)
    ctes = []
    for key in order:
        cell = book.cells[key]
        if cell.kind == "blank":
            unresolved[key] = "no saved value in the file - open and save it in Excel"
            continue
        try:
            e = comp.compile_cell(cell)
        except Unsupported as err:
            unresolved[key] = str(err)
            continue
        name = f"c{len(ctes) + 1}"
        comp.done[key] = (name, e.type)
        ctes.append((name, key, e))
    return Compiled(ctes, unresolved, comp.findings)


def render_model(compiled: Compiled, input_table: str) -> str:
    if not compiled.ctes:
        return ("SELECT CAST(NULL AS VARCHAR) AS sheet, CAST(NULL AS INTEGER) AS row_num, "
                "CAST(NULL AS INTEGER) AS col_num, CAST(NULL AS VARCHAR) AS kind, "
                "CAST(NULL AS DOUBLE) AS num, CAST(NULL AS VARCHAR) AS txt WHERE FALSE")
    ctes = ",\n".join(f"{name} AS MATERIALIZED (SELECT {e.sql} AS v)" for name, _, e in compiled.ctes)
    rows = []
    for name, (sheet, row, col), e in compiled.ctes:
        v = f"(SELECT v FROM {name})"
        num_sql = "CAST(NULL AS DOUBLE)" if e.type == "txt" else f"CAST({v} AS DOUBLE)"
        txt_sql = f"CAST({v} AS VARCHAR)" if e.type == "txt" else "CAST(NULL AS VARCHAR)"
        rows.append(f"SELECT {lit_str(sheet).sql} AS sheet, {row} AS row_num, {col} AS col_num, "
                    f"'{KIND[e.type]}' AS kind, {num_sql} AS num, {txt_sql} AS txt")
    return (f"WITH\n{ctes}\n" + "\nUNION ALL\n".join(rows)).replace(INPUT, input_table)


def input_rows(book: Book) -> list[tuple]:
    """Every non-formula cell as a row of the input table (see INPUT_COLUMNS)."""
    rows = []
    for (sheet, row, col), cell in sorted(book.cells.items()):
        if cell.formula is not None or cell.kind == "blank":
            continue
        if cell.kind == "b":
            num_v, txt_v = (1.0 if cell.value else 0.0), None
        elif cell.kind == "n":
            num_v, txt_v = cell.value, None
        else:
            num_v, txt_v = None, cell.value
        rows.append((sheet, row, col, cell.kind, num_v, txt_v))
    return rows
```

- [ ] **Step 6: Write `src/unexcel/verifier.py`** (execution only; comparison comes in Task 8)

```python
"""Run the compiled SQL on DuckDB and compare every cell with Excel's saved value."""
from __future__ import annotations

import duckdb

from unexcel.compiler import INPUT_COLUMNS, Compiled, input_rows, render_model
from unexcel.model import Book, Key


def execute(book: Book, compiled: Compiled) -> dict[Key, tuple]:
    con = duckdb.connect()
    con.execute(f"CREATE TABLE input_cells ({INPUT_COLUMNS})")
    rows = input_rows(book)
    if rows:
        # ponytail: executemany is slow past ~100k rows; switch to Arrow/CSV bulk load in Plan 2 perf work
        con.executemany("INSERT INTO input_cells VALUES (?, ?, ?, ?, ?, ?)", rows)
    result = con.execute(render_model(compiled, "input_cells")).fetchall()
    return {(s, r, c): (k, n, t) for s, r, c, k, n, t in result}
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `python -m pytest tests/test_compiler.py -v`
Expected: 7 passed.

- [ ] **Step 8: Commit**

```bash
git add src/unexcel/expr.py src/unexcel/functions.py src/unexcel/compiler.py src/unexcel/verifier.py tests/test_compiler.py
git commit -m "feat: cell-mode compiler for operators and references, DuckDB execution" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: SUM, IF, ROUND, AND, OR, IFERROR (+ numbers stored as text finding)

**Files:**
- Modify: `src/unexcel/functions.py` (replace whole file)
- Test: `tests/test_functions.py`

**Interfaces:**
- Consumes: `Compiler.expr`, `Compiler.cell_value`, `Compiler.book`, `Compiler.findings`, everything in `expr.py`, `resolve`, `Area`.
- Produces: `FUNCS` entries `SUM IF ROUND AND OR IFERROR`; helpers `functions.range_numbers(c, area) -> list[str]`, `functions.note_text_numbers(c, area)`, `functions.arity(name, args, lo, hi)`. Finding text format: `"{Sheet!A1}: number stored as text ('15,000') is skipped by Excel"`.

- [ ] **Step 1: Write the failing tests `tests/test_functions.py`**

```python
import pytest
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute


def run(sheets):
    book = book_of(sheets)
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def test_sum_mixes_ranges_scalars_and_formula_cells():
    got, c = run({"S": {"A1": 1, "A2": 2, "A3": "note", "A4": True, "A5": ("=A1*10", 10),
                        "B1": ("=SUM(A1:A5, 100)", 113)}})
    assert c.unresolved == {} and got[("S", 1, 2)] == ("n", 113.0, None)


def test_sum_propagates_errors_and_handles_empty_ranges():
    got, _ = run({"S": {"A1": ("=1/0", "#DIV/0!"), "B1": ("=SUM(A1:A2)", "#DIV/0!"), "C1": ("=SUM(Z1:Z9)", 0)}})
    assert got[("S", 1, 2)] == ("n", None, None)
    assert got[("S", 1, 3)] == ("n", 0.0, None)


def test_sum_reports_numbers_stored_as_text():
    got, c = run({"S": {"A1": 5, "A2": "15,000", "A3": ("=SUM(A1:A2)", 5)}})
    assert got[("S", 3, 1)] == ("n", 5.0, None)
    assert c.findings == ["S!A2: number stored as text ('15,000') is skipped by Excel"]


def test_if_round_and_or_iferror():
    got, c = run({"S": {
        "A1": 120000, "A2": 2.675,
        "B1": ("=IF(A1>100000, A1*5%, A1*2%)", 6000),
        "B2": ("=ROUND(A2, 2)", 2.68),
        "B3": ("=AND(A1>0, OR(A2<0, TRUE))", True),
        "B4": ("=IFERROR(A1/0, -1)", -1),
        "B5": ('=IF(A1>0, "big", "small")', "big"),
    }})
    assert c.unresolved == {}
    assert got[("S", 1, 2)][1] == pytest.approx(6000)
    assert got[("S", 2, 2)][1] == pytest.approx(2.68)
    assert got[("S", 3, 2)] == ("b", 1.0, None)
    assert got[("S", 4, 2)] == ("n", -1.0, None)
    assert got[("S", 5, 2)] == ("s", None, "big")


def test_if_with_different_branch_types_is_unresolved():
    _, c = run({"S": {"A1": 1, "B1": ('=IF(A1>0, "yes", 0)', "yes"), "B2": ("=IF(A1)", 0)}})
    assert c.unresolved[("S", 1, 2)] == "mixes txt and num"
    assert c.unresolved[("S", 2, 2)] == "IF needs 2-3 arguments, got 1"
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_functions.py -v`
Expected: FAIL. Cells come back `UNRESOLVED` with "function SUM not supported yet" (e.g. `KeyError: ('S', 1, 2)` on `got[...]`).

- [ ] **Step 3: Replace `src/unexcel/functions.py`**

```python
"""Excel function -> SQL. Each entry: fn(compiler, args, sheet) -> Expr."""
from __future__ import annotations

from unexcel.expr import INPUT, ZERO, Expr, as_bool, lit_str, num, unify
from unexcel.model import Unsupported, a1
from unexcel.parser import RefNode
from unexcel.refs import Area, resolve

FUNCS = {}


def register(name):
    def deco(fn):
        FUNCS[name] = fn
        return fn

    return deco


def arity(name, args, lo, hi):
    if not lo <= len(args) <= hi:
        raise Unsupported(f"{name} needs {lo}-{hi} arguments, got {len(args)}")


def _looks_numeric(text: str) -> bool:
    try:
        float(text.replace(",", "").strip())
        return True
    except ValueError:
        return False


def note_text_numbers(c, area: Area):
    """Excel skips text in SUM-like ranges, even text that looks like a number. Tell the user."""
    for key in area.cells():
        cell = c.book.cells.get(key)
        if cell is not None and cell.formula is None and cell.kind == "s" and _looks_numeric(cell.value):
            c.findings.append(f"{a1(key)}: number stored as text ({cell.value!r}) is skipped by Excel")


def range_numbers(c, area: Area) -> list[str]:
    """SQL terms for the numbers in a range as SUM sees them: text, TRUE/FALSE and blanks are skipped."""
    terms, has_inputs = [], False
    for key in area.cells():
        cell = c.book.cells.get(key)
        if cell is None:
            continue
        if cell.formula is not None:
            value = c.cell_value(key)
            if value.type == "num":
                terms.append(value.sql)
        elif cell.kind == "n":
            has_inputs = True
        elif cell.kind == "e":
            raise Unsupported(f"input cell {a1(key)} holds an error")
    note_text_numbers(c, area)
    if has_inputs:
        terms.insert(0, f"(SELECT SUM(num) FROM {INPUT} WHERE sheet = {lit_str(area.sheet).sql} AND kind = 'n' "
                        f"AND row_num BETWEEN {area.r1} AND {area.r2} AND col_num BETWEEN {area.c1} AND {area.c2})")
    return terms


@register("SUM")
def sum_(c, args, sheet):
    parts = []
    for arg in args:
        if isinstance(arg, RefNode):
            parts.extend(range_numbers(c, resolve(arg.text, sheet, c.book)))
        else:
            parts.append(num(c.expr(arg, sheet)).sql)
    return Expr("(" + " + ".join(parts) + ")", "num") if parts else ZERO


@register("IF")
def if_(c, args, sheet):
    arity("IF", args, 2, 3)
    cond = as_bool(c.expr(args[0], sheet))
    yes = c.expr(args[1], sheet)
    no = c.expr(args[2], sheet) if len(args) == 3 else Expr("FALSE", "bool")
    yes, no = unify(yes, no)
    return Expr(f"(CASE WHEN ({cond.sql}) IS NULL THEN NULL WHEN {cond.sql} THEN {yes.sql} ELSE {no.sql} END)",
                yes.type)


@register("ROUND")
def round_(c, args, sheet):
    arity("ROUND", args, 2, 2)
    x, digits = num(c.expr(args[0], sheet)), num(c.expr(args[1], sheet))
    # DuckDB ROUND on DOUBLE rounds half away from zero like Excel (2.675 -> 2.68, verified on 1.5.5)
    return Expr(f"ROUND({x.sql}, CAST(TRUNC({digits.sql}) AS INTEGER))", "num")


def _logic(name, joiner):
    def fn(c, args, sheet):
        if not args:
            raise Unsupported(f"{name} needs arguments")
        parts = [as_bool(c.expr(a, sheet)).sql for a in args]
        return Expr("(" + f" {joiner} ".join(parts) + ")", "bool")

    return fn


FUNCS["AND"] = _logic("AND", "AND")
FUNCS["OR"] = _logic("OR", "OR")


@register("IFERROR")
def iferror(c, args, sheet):
    arity("IFERROR", args, 2, 2)
    value, fallback = unify(c.expr(args[0], sheet), c.expr(args[1], sheet))
    return Expr(f"COALESCE({value.sql}, {fallback.sql})", value.type)  # SQL NULL = Excel error
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_functions.py tests/test_compiler.py -v`
Expected: all passed (5 new + 7 from Task 5).

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/functions.py tests/test_functions.py
git commit -m "feat: SUM IF ROUND AND OR IFERROR with text-number findings" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: SUMIFS and VLOOKUP (exact match)

**Files:**
- Modify: `src/unexcel/functions.py` (append)
- Modify: `tests/xlsxkit.py` (append `SALES`)
- Test: `tests/test_lookup.py`

**Interfaces:**
- Consumes: `functions.arity`, `functions.note_text_numbers`, `Compiler.expr`, `expr.*`, `resolve`, `Area`, `parser.Bool/Num/RefNode/Str`.
- Produces: `FUNCS` entries `SUMIFS VLOOKUP`; `xlsxkit.SALES` (Sales sheet spec reused by the CLI tests).

- [ ] **Step 1: Append `SALES` to `tests/xlsxkit.py`**

```python
# The commission example from the design: header row, mixed-case names, one number stored as text.
SALES = {
    "A1": "Rep", "B1": "Region", "C1": "Amount",
    "A2": "Asha", "B2": "West", "C2": 80000,
    "A3": "asha", "B3": "West", "C3": 40000,
    "A4": "Vikram", "B4": "South", "C4": 60000,
    "A5": "Vikram", "B5": "South", "C5": "15,000",
}
```

- [ ] **Step 2: Write the failing tests `tests/test_lookup.py`** (Review Focus 4: whole columns with header rows)

```python
from xlsxkit import SALES, book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute


def run(sheets):
    book = book_of(sheets)
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def test_sumifs_whole_columns_case_insensitive_and_skips_text():
    got, c = run({"Sales": SALES, "Commission": {
        "A1": "Asha", "B1": ("=SUMIFS(Sales!C:C, Sales!A:A, A1)", 120000),
        "A2": "Vikram", "B2": ('=SUMIFS(Sales!C:C, Sales!A:A, A2, Sales!B:B, "south")', 60000),
    }})
    assert c.unresolved == {}
    assert got[("Commission", 1, 2)] == ("n", 120000.0, None)
    assert got[("Commission", 2, 2)] == ("n", 60000.0, None)
    assert "Sales!C5: number stored as text ('15,000') is skipped by Excel" in c.findings


def test_sumifs_rejects_what_it_cannot_prove():
    _, c = run({"Sales": SALES, "X": {
        "A1": ('=SUMIFS(Sales!C:C, Sales!A:A, ">5")', 0),
        "A2": ('=SUMIFS(Sales!C2:C5, Sales!A2:A3, "Asha")', 0),
        "A3": ("=SUMIFS(Sales!C:C, Sales!A:A)", 0),
    }})
    assert "comparison/wildcard" in c.unresolved[("X", 1, 1)]
    assert "same height" in c.unresolved[("X", 2, 1)]
    assert "criteria pairs" in c.unresolved[("X", 3, 1)]


def test_vlookup_exact_with_header_row_and_missing_key():
    rates = {"A1": "Region", "B1": "Rate", "A2": "West", "B2": 0.05, "A3": "South", "B3": 0.02}
    got, c = run({"Rates": rates, "X": {
        "A1": "south", "B1": ("=VLOOKUP(A1, Rates!A:B, 2, FALSE)", 0.02),
        "A2": "North", "B2": ("=VLOOKUP(A2, Rates!A:B, 2, FALSE)", "#N/A"),
        "B3": ("=VLOOKUP(A1, Rates!A:B, 2)", 0.02),
        "B4": ("=VLOOKUP(A1, Rates!A:B, 3, FALSE)", "#REF!"),
    }})
    assert got[("X", 1, 2)] == ("n", 0.02, None)
    assert got[("X", 2, 2)] == ("n", None, None)
    assert c.unresolved[("X", 3, 2)] == "VLOOKUP approximate match not supported yet"
    assert c.unresolved[("X", 4, 2)] == "VLOOKUP column index outside the table"
```

- [ ] **Step 3: Run to verify it fails**

Run: `python -m pytest tests/test_lookup.py -v`
Expected: FAIL, `KeyError` on `got[...]`, because the cells are UNRESOLVED with "function SUMIFS not supported yet".

- [ ] **Step 4: Append to `src/unexcel/functions.py`**

```python
from unexcel.parser import Bool, Num, Str  # noqa: E402  (kept next to the functions that use them)


def _input_column(c, node, sheet, fname) -> Area:
    if not isinstance(node, RefNode):
        raise Unsupported(f"{fname} needs a cell range")
    area = resolve(node.text, sheet, c.book)
    for key in area.cells():
        cell = c.book.cells.get(key)
        if cell is not None and cell.formula is not None:
            raise Unsupported(f"{fname} over formula cells ({a1(key)}) not supported yet")
    return area


def _match(alias: str, crit: Expr) -> str:
    if crit.type == "txt":  # Excel matching ignores case
        return f"{alias}.kind = 's' AND LOWER({alias}.txt) = LOWER({crit.sql})"
    if crit.type == "num":
        return f"{alias}.kind = 'n' AND {alias}.num = {crit.sql}"
    raise Unsupported(f"criteria of type {crit.type}")


@register("SUMIFS")
def sumifs(c, args, sheet):
    if len(args) < 3 or len(args) % 2 == 0:
        raise Unsupported("SUMIFS needs a sum range and criteria pairs")
    total = _input_column(c, args[0], sheet, "SUMIFS")
    if total.c1 != total.c2:
        raise Unsupported("SUMIFS sum range wider than one column")
    joins, conds = [], []
    for i in range(1, len(args), 2):
        area = _input_column(c, args[i], sheet, "SUMIFS")
        if area.c1 != area.c2 or area.r2 - area.r1 != total.r2 - total.r1:
            raise Unsupported("SUMIFS ranges must be single columns of the same height")
        node = args[i + 1]
        if isinstance(node, Str) and node.value and (node.value[0] in "<>=" or any(ch in node.value for ch in "*?~")):
            raise Unsupported("SUMIFS comparison/wildcard criteria not supported yet")
        k = f"k{len(joins) + 1}"
        joins.append(f"JOIN {INPUT} {k} ON {k}.sheet = {lit_str(area.sheet).sql} AND {k}.col_num = {area.c1} "
                     f"AND {k}.row_num = s.row_num + {area.r1 - total.r1}")
        conds.append(_match(k, c.expr(node, sheet)))
    note_text_numbers(c, total)
    where = " AND ".join([f"s.sheet = {lit_str(total.sheet).sql}", f"s.col_num = {total.c1}",
                          f"s.row_num BETWEEN {total.r1} AND {total.r2}", "s.kind = 'n'", *conds])
    return Expr(f"(SELECT COALESCE(SUM(s.num), CAST(0 AS DOUBLE)) FROM {INPUT} s {' '.join(joins)} WHERE {where})",
                "num")


@register("VLOOKUP")
def vlookup(c, args, sheet):
    if len(args) == 3:
        raise Unsupported("VLOOKUP approximate match not supported yet")
    arity("VLOOKUP", args, 4, 4)
    exact = args[3]
    if not (exact is None or (isinstance(exact, Bool) and not exact.value)
            or (isinstance(exact, Num) and exact.value == 0)):
        raise Unsupported("VLOOKUP approximate match not supported yet")
    if not isinstance(args[2], Num):
        raise Unsupported("VLOOKUP column index must be a number")
    table = _input_column(c, args[1], sheet, "VLOOKUP")
    col = table.c1 + int(args[2].value) - 1
    if not table.c1 <= col <= table.c2:
        raise Unsupported("VLOOKUP column index outside the table")
    kinds = [cell.kind for r in range(table.r1, table.r2 + 1)
             if (cell := c.book.cells.get((table.sheet, r, col))) is not None and cell.kind != "blank"]
    if len(kinds) > 1 and kinds[0] == "s" and set(kinds[1:]) == {"n"}:
        kinds = kinds[1:]  # text header above a numeric column
    if set(kinds) == {"n"}:
        column, typ, empty = "num", "num", "CAST(0 AS DOUBLE)"
    elif set(kinds) == {"s"}:
        column, typ, empty = "txt", "txt", "''"
    else:
        raise Unsupported("VLOOKUP result column mixes types")
    lookup = c.expr(args[0], sheet)
    return Expr(
        f"(SELECT COALESCE(r.{column}, {empty}) FROM {INPUT} k LEFT JOIN {INPUT} r "
        f"ON r.sheet = k.sheet AND r.row_num = k.row_num AND r.col_num = {col} "
        f"WHERE k.sheet = {lit_str(table.sheet).sql} AND k.col_num = {table.c1} "
        f"AND k.row_num BETWEEN {table.r1} AND {table.r2} AND {_match('k', lookup)} "
        f"ORDER BY k.row_num LIMIT 1)",
        typ,
    )
```

Then move the `from unexcel.parser import Bool, Num, Str` line up into the existing import block (`from unexcel.parser import Bool, Num, RefNode, Str`) and delete the `# noqa` line, so all imports sit at the top of the file.

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 6: Commit**

```bash
git add src/unexcel/functions.py tests/xlsxkit.py tests/test_lookup.py
git commit -m "feat: SUMIFS (equality criteria) and exact-match VLOOKUP" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 8: Verifier verdicts

**Files:**
- Modify: `src/unexcel/verifier.py` (append)
- Test: `tests/test_verifier.py`

**Interfaces:**
- Consumes: `execute`, `Compiled`, `Book`, `Cell`, `Verdict`.
- Produces: `verifier.matches(cell, kind, num, txt) -> bool`; `verifier.verify(book, compiled) -> list[Verdict]` sorted by key. For FAIL/PASS, `actual` is the SQL value (`bool` for kind `b`, `None` for a SQL error).

- [ ] **Step 1: Write the failing tests `tests/test_verifier.py`**

```python
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.model import Cell, Verdict
from unexcel.verifier import matches, verify


def test_pass_fail_and_unresolved():
    book = book_of({"S": {
        "A1": 2,
        "A2": ("=A1*2", 4),
        "A3": ("=A1*2", 5),        # Excel says 5: someone typed over the result
        "A4": ("=1/0", "#DIV/0!"),
        "A5": ("=A1", "#N/A"),
        "A6": ("=NOW()", 1),
    }})
    v = {x.key: x for x in verify(book, compile_book(book))}
    assert v[("S", 2, 1)].status == "PASS"
    assert (v[("S", 3, 1)].status, v[("S", 3, 1)].expected, v[("S", 3, 1)].actual) == ("FAIL", 5.0, 4.0)
    assert v[("S", 4, 1)].status == "PASS"
    assert (v[("S", 5, 1)].status, v[("S", 5, 1)].actual) == ("FAIL", 2.0)
    assert v[("S", 6, 1)] == Verdict(("S", 6, 1), "UNRESOLVED", 1.0, None, "volatile function NOW")


def test_matching_rules():
    n = Cell("S", 1, 1, "x", 1.0, "n")
    assert matches(n, "n", 1.0 + 1e-12, None)
    assert not matches(n, "n", 1.0001, None)
    assert not matches(n, "s", None, "1")
    assert not matches(n, "n", None, None)
    assert matches(Cell("S", 1, 1, "x", "Hi", "s"), "s", None, "Hi")
    assert not matches(Cell("S", 1, 1, "x", "Hi", "s"), "s", None, "hi")
    assert matches(Cell("S", 1, 1, "x", True, "b"), "b", 1.0, None)
    assert matches(Cell("S", 1, 1, "x", "#N/A", "e"), "n", None, None)
    assert not matches(Cell("S", 1, 1, "x", "#N/A", "e"), "n", 0.0, None)
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_verifier.py -v`
Expected: FAIL, `ImportError: cannot import name 'matches'`.

- [ ] **Step 3: Append to `src/unexcel/verifier.py`**

Add `import math` to the imports, and `Cell, Verdict` to the `unexcel.model` import, then append:

```python
def matches(cell: Cell, kind: str, num, txt) -> bool:
    """Excel's saved value vs the SQL result. Plan 1: any Excel error matches any SQL NULL."""
    if cell.kind == "e":
        return num is None and txt is None
    if num is None and txt is None:
        return False
    if cell.kind == "n":
        return kind == "n" and math.isclose(cell.value, num, rel_tol=1e-9, abs_tol=1e-12)
    if cell.kind == "b":
        return kind == "b" and num == (1.0 if cell.value else 0.0)
    if cell.kind == "s":
        return kind == "s" and txt == cell.value
    return False


def verify(book: Book, compiled: Compiled) -> list[Verdict]:
    got = execute(book, compiled)
    out = []
    for _, key, _ in compiled.ctes:
        cell = book.cells[key]
        kind, n, t = got[key]
        actual = t if kind == "s" else (bool(n) if kind == "b" and n is not None else n)
        status = "PASS" if matches(cell, kind, n, t) else "FAIL"
        out.append(Verdict(key, status, cell.value, actual))
    for key, reason in compiled.unresolved.items():
        out.append(Verdict(key, "UNRESOLVED", book.cells[key].value, None, reason))
    return sorted(out, key=lambda v: v.key)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/verifier.py tests/test_verifier.py
git commit -m "feat: per-cell PASS/FAIL/UNRESOLVED verdicts" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 9: dbt project emitter

**Files:**
- Create: `src/unexcel/emitter.py`
- Test: `tests/test_emitter.py`

**Interfaces:**
- Consumes: `render_model`, `input_rows`, `Compiled`, `expr.INPUT`.
- Produces: `emitter.project_name(stem) -> str`, `emitter.dbt_sql(compiled) -> str`, `emitter.write_project(book, compiled, out: Path, stem: str)`. Writes `dbt_project.yml`, `profiles.yml`, `seeds/input_cells.csv`, `models/workbook_cells.sql`, `models/schema.yml`.

- [ ] **Step 1: Write the failing tests `tests/test_emitter.py`** (includes Review Focus 2)

```python
import pytest
from xlsxkit import book_of

from unexcel.compiler import compile_book, render_model
from unexcel.emitter import dbt_sql, project_name, write_project


def test_writes_dbt_project(tmp_path):
    book = book_of({"S": {"A1": 2, "A2": ("=A1*2", 4)}})
    write_project(book, compile_book(book), tmp_path, "My Book 2026")
    assert "name: my_book_2026" in (tmp_path / "dbt_project.yml").read_text(encoding="utf-8")
    assert (tmp_path / "profiles.yml").exists() and (tmp_path / "models" / "schema.yml").exists()
    seed = (tmp_path / "seeds" / "input_cells.csv").read_text(encoding="utf-8").splitlines()
    assert seed == ["sheet,row_num,col_num,kind,num,txt", "S,1,1,n,2.0,"]
    assert "{{ ref('input_cells') }}" in (tmp_path / "models" / "workbook_cells.sql").read_text(encoding="utf-8")


def test_project_name():
    assert project_name("My Book 2026") == "my_book_2026"
    assert project_name("2026 plan!") == "wb_2026_plan"
    assert project_name("***") == "workbook"


def test_dbt_sql_survives_jinja():
    jinja2 = pytest.importorskip("jinja2")
    book = book_of({"S": {"A1": "a", "A2": ('=A1&"{{ oops }} {% x %} {#"', "a{{ oops }} {% x %} {#")}})
    compiled = compile_book(book)
    rendered = jinja2.Template(dbt_sql(compiled)).render(ref=lambda name: name)
    assert rendered == render_model(compiled, "input_cells")
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_emitter.py -v`
Expected: FAIL, `ModuleNotFoundError: No module named 'unexcel.emitter'`.

- [ ] **Step 3: Write `src/unexcel/emitter.py`**

```python
"""Write the compiled workbook as a dbt project (DuckDB profile included)."""
from __future__ import annotations

import csv
import re
from pathlib import Path

from unexcel.compiler import Compiled, input_rows, render_model
from unexcel.expr import INPUT
from unexcel.model import Book

PROJECT = """name: {name}
version: "1.0.0"
config-version: 2
profile: unexcel
seeds:
  {name}:
    +column_types:
      sheet: varchar
      row_num: integer
      col_num: integer
      kind: varchar
      num: double
      txt: varchar
models:
  {name}:
    +materialized: table
"""

PROFILES = """unexcel:
  target: dev
  outputs:
    dev:
      type: duckdb
      path: unexcel.duckdb
"""

SCHEMA = """version: 2
models:
  - name: workbook_cells
    description: "Every formula cell of the workbook, recomputed in SQL. One row per cell."
    columns:
      - name: sheet
        data_tests: [not_null]
      - name: row_num
        data_tests: [not_null]
      - name: col_num
        data_tests: [not_null]
"""


def project_name(stem: str) -> str:
    name = re.sub(r"[^A-Za-z0-9]+", "_", stem).strip("_").lower() or "workbook"
    return f"wb_{name}" if name[0].isdigit() else name


def dbt_sql(compiled: Compiled) -> str:
    """Model SQL for dbt. Braces from workbook text are escaped so Jinja renders them literally."""
    sql = re.sub(r"[{}]", lambda m: "{{ '" + m.group() + "' }}", render_model(compiled, INPUT))
    return sql.replace(INPUT, "{{ ref('input_cells') }}")


def write_project(book: Book, compiled: Compiled, out: Path, stem: str) -> None:
    name = project_name(stem)
    (out / "seeds").mkdir(parents=True, exist_ok=True)
    (out / "models").mkdir(exist_ok=True)
    (out / "dbt_project.yml").write_text(PROJECT.format(name=name), encoding="utf-8")
    (out / "profiles.yml").write_text(PROFILES, encoding="utf-8")
    with open(out / "seeds" / "input_cells.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["sheet", "row_num", "col_num", "kind", "num", "txt"])
        for row in input_rows(book):
            w.writerow(["" if x is None else x for x in row])
    (out / "models" / "workbook_cells.sql").write_text(dbt_sql(compiled) + "\n", encoding="utf-8")
    (out / "models" / "schema.yml").write_text(SCHEMA, encoding="utf-8")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_emitter.py -v`
Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/emitter.py tests/test_emitter.py
git commit -m "feat: emit dbt project with seed, model and DuckDB profile" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 10: Report

**Files:**
- Create: `src/unexcel/report.py`
- Test: `tests/test_report.py`

**Interfaces:**
- Consumes: `Verdict`, `a1`.
- Produces: `report.write_report(verdicts, findings, out: Path, source: str) -> Counter` with keys `PASS FAIL UNRESOLVED` always present. Writes `REPORT.md` and `report.json` (`{"source", "counts", "findings", "cells": [{"cell","status","expected","actual","reason"}]}`). Findings are de-duplicated, keeping order.

- [ ] **Step 1: Write the failing tests `tests/test_report.py`**

```python
import json

from unexcel.model import Verdict
from unexcel.report import write_report


def test_report_counts_sections_and_dedupes_findings(tmp_path):
    verdicts = [
        Verdict(("S", 1, 1), "PASS", 1.0, 1.0),
        Verdict(("S", 2, 1), "FAIL", 5.0, 4.0),
        Verdict(("S", 3, 1), "UNRESOLVED", 1.0, None, "volatile function NOW"),
        Verdict(("S", 4, 1), "FAIL", 2.0, None),
    ]
    counts = write_report(verdicts, ["S!A9: x", "S!A9: x"], tmp_path, "book.xlsx")
    assert counts == {"PASS": 1, "FAIL": 2, "UNRESOLVED": 1}
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "| S!A2 | 5.0 | 4.0 |" in md
    assert "| S!A4 | 2.0 | (error) |" in md
    assert "| S!A3 | volatile function NOW |" in md
    assert md.count("S!A9: x") == 1
    data = json.loads((tmp_path / "report.json").read_text(encoding="utf-8"))
    assert data["counts"]["FAIL"] == 2 and data["cells"][1]["cell"] == "S!A2"


def test_empty_sections_are_left_out(tmp_path):
    write_report([Verdict(("S", 1, 1), "PASS", 1.0, 1.0)], [], tmp_path, "b.xlsx")
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "Failed cells" not in md and "Unresolved" not in md and "Findings" not in md
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_report.py -v`
Expected: FAIL, `ModuleNotFoundError: No module named 'unexcel.report'`.

- [ ] **Step 3: Write `src/unexcel/report.py`**

```python
"""REPORT.md for people, report.json for tools."""
from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

from unexcel.model import Verdict, a1


def _md(value) -> str:
    return "(error)" if value is None else str(value).replace("|", "\\|")


def write_report(verdicts: list[Verdict], findings: list[str], out: Path, source: str) -> Counter:
    counts = Counter({"PASS": 0, "FAIL": 0, "UNRESOLVED": 0})
    counts.update(v.status for v in verdicts)
    findings = list(dict.fromkeys(findings))

    data = {
        "source": source,
        "counts": dict(counts),
        "findings": findings,
        "cells": [{"cell": a1(v.key), "status": v.status, "expected": v.expected, "actual": v.actual,
                   "reason": v.reason} for v in verdicts],
    }
    (out / "report.json").write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    lines = [f"# unexcel report: {source}", "", "| Status | Cells |", "|---|---|"]
    lines += [f"| {status} | {counts[status]} |" for status in ("PASS", "FAIL", "UNRESOLVED")]
    if findings:
        lines += ["", "## Findings in the workbook", ""] + [f"- {f}" for f in findings]
    failed = [v for v in verdicts if v.status == "FAIL"]
    if failed:
        lines += ["", "## Failed cells", "", "| Cell | Excel | SQL |", "|---|---|---|"]
        lines += [f"| {a1(v.key)} | {_md(v.expected)} | {_md(v.actual)} |" for v in failed]
    unresolved = [v for v in verdicts if v.status == "UNRESOLVED"]
    if unresolved:
        lines += ["", "## Unresolved cells", "", "| Cell | Reason |", "|---|---|"]
        lines += [f"| {a1(v.key)} | {_md(v.reason)} |" for v in unresolved]
    (out / "REPORT.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return counts
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_report.py -v`
Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/report.py tests/test_report.py
git commit -m "feat: markdown and JSON verification report" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 11: CLI and end-to-end golden test

**Files:**
- Create: `src/unexcel/cli.py`, `src/unexcel/__main__.py`
- Test: `tests/test_cli.py`

**Interfaces:**
- Consumes: `load`, `UnreadableWorkbook`, `compile_book`, `verify`, `write_project`, `write_report`.
- Produces: `cli.main(argv: list[str] | None = None) -> int`. Command `unexcel convert WORKBOOK [--out DIR]`, default out `<workbook stem>_dbt` next to the workbook. Summary line: `"{p} pass, {f} fail, {u} unresolved -> {path}"`.

- [ ] **Step 1: Write the failing tests `tests/test_cli.py`** (includes Review Focus 1)

```python
import io
import sys

from xlsxkit import SALES, write_xlsx

from unexcel.cli import main

COMMISSION = {
    "A1": "Rep", "B1": "Total", "C1": "Bonus",
    "A2": "Asha", "B2": ("=SUMIFS(Sales!C:C, Sales!A:A, A2)", 120000), "C2": ("=IF(B2>100000, B2*5%, B2*2%)", 6000),
    "A3": "Vikram", "B3": ("=SUMIFS(Sales!C:C, Sales!A:A, A3)", 60000), "C3": 1500,
    "B5": ("=SUM(B2:B3)", 180000), "C5": ("=TODAY()", 46000),
}


def test_convert_commission_end_to_end(tmp_path, capsys):
    wb = tmp_path / "commission.xlsx"
    write_xlsx(wb, {"Sales": SALES, "Commission": COMMISSION})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out")]) == 0
    assert "4 pass, 0 fail, 1 unresolved" in capsys.readouterr().out
    md = (tmp_path / "out" / "REPORT.md").read_text(encoding="utf-8")
    assert "Sales!C5: number stored as text ('15,000')" in md
    assert "| Commission!C5 | volatile function TODAY |" in md
    assert (tmp_path / "out" / "models" / "workbook_cells.sql").exists()


def test_failures_exit_1(tmp_path):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2, "A2": ("=A1*2", 5)}})
    assert main(["convert", str(wb)]) == 1
    assert (tmp_path / "b_dbt" / "REPORT.md").exists()


def test_unreadable_file_exits_2(tmp_path, capsys):
    bad = tmp_path / "x.xls"
    bad.write_bytes(b"x")
    assert main(["convert", str(bad)]) == 2
    assert "save the workbook as .xlsx" in capsys.readouterr().err


def test_console_that_cannot_print_unicode(tmp_path, monkeypatch):
    wb = tmp_path / "Ümsatz✓.xlsx"
    write_xlsx(wb, {"S": {"A1": 1, "A2": ("=A1+1", 2)}})
    monkeypatch.setattr(sys, "stdout", io.TextIOWrapper(io.BytesIO(), encoding="cp1252"))
    assert main(["convert", str(wb)]) == 0
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_cli.py -v`
Expected: FAIL, `ModuleNotFoundError: No module named 'unexcel.cli'`.

- [ ] **Step 3: Write `src/unexcel/cli.py`**

```python
"""unexcel convert WORKBOOK [--out DIR]"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from unexcel.compiler import compile_book
from unexcel.emitter import write_project
from unexcel.reader import UnreadableWorkbook, load
from unexcel.report import write_report
from unexcel.verifier import verify


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")  # Windows consoles can't print every character
    parser = argparse.ArgumentParser(prog="unexcel", description="Turn an Excel workbook into a verified dbt project.")
    sub = parser.add_subparsers(dest="command", required=True)
    convert = sub.add_parser("convert", help="convert a workbook and verify every formula cell")
    convert.add_argument("workbook", type=Path)
    convert.add_argument("--out", type=Path, help="output folder (default: <workbook>_dbt next to the file)")
    args = parser.parse_args(argv)

    try:
        book = load(args.workbook)
    except UnreadableWorkbook as err:
        print(f"error: {err}", file=sys.stderr)
        return 2
    compiled = compile_book(book)
    verdicts = verify(book, compiled)
    out = args.out or args.workbook.with_name(args.workbook.stem + "_dbt")
    write_project(book, compiled, out, args.workbook.stem)
    counts = write_report(verdicts, compiled.findings, out, args.workbook.name)
    print(f"{counts['PASS']} pass, {counts['FAIL']} fail, {counts['UNRESOLVED']} unresolved -> {out / 'REPORT.md'}")
    return 1 if counts["FAIL"] else 0
```

`src/unexcel/__main__.py`:
```python
from unexcel.cli import main

raise SystemExit(main())
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 5: Smoke-test the installed command**

Run: `unexcel --help`
Expected: usage text listing `convert`.

- [ ] **Step 6: Commit**

```bash
git add src/unexcel/cli.py src/unexcel/__main__.py tests/test_cli.py
git commit -m "feat: unexcel convert command with end-to-end commission test" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 12: Optional `dbt build` check + real-Excel ground truth

**Files:**
- Modify: `src/unexcel/emitter.py` (append `run_dbt`), `src/unexcel/cli.py` (add `--dbt`)
- Create: `examples/make_commission.py`, `tests/test_excel_saved.py`
- Test: `tests/test_emitter.py` (append), `tests/test_excel_saved.py`

**Interfaces:**
- Consumes: everything above.
- Produces: `emitter.run_dbt(out: Path) -> tuple[bool, str]`; CLI flag `--dbt` (a failing dbt build makes the exit code 1); `examples/commission.xlsx` (saved by real Excel, committed).

- [ ] **Step 1: Append the failing tests to `tests/test_emitter.py`**

```python
import shutil

from xlsxkit import SALES, write_xlsx

from unexcel.cli import main
from unexcel.emitter import run_dbt


def test_run_dbt_without_dbt_installed(tmp_path, monkeypatch):
    monkeypatch.setattr("unexcel.emitter.shutil.which", lambda name: None)
    assert run_dbt(tmp_path) == (False, "dbt is not installed - run: pip install unexcel[dbt]")


@pytest.mark.skipif(shutil.which("dbt") is None, reason="needs: pip install -e .[dbt]")
def test_dbt_build_passes_on_generated_project(tmp_path):
    wb = tmp_path / "commission.xlsx"
    write_xlsx(wb, {"Sales": SALES, "C": {"A1": "Asha", "B1": ("=SUMIFS(Sales!C:C, Sales!A:A, A1)", 120000)}})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out"), "--dbt"]) == 0
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_emitter.py -v`
Expected: FAIL, `ImportError: cannot import name 'run_dbt'`.

- [ ] **Step 3: Append `run_dbt` to `src/unexcel/emitter.py`**

Add `import shutil` and `import subprocess` to the imports, then append:

```python
def run_dbt(out: Path) -> tuple[bool, str]:
    """End-to-end check: dbt seeds the inputs and builds the model on DuckDB."""
    exe = shutil.which("dbt")
    if exe is None:
        return False, "dbt is not installed - run: pip install unexcel[dbt]"
    proc = subprocess.run([exe, "build", "--project-dir", str(out), "--profiles-dir", str(out)],
                          capture_output=True, text=True, cwd=out)
    return proc.returncode == 0, (proc.stdout + proc.stderr)[-4000:]
```

- [ ] **Step 4: Add `--dbt` to `src/unexcel/cli.py`**

Change the emitter import to `from unexcel.emitter import run_dbt, write_project`. After the `--out` argument add:
```python
    convert.add_argument("--dbt", action="store_true", help="also run `dbt build` on the result")
```
Replace the final `return 1 if counts["FAIL"] else 0` with:
```python
    code = 1 if counts["FAIL"] else 0
    if args.dbt:
        ok, log = run_dbt(out)
        print("dbt build: ok" if ok else f"dbt build failed:\n{log}")
        code = code if ok else 1
    return code
```

- [ ] **Step 5: Run tests**

Run: `python -m pip install -e ".[dev,dbt]"` then `python -m pytest tests -v`
Expected: all passed. The dbt test runs because dbt is now installed.

- [ ] **Step 6: Create `examples/make_commission.py`**

```python
"""Writes examples/commission_unsaved.xlsx (formulas, no saved values).

Open it in real Microsoft Excel, then File > Save As examples/commission.xlsx.
Excel stores its own answers in the file; tests/test_excel_saved.py checks unexcel against them.
"""
from pathlib import Path

import openpyxl

wb = openpyxl.Workbook()
sales = wb.active
sales.title = "Sales"
for row in [["Rep", "Region", "Amount"], ["Asha", "West", 80000], ["asha", "West", 40000],
            ["Vikram", "South", 60000], ["Vikram", "South", "15,000"]]:
    sales.append(row)
c = wb.create_sheet("Commission")
c.append(["Rep", "Total", "Bonus"])
c.append(["Asha", "=SUMIFS(Sales!C:C,Sales!A:A,A2)", "=IF(B2>100000,B2*5%,B2*2%)"])
c.append(["Vikram", "=SUMIFS(Sales!C:C,Sales!A:A,A3)", 1500])
c["B5"] = "=SUM(B2:B3)"
wb.save(Path(__file__).parent / "commission_unsaved.xlsx")
```

- [ ] **Step 7: Write `tests/test_excel_saved.py`**

```python
from pathlib import Path

import pytest

from unexcel.compiler import compile_book
from unexcel.reader import load
from unexcel.verifier import verify

EXAMPLE = Path(__file__).parent.parent / "examples" / "commission.xlsx"


@pytest.mark.skipif(not EXAMPLE.exists(), reason="save examples/commission.xlsx from real Excel (PLAN Task 12)")
def test_real_excel_answers_match():
    book = load(EXAMPLE)
    verdicts = verify(book, compile_book(book))
    assert [v for v in verdicts if v.status != "PASS"] == []
    assert len(verdicts) == 4  # B2, C2, B3, B5
```

- [ ] **Step 8: MANUAL, needs the user's Excel.** Run `python examples/make_commission.py`, open `examples/commission_unsaved.xlsx` in Microsoft Excel, and save it as `examples/commission.xlsx`. Then:

Run: `python -m pytest tests/test_excel_saved.py -v`
Expected: 1 passed. If anything FAILs, real Excel disagrees with the cached values we assumed in the synthetic tests. Fix the compiler, not the expectation.

- [ ] **Step 9: Commit**

```bash
git add src/unexcel/emitter.py src/unexcel/cli.py examples/make_commission.py examples/commission.xlsx tests/test_emitter.py tests/test_excel_saved.py
git commit -m "feat: optional dbt build check and real-Excel ground-truth fixture" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## Self-review (done while writing)

- **Spec coverage, Plan 1 scope:** reader §4 → T1; graph + cycles §4/§6 → T4; compiler §4 → T5–7; verifier + match rule §2.3 → T8; emitter + dbt check §4/§5 → T9, T12; report §4 → T10; CLI + exit codes §5/§6 → T11; error table §6: unsupported format/corrupt → T1, volatile → T5, external links → T3, circular → T4, Excel errors → T5/T8, text numbers → T6/T7, whole-column refs → T3/T7, dates + 1900 bug → T1, no saved value → T5; testing §7: fixtures via `xlsxkit` and real Excel → T12, golden test → T11, rules-only baseline (no AI anywhere) ✓. Deferred items are listed at the top with the plan that owns them.
- **Placeholders:** none. Every code step has full code.
- **Name consistency:** `row_num`/`col_num` everywhere (`row` is a DuckDB keyword); `INPUT = "__unexcel_input__"`; `Compiled.ctes` triples `(name, key, expr)` used the same way in T5, T8 and T9; `note_text_numbers`/`range_numbers`/`arity` defined in T6 and used in T7.
- **Verified against real libraries (2026-10-11, openpyxl 3.1.5, DuckDB 1.5.5):** hand-written xlsx with saved values reads back correctly; tokenizer token types; `range_boundaries("A:A") == (1, None, 1, None)`; `to_excel` handles the 1900 bug; DuckDB `x/0 = inf` (so `NULLIF`), literals are DECIMAL (so `CAST`), `ROUND(2.675::DOUBLE, 2) = 2.68`, `AS MATERIALIZED` CTEs and the SUMIFS/VLOOKUP subquery shapes run.
