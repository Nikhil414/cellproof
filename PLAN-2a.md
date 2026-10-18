# unexcel — Plan 2a: Scale and Excel errors (cell mode) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `unexcel convert` handles a 200k-formula workbook in under 5 minutes, and Excel errors are proven code by code (`#N/A` is no longer the same as `#DIV/0!`). Workbooks with no saved values get an answer key computed by the `formulas` library, labelled `computed`.

**Architecture:** Plan 1 compiled the whole workbook into one SQL statement. That statement breaks DuckDB at about 1,000 formula cells (measured below). Plan 2a keeps cell mode but runs it in **batches**: every known cell (inputs, then the results of earlier batches) lives in one table, and each batch of ≤250 formula cells is one `CREATE TABLE b_k AS SELECT * FROM b_{k-1} UNION ALL <new cells>`. The verifier and the dbt project run the same batch SQL, one dbt model per batch. Each expression carries a second SQL expression for its Excel error code, so errors flow through operators and functions the way Excel's do.

**Setup:** every command assumes the venv is active (`source .venv/bin/activate` on macOS/Linux, `.venv\Scripts\activate` on Windows). Install with `python -m pip install -e ".[dev,dbt]"`; Task 5 adds the `compute` extra.

**Tech Stack:** Python ≥3.10, openpyxl 3.1, DuckDB ≥1.1 (measured on 1.5.6), pytest, dbt-duckdb (optional), formulas ≥1.2 (optional, Task 5).

**Spec:** `DESIGN.md` (accepted 2026-10-11) and the roadmap in `PLAN.md`. This plan builds on the Plan 1 code (branch `plan-1-core`, PR #1).

### Why this plan comes before the rest of Plan 2

The Plan 2 row in `PLAN.md`'s roadmap holds five separate pieces of work. They are split into three plans, each producing working software on its own:

| Plan | Scope |
|---|---|
| **2a (this)** | Batched execution (fixes the crash), cell index, Excel error kinds, computed answer key, 200k perf check |
| 2b | Function library: ~80 functions on top of 2a's error-aware `Expr` |
| 2c | Table mode + region detection: readable `GROUP BY` / join models; the big speed-up |

2a goes first because every function in 2b must produce error codes the 2a way, and because Plan 1 has two crashes that real workbooks will hit.

**Measured 2026-10-11 on the Plan 1 code (DuckDB 1.5.6, Python 3.12, Apple Silicon):**
- 500+ formula cells → `Binder Error: Max expression depth limit of 1000 exceeded`; 1,000 cells with the limit raised → 2.5 s; 4,000 cells → DuckDB segfault (exit 139).
- One formula with 500 `+` terms → uncaught Python `RecursionError` crashes the whole run (DuckDB itself handles 600–900 nesting levels).
- `executemany` input load: 20k rows 6.8 s. CSV + `read_csv`: 201k rows 0.1 s, exact round-trip of commas, quotes, apostrophes and newlines.
- Batched chain, 250 cells per batch: 200k flat cells 71 s; 100k-row running total 50 s (both exact).
- DuckDB `POWER(-8, 1/3) = nan`, `POWER(0, -1) = inf`, `POWER(0, 0) = 1` — no exceptions, so Excel's `#NUM!`/`#DIV/0!` must be computed explicitly.
- A seed CSV text cell `""` is read back as `NULL`.
- `formulas` 1.3.4: correct on the commission example (but 7.5 s, because it expands `C:C` to a million rows); returns `#REF!` for every formula on a sheet whose name contains an apostrophe; 166 MB of dependencies; licence EUPL-1.1+.

### Deliberate limits in Plan 2a (each owned by a later plan)
- Text used as a number (`"3"+1`) stays `UNRESOLVED` (2b).
- `SUMIFS` over ranges or criteria that can hold errors is `UNRESOLVED` (2b).
- Cell mode SQL stays unreadable; 2c replaces it for table-shaped regions.
- dbt gets one model per batch (800 models for 200k cells). 2c reduces this.

## Global Constraints

- Python `>=3.10`. Runtime dependencies: **only** `openpyxl>=3.1`, `duckdb>=1.1`. Optional extras: `dbt` → `dbt-duckdb>=1.8`; `dev` → `pytest>=8`, `jinja2>=3.1`; **new** `compute` → `formulas>=1.2` (optional only, never a core dependency: EUPL licence, 166 MB).
- Supported inputs: `.xlsx`, `.xlsm`. Anything else → `UnreadableWorkbook`, exit code 2.
- Match rule: numbers `math.isclose(rel_tol=1e-9, abs_tol=1e-12)`; text exact (case-sensitive); booleans exact; **an Excel error matches only the same error code** (Plan 1 matched any error to any SQL NULL).
- Never guess: any cell that can't be translated → `UNRESOLVED` with a human-readable reason. Never silently skipped. A computed answer key never overrides a value Excel saved.
- Cell values never leave the process (`formulas` runs locally; no network).
- Exit codes: 0 = no FAIL, 1 = at least one FAIL (or dbt build failed), 2 = unreadable workbook.
- Console output is ASCII-only, and stdout/stderr use `errors="replace"`.
- In SQL every numeric literal is `CAST(x AS DOUBLE)`; every division is `a / NULLIF(b, 0)` plus an explicit `'#DIV/0!'` error; `POWER` errors are explicit (DuckDB returns `nan`/`inf`).
- Cell table schema (seed, verifier and every batch model): `sheet VARCHAR, row_num INTEGER, col_num INTEGER, kind VARCHAR, num DOUBLE, txt VARCHAR`. `kind` is `n`/`s`/`b`/`e`; `num` holds numbers and TRUE/FALSE as 1/0; `txt` holds text and, for kind `e`, the Excel error code.
- The verifier and the dbt project run the same batch SQL.
- Performance target (DESIGN §2.1): 200k formula cells under 5 minutes on a laptop. Measured in Task 7, not promised.

## Review Focus

1. **A sheet formatted down to row 1,048,576** with whole-column references (`=SUM(A:A)`) → converts in seconds. Plan 1 walked every row of the column for every reference. Test in Task 2.
2. **One very long or deeply nested formula** (1,500 `+` terms, 600 nested parentheses) → that cell is `UNRESOLVED` ("formula is nested too deeply to translate") and every other cell still converts. Test in Task 1.
3. **A running-total column thousands of rows long** (`=C1+B2` filled down) → converts and passes; batch boundaries don't break the chain. Test in Task 3.
4. **A `SUM` over cells computed in the same batch** → the batch closes first, so the total includes them. Test in Task 3.
5. **No saved values and `formulas` not installed, or an apostrophe in a sheet name** → `UNRESOLVED` with an actionable reason, never a FAIL against a guessed answer. Test in Task 5.

---
## File structure

```
src/unexcel/
  model.py       + Book.present(): non-empty cells of a rectangle via a per-column index        (Task 2)
  refs.py        + Area.present(book)                                                           (Task 2)
  graph.py       deps via Area.present; RecursionError tolerated                                (Tasks 1, 2)
  expr.py        CELLS placeholder, where_cell/where_area (T3); Expr.err, error(), errs() (T4)
  compiler.py    Step, Compiled.batches, NewBatch, render_batch, write_seed, READ_SEED (T3); errors (T4)
  functions.py   ranges read the cells table (T3); error-aware SUM/IF/... (T4)
  verifier.py    execute runs the batch chain (T3); error-code matching (T4)
  emitter.py     one dbt model per batch + workbook_cells (T3)
  report.py      error codes, "(no value)" (T4); computed answers labelled (T5)
  compute.py     NEW: answer key from the formulas library (T5)
  cli.py         computed answer key wiring (T5)
tests/
  test_limits.py test_index.py test_batches.py test_errors.py test_compute.py test_perf.py   NEW
  test_compiler.py test_functions.py test_lookup.py test_verifier.py test_report.py
  test_emitter.py test_excel_saved.py                                                        MODIFIED
examples/
  make_errors.py   builds the error workbook you save in real Excel (T6)
```

---

### Task 1: A formula nested too deeply is UNRESOLVED, not a crash

**Files:**
- Modify: `src/unexcel/graph.py`, `src/unexcel/compiler.py`
- Test: `tests/test_limits.py`

**Interfaces:**
- Consumes: Plan 1 `compile_book`, `verify`, `formula_order`.
- Produces: `compiler.DEEP = "formula is nested too deeply to translate"`. `formula_order` and `compile_book` never raise `RecursionError`.

- [ ] **Step 1: Write the failing tests `tests/test_limits.py`** (Review Focus 2)

```python
from xlsxkit import book_of

from unexcel.compiler import DEEP, compile_book
from unexcel.verifier import verify


def test_very_long_formula_is_unresolved_not_a_crash():
    terms = "+".join(["A1"] * 1500)
    book = book_of({"S": {"A1": 1, "B1": ("=" + terms, 1500), "C1": ("=A1*2", 2)}})
    compiled = compile_book(book)
    assert compiled.unresolved == {("S", 1, 2): DEEP}
    assert [v.status for v in verify(book, compiled)] == ["UNRESOLVED", "PASS"]  # B1, C1


def test_deeply_nested_parentheses_are_unresolved():
    book = book_of({"S": {"A1": ("=" + "(" * 600 + "1" + ")" * 600, 1)}})
    assert compile_book(book).unresolved == {("S", 1, 1): DEEP}
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_limits.py -v`
Expected: FAIL, `ImportError: cannot import name 'DEEP'`.

- [ ] **Step 3: Tolerate `RecursionError` in `src/unexcel/graph.py`**

Replace
```python
        except Unsupported:
            pass  # the compiler reports the real reason for this cell
```
with
```python
        except (Unsupported, RecursionError):
            pass  # the compiler reports the real reason for this cell
```

- [ ] **Step 4: Catch `RecursionError` in `src/unexcel/compiler.py`**

Below `VOLATILE = ...` add:
```python
DEEP = "formula is nested too deeply to translate"
```
In `compile_book`, replace
```python
        except Unsupported as err:
            unresolved[key] = str(err)
            continue
```
with
```python
        except Unsupported as err:
            unresolved[key] = str(err)
            continue
        except RecursionError:  # Python's stack, not DuckDB's: DuckDB took 900 nesting levels in tests
            unresolved[key] = DEEP
            continue
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 6: Commit**

```bash
git add src/unexcel/graph.py src/unexcel/compiler.py tests/test_limits.py
git commit -m "fix: deeply nested formula is UNRESOLVED instead of crashing the run" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Cell index — ranges only visit cells that exist

**Files:**
- Modify: `src/unexcel/model.py`, `src/unexcel/refs.py`, `src/unexcel/graph.py`, `src/unexcel/functions.py`
- Test: `tests/test_index.py`

**Interfaces:**
- Consumes: Plan 1 `Book`, `Area`.
- Produces: `Book.present(sheet, r1, c1, r2, c2) -> list[Key]` (non-empty cells, row-major); `Area.present(book) -> list[Key]`. Every loop over a range (graph, SUM, SUMIFS, VLOOKUP, text-number findings) uses `present` instead of `Area.cells()`.

- [ ] **Step 1: Write the failing tests `tests/test_index.py`** (Review Focus 1)

```python
import time

from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.refs import Area


def test_present_lists_only_filled_cells_row_major():
    book = book_of({"S": {"B2": 1, "A3": 2, "B1": 3, "C9": 4}})
    assert Area("S", 1, 1, 5, 2).present(book) == [("S", 1, 2), ("S", 2, 2), ("S", 3, 1)]
    assert Area("S", 1, 3, 1_048_576, 3).present(book) == [("S", 9, 3)]
    assert Area("Other", 1, 1, 9, 9).present(book) == []


def test_whole_column_refs_on_a_sheet_formatted_to_the_last_row_stay_fast():
    spec = {f"A{r}": r for r in range(1, 201)}
    spec.update({f"B{r}": ("=SUM(A:A)", 20100) for r in range(1, 201)})
    book = book_of({"S": spec})
    book.max_row["S"] = 1_048_576  # someone formatted the whole sheet, so A:A ends at the last row
    start = time.perf_counter()
    compiled = compile_book(book)
    assert compiled.unresolved == {}
    assert time.perf_counter() - start < 5  # Plan 1 walked 1,048,576 rows per reference: minutes
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_index.py -v`
Expected: FAIL, `AttributeError: 'Area' object has no attribute 'present'`.

- [ ] **Step 3: Add the index to `src/unexcel/model.py`**

Replace the imports at the top with:
```python
from __future__ import annotations

from bisect import bisect_left, bisect_right
from collections import defaultdict
from dataclasses import dataclass, field
from functools import cached_property
from typing import Literal, Union

from openpyxl.utils.cell import get_column_letter
```
Replace the `Book` class with:
```python
@dataclass
class Book:
    cells: dict[Key, Cell]
    sheets: list[str]
    names: dict[str, tuple[str, str]] = field(default_factory=dict)  # NAME (upper) -> (sheet, ref)
    max_row: dict[str, int] = field(default_factory=dict)

    @cached_property
    def _rows_by_col(self) -> dict[tuple[str, int], list[int]]:
        """(sheet, col) -> sorted rows that hold a cell. Built once; build a new Book if cells change."""
        cols = defaultdict(list)
        for sheet, row, col in self.cells:
            cols[(sheet, col)].append(row)
        return {k: sorted(v) for k, v in cols.items()}

    def present(self, sheet: str, r1: int, c1: int, r2: int, c2: int) -> list[Key]:
        """Keys of the non-empty cells inside a rectangle, row-major. Fast on whole-column ranges."""
        found = []
        for col in range(c1, c2 + 1):
            rows = self._rows_by_col.get((sheet, col), [])
            found.extend((sheet, r, col) for r in rows[bisect_left(rows, r1):bisect_right(rows, r2)])
        return sorted(found)
```

- [ ] **Step 4: Add `Area.present` to `src/unexcel/refs.py`**

Inside `class Area`, after `cells()`, add:
```python
    def present(self, book: Book) -> list[Key]:
        """Non-empty cells in the area, row-major. Use this, not cells(), to read a range."""
        return book.present(self.sheet, self.r1, self.c1, self.r2, self.c2)
```

- [ ] **Step 5: Use it in `src/unexcel/graph.py`**

Replace
```python
                found.update(resolve(text, cell.sheet, book).cells())
```
with
```python
                found.update(resolve(text, cell.sheet, book).present(book))
```

- [ ] **Step 6: Use it in `src/unexcel/functions.py`**

In `note_text_numbers` replace the loop header and lookup
```python
    for key in area.cells():
        cell = c.book.cells.get(key)
        if cell is not None and cell.formula is None and cell.kind == "s" and _looks_numeric(cell.value):
```
with
```python
    for key in area.present(c.book):
        cell = c.book.cells[key]
        if cell.formula is None and cell.kind == "s" and _looks_numeric(cell.value):
```
In `range_numbers` replace
```python
    for key in area.cells():
        cell = c.book.cells.get(key)
        if cell is None:
            continue
        if cell.formula is not None:
```
with
```python
    for key in area.present(c.book):
        cell = c.book.cells[key]
        if cell.formula is not None:
```
In `_input_column` replace
```python
    for key in area.cells():
        cell = c.book.cells.get(key)
        if cell is not None and cell.formula is not None:
```
with
```python
    for key in area.present(c.book):
        if c.book.cells[key].formula is not None:
```
In `vlookup` replace
```python
    kinds = [cell.kind for r in range(table.r1, table.r2 + 1)
             if (cell := c.book.cells.get((table.sheet, r, col))) is not None and cell.kind != "blank"]
```
with
```python
    kinds = [c.book.cells[k].kind for k in Area(table.sheet, table.r1, col, table.r2, col).present(c.book)
             if c.book.cells[k].kind != "blank"]
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 8: Commit**

```bash
git add src/unexcel/model.py src/unexcel/refs.py src/unexcel/graph.py src/unexcel/functions.py tests/test_index.py
git commit -m "perf: per-column cell index so ranges only visit cells that exist" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Batched execution over one cells table (verifier and dbt run the same SQL)

**Files:**
- Modify: `src/unexcel/expr.py` (rename the placeholder, add `where_cell`/`where_area`)
- Replace: `src/unexcel/compiler.py`, `src/unexcel/functions.py`, `src/unexcel/verifier.py`, `src/unexcel/emitter.py`
- Test: `tests/test_batches.py` (new), `tests/test_emitter.py` (modify)

**Interfaces:**
- Consumes: `Area.present`, `formula_order`, `parse`, `resolve`, `DEEP`.
- Produces:
  - `expr.CELLS = "__unexcel_cells__"` (replaces `expr.INPUT`): placeholder for "the table of every cell known before this batch". `expr.where_cell(key) -> str`, `expr.where_area(area, alias="") -> str`.
  - `compiler.Step(key, expr, cte)`; `compiler.Compiled(batches: list[list[Step]], unresolved, findings)` with property `.steps`; `compiler.NewBatch`; `compiler.BATCH = 250`; `compiler.NO_VALUE`; `compiler.render_batch(steps, prev) -> str`; `compiler.write_seed(book, path)`; `compiler.READ_SEED` (DuckDB `read_csv(?, ...)` for the seed); `compiler.SEED_HEADER`.
  - `Compiler.ready(area) -> list[Key]` (non-empty cells of a range; raises `NewBatch` if one is in the open batch, `Unsupported` if one is unresolved); `Compiler.kind_of(key) -> str` (`n`/`s`/`b`).
  - `Compiler.done: dict[Key, str]` (type of each formula cell in a closed batch); `Compiler.open: dict[Key, tuple[str, str]]` (cte, type).
  - `verifier.execute(book, compiled)` runs `b0` (seed) → `b1` → … and returns `{key: (kind, num, txt)}` for formula cells.
  - `emitter.model_name(k)`, `emitter.dbt_sql(steps, prev_model)`, `emitter.final_sql(last_model)`. Project layout: `seeds/input_cells.csv`, `models/unexcel_cells/cells_0001.sql …`, `models/workbook_cells.sql`.
- Plan 1 error behaviour is kept in this task (a formula error is a NULL number); Task 4 replaces it.

- [ ] **Step 1: Write the failing tests `tests/test_batches.py`** (Review Focus 3 and 4)

```python
from xlsxkit import book_of

import unexcel.compiler as compiler
from unexcel.compiler import compile_book
from unexcel.verifier import execute, verify


def run(sheets):
    book = book_of(sheets)
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def test_thousands_of_cells_no_longer_crash():  # Plan 1: DuckDB depth error at ~1k cells, segfault at 4k
    n = 3000
    spec = {f"A{r}": float(r) for r in range(1, n + 1)}
    spec.update({f"B{r}": (f"=A{r}*2", 2.0 * r) for r in range(1, n + 1)})
    book = book_of({"S": spec})
    compiled = compile_book(book)
    assert len(compiled.batches) == n // compiler.BATCH
    assert {v.status for v in verify(book, compiled)} == {"PASS"}


def test_running_total_thousands_of_rows_deep():  # Review Focus 3
    n = 3000
    spec = {f"A{r}": 1 for r in range(1, n + 1)}
    spec["B1"] = ("=A1", 1)
    spec.update({f"B{r}": (f"=B{r - 1}+A{r}", float(r)) for r in range(2, n + 1)})
    book = book_of({"S": spec})
    verdicts = verify(book, compile_book(book))
    assert len(verdicts) == n and {v.status for v in verdicts} == {"PASS"}


def test_batch_size_limit(monkeypatch):
    monkeypatch.setattr(compiler, "BATCH", 3)
    book = book_of({"S": {f"A{r}": (f"={r}", float(r)) for r in range(1, 8)}})
    assert [len(b) for b in compile_book(book).batches] == [3, 3, 1]


def test_range_over_the_open_batch_closes_it_first():  # Review Focus 4
    book = book_of({"S": {"A1": 1, "B1": ("=A1+1", 2), "B2": ("=A1+2", 3),
                          "C1": ("=SUM(B1:B2)", 5), "C2": ("=C1*2", 10)}})
    compiled = compile_book(book)
    assert [[s.key for s in b] for b in compiled.batches] == [
        [("S", 1, 2), ("S", 2, 2)],  # B1, B2
        [("S", 1, 3), ("S", 2, 3)],  # C1 (range over B1:B2 -> new batch), C2 (single ref -> same batch)
    ]
    assert {v.status for v in verify(book, compiled)} == {"PASS"}


def test_formula_results_feed_sumifs_and_vlookup():  # Plan 1 rejected ranges holding formula cells
    got, c = run({"S": {
        "A1": "x", "A2": "y", "A3": "x",
        "B1": 10, "B2": 20, "B3": 30,
        "C1": ("=B1*2", 20), "C2": ("=B2*2", 40), "C3": ("=B3*2", 60),
        "D1": ('=SUMIFS(C1:C3, A1:A3, "x")', 80),
        "D2": ('=VLOOKUP("y", A1:C3, 3, FALSE)', 40),
    }})
    assert c.unresolved == {}
    assert got[("S", 1, 4)] == ("n", 80.0, None)
    assert got[("S", 2, 4)] == ("n", 40.0, None)


def test_empty_text_input_stays_empty_text():  # the seed CSV reads "" back as NULL
    got, _ = run({"S": {"A1": "", "A2": ('=A1&"x"', "x")}})
    assert got[("S", 2, 1)] == ("s", None, "x")
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_batches.py -v`
Expected: FAIL. `AttributeError: 'Compiled' object has no attribute 'batches'` (and the 3,000-cell tests hit the Plan 1 DuckDB depth error).

- [ ] **Step 3: Edit `src/unexcel/expr.py`**

Replace
```python
from unexcel.model import Unsupported

INPUT = "__unexcel_input__"  # replaced by the real input table when SQL is rendered
```
with
```python
from unexcel.model import Key, Unsupported

CELLS = "__unexcel_cells__"  # every cell known before the current batch; replaced when SQL is rendered
```
Append at the end of the file:
```python
def where_cell(key: Key) -> str:
    sheet, row, col = key
    return f"sheet = {lit_str(sheet).sql} AND row_num = {row} AND col_num = {col}"


def where_area(area, alias: str = "") -> str:
    p = f"{alias}." if alias else ""
    return (f"{p}sheet = {lit_str(area.sheet).sql} AND {p}row_num BETWEEN {area.r1} AND {area.r2} "
            f"AND {p}col_num BETWEEN {area.c1} AND {area.c2}")
```

- [ ] **Step 4: Replace `src/unexcel/compiler.py`**

```python
"""Formula AST -> SQL. Cell mode, run in batches.

Every cell known so far (inputs, then the results of earlier batches) lives in one table with the
columns of INPUT_COLUMNS. A batch is up to BATCH consecutive formula cells in dependency order. It
becomes one SQL statement: the cells it computes are materialized CTEs, appended to the table it
read from. One statement per workbook broke DuckDB at ~1k formula cells; batches keep every
statement small.
"""
from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from unexcel.expr import BLANK, CELLS, Expr, finalize, lit_num, lit_str, num, txt, unify, where_cell
from unexcel.functions import FUNCS
from unexcel.graph import formula_order
from unexcel.model import Book, Cell, Key, Unsupported, a1
from unexcel.parser import Binary, Bool, Call, ErrLit, Num, Postfix, RefNode, Str, Unary, parse
from unexcel.refs import Area, resolve

INPUT_COLUMNS = "sheet VARCHAR, row_num INTEGER, col_num INTEGER, kind VARCHAR, num DOUBLE, txt VARCHAR"
SEED_HEADER = ["sheet", "row_num", "col_num", "kind", "num", "txt"]
READ_SEED = ("read_csv(?, header = true, quote = '\"', escape = '\"', columns = {'sheet': 'VARCHAR', "
             "'row_num': 'INTEGER', 'col_num': 'INTEGER', 'kind': 'VARCHAR', 'num': 'DOUBLE', 'txt': 'VARCHAR'})")
VOLATILE = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT"}
ARITH = {"+", "-", "*"}
KIND = {"num": "n", "txt": "s", "bool": "b"}
TYPE = {"n": "num", "s": "txt", "b": "bool"}
NO_VALUE = "no saved value in the file - open and save it in Excel"
DEEP = "formula is nested too deeply to translate"
# ponytail: fixed batch size. Measured on DuckDB 1.5.6: 200k flat cells 71 s, a 100k-row running total 50 s.
# Table mode (Plan 2c) is the real speed-up; change this only with a new measurement.
BATCH = 250


@dataclass(frozen=True)
class Step:
    key: Key
    expr: Expr
    cte: str  # its CTE name inside its batch


@dataclass
class Compiled:
    batches: list[list[Step]]
    unresolved: dict[Key, str]
    findings: list[str]

    @property
    def steps(self) -> list[Step]:
        return [s for batch in self.batches for s in batch]


class NewBatch(Exception):
    """A range reads a cell of the open batch. Ranges read the table, so the batch must close first."""


class Compiler:
    def __init__(self, book: Book):
        self.book = book
        self.done: dict[Key, str] = {}  # formula cell in a closed batch -> type
        self.open: dict[Key, tuple[str, str]] = {}  # formula cell in the open batch -> (cte, type)
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
        if key in self.open:
            cte, typ = self.open[key]
            return Expr(f"(SELECT v FROM {cte})", typ)
        if cell.formula is None:
            if cell.kind == "e":
                raise Unsupported(f"input cell {a1(key)} holds an error")
            typ = TYPE[cell.kind]
        elif key in self.done:
            typ = self.done[key]
        else:
            raise Unsupported(f"depends on unresolved cell {a1(key)}")
        where = f"FROM {CELLS} WHERE {where_cell(key)}"
        if typ == "num":
            return Expr(f"(SELECT num {where})", "num")
        if typ == "txt":
            return Expr(f"(SELECT COALESCE(txt, '') {where})", "txt")  # the seed CSV reads '' back as NULL
        return Expr(f"((SELECT num {where}) <> 0)", "bool")

    def ready(self, area: Area) -> list[Key]:
        """The non-empty cells of a range that SQL will read from the table. All must be computed already."""
        keys = area.present(self.book)
        for key in keys:
            if self.book.cells[key].formula is None:
                continue
            if key in self.open:
                raise NewBatch
            if key not in self.done:
                raise Unsupported(f"depends on unresolved cell {a1(key)}")
        return keys

    def kind_of(self, key: Key) -> str:
        """Value kind (n/s/b/e) a cell has in the table. Call ready() on its range first."""
        cell = self.book.cells[key]
        return KIND[self.done[key]] if cell.formula is not None else cell.kind


def _close(comp: Compiler, batches: list[list[Step]], current: list[Step]) -> None:
    """Seal the open batch: later cells read its results from the table."""
    if current:
        batches.append(list(current))
        comp.done.update((s.key, s.expr.type) for s in current)
        comp.open.clear()
        current.clear()


def compile_book(book: Book) -> Compiled:
    order, unresolved = formula_order(book)
    comp = Compiler(book)
    batches: list[list[Step]] = []
    current: list[Step] = []
    for key in order:
        cell = book.cells[key]
        if cell.kind == "blank":
            unresolved[key] = NO_VALUE
            continue
        if len(current) >= BATCH:
            _close(comp, batches, current)
        try:
            try:
                e = comp.compile_cell(cell)
            except NewBatch:
                _close(comp, batches, current)
                e = comp.compile_cell(cell)  # the open batch is empty now, so this can't raise NewBatch
        except Unsupported as err:
            unresolved[key] = str(err)
            continue
        except RecursionError:  # Python's stack, not DuckDB's: DuckDB took 900 nesting levels in tests
            unresolved[key] = DEEP
            continue
        step = Step(key, e, f"c{len(current) + 1}")
        comp.open[key] = (step.cte, e.type)
        current.append(step)
    _close(comp, batches, current)
    return Compiled(batches, unresolved, comp.findings)


def render_batch(steps: list[Step], prev: str) -> str:
    """One batch as a SELECT: every row of `prev` plus one row per cell this batch computes."""
    ctes = ",\n".join(f"{s.cte} AS MATERIALIZED (SELECT {s.expr.sql} AS v)" for s in steps)
    rows = []
    for s in steps:
        sheet, row, col = s.key
        v = f"(SELECT v FROM {s.cte})"
        num_sql = "CAST(NULL AS DOUBLE)" if s.expr.type == "txt" else f"CAST({v} AS DOUBLE)"
        txt_sql = f"CAST({v} AS VARCHAR)" if s.expr.type == "txt" else "CAST(NULL AS VARCHAR)"
        rows.append(f"SELECT {lit_str(sheet).sql}, {row}, {col}, '{KIND[s.expr.type]}', {num_sql}, {txt_sql}")
    sql = f"WITH\n{ctes}\nSELECT * FROM {CELLS}\nUNION ALL\n" + "\nUNION ALL\n".join(rows)
    return sql.replace(CELLS, prev)


def input_rows(book: Book) -> list[tuple]:
    """Every non-formula cell as a row of the cells table (see INPUT_COLUMNS)."""
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


def write_seed(book: Book, path: Path) -> None:
    """The input cells as CSV: the dbt seed, and what the verifier loads (0.1 s for 200k rows)."""
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(SEED_HEADER)
        for row in input_rows(book):
            w.writerow(["" if x is None else x for x in row])
```

- [ ] **Step 5: Replace `src/unexcel/functions.py`**

```python
"""Excel function -> SQL. Each entry: fn(compiler, args, sheet) -> Expr.

Ranges are read from the cells table (inputs plus the formula cells of earlier batches).
Compiler.ready() checks every formula cell in a range was computed in an earlier batch.
"""
from __future__ import annotations

from unexcel.expr import CELLS, ZERO, Expr, as_bool, lit_str, num, unify, where_area
from unexcel.model import Unsupported, a1
from unexcel.parser import Bool, Num, RefNode, Str
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


def note_text_numbers(c, keys):
    """Excel skips text in SUM-like ranges, even text that looks like a number. Tell the user."""
    for key in keys:
        cell = c.book.cells[key]
        if cell.formula is None and cell.kind == "s" and _looks_numeric(cell.value):
            c.findings.append(f"{a1(key)}: number stored as text ({cell.value!r}) is skipped by Excel")


def range_sum(c, area: Area) -> Expr:
    """SUM over a range: text, TRUE/FALSE and blanks are skipped, like Excel."""
    keys = c.ready(area)
    for key in keys:
        cell = c.book.cells[key]
        if cell.formula is None and cell.kind == "e":
            raise Unsupported(f"input cell {a1(key)} holds an error")
    note_text_numbers(c, keys)
    # Until Task 4 a formula error is a NULL number, and it makes the whole sum an error.
    return Expr(f"(SELECT CASE WHEN COUNT(*) FILTER (WHERE kind = 'n' AND num IS NULL) > 0 THEN NULL "
                f"ELSE COALESCE(SUM(num) FILTER (WHERE kind = 'n'), CAST(0 AS DOUBLE)) END "
                f"FROM {CELLS} WHERE {where_area(area)})", "num")


@register("SUM")
def sum_(c, args, sheet):
    parts = []
    for arg in args:
        if isinstance(arg, RefNode):
            parts.append(range_sum(c, resolve(arg.text, sheet, c.book)).sql)
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
    # DuckDB ROUND on DOUBLE rounds half away from zero like Excel (2.675 -> 2.68, verified on 1.5.6)
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
    return Expr(f"COALESCE({value.sql}, {fallback.sql})", value.type)  # until Task 4: SQL NULL = Excel error


def _column(c, node, sheet, fname) -> tuple[Area, list]:
    if not isinstance(node, RefNode):
        raise Unsupported(f"{fname} needs a cell range")
    area = resolve(node.text, sheet, c.book)
    return area, c.ready(area)


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
    total, total_keys = _column(c, args[0], sheet, "SUMIFS")
    if total.c1 != total.c2:
        raise Unsupported("SUMIFS sum range wider than one column")
    joins, conds = [], []
    for i in range(1, len(args), 2):
        area, _ = _column(c, args[i], sheet, "SUMIFS")
        if area.c1 != area.c2 or area.r2 - area.r1 != total.r2 - total.r1:
            raise Unsupported("SUMIFS ranges must be single columns of the same height")
        node = args[i + 1]
        if isinstance(node, Str) and node.value and (node.value[0] in "<>=" or any(ch in node.value for ch in "*?~")):
            raise Unsupported("SUMIFS comparison/wildcard criteria not supported yet")
        k = f"k{len(joins) + 1}"
        joins.append(f"JOIN {CELLS} {k} ON {k}.sheet = {lit_str(area.sheet).sql} AND {k}.col_num = {area.c1} "
                     f"AND {k}.row_num = s.row_num + {area.r1 - total.r1}")
        conds.append(_match(k, c.expr(node, sheet)))
    note_text_numbers(c, total_keys)
    where = " AND ".join([where_area(total, "s"), "s.kind = 'n'", *conds])
    return Expr(f"(SELECT COALESCE(SUM(s.num), CAST(0 AS DOUBLE)) FROM {CELLS} s {' '.join(joins)} WHERE {where})",
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
    table, keys = _column(c, args[1], sheet, "VLOOKUP")
    col = table.c1 + int(args[2].value) - 1
    if not table.c1 <= col <= table.c2:
        raise Unsupported("VLOOKUP column index outside the table")
    kinds = [c.kind_of(k) for k in keys if k[2] == col]
    if len(kinds) > 1 and kinds[0] == "s" and set(kinds[1:]) == {"n"}:
        kinds = kinds[1:]  # text header above a numeric column
    if set(kinds) == {"n"}:
        column, typ, empty = "num", "num", "CAST(0 AS DOUBLE)"
    elif set(kinds) == {"s"}:
        column, typ, empty = "txt", "txt", "''"
    else:
        raise Unsupported("VLOOKUP result column mixes types")
    lookup = c.expr(args[0], sheet)
    first_col = Area(table.sheet, table.r1, table.c1, table.r2, table.c1)
    return Expr(
        f"(SELECT COALESCE(r.{column}, {empty}) FROM {CELLS} k LEFT JOIN {CELLS} r "
        f"ON r.sheet = k.sheet AND r.row_num = k.row_num AND r.col_num = {col} "
        f"WHERE {where_area(first_col, 'k')} AND {_match('k', lookup)} "
        f"ORDER BY k.row_num LIMIT 1)",
        typ,
    )
```

- [ ] **Step 6: Replace `src/unexcel/verifier.py`**

```python
"""Run the compiled batches on DuckDB and compare every cell with Excel's saved value."""
from __future__ import annotations

import math
import tempfile
from pathlib import Path

import duckdb

from unexcel.compiler import READ_SEED, Compiled, render_batch, write_seed
from unexcel.model import Book, Cell, Key, Verdict


def execute(book: Book, compiled: Compiled) -> dict[Key, tuple]:
    """The same statements as the dbt project: seed -> b1 -> b2 -> ... Returns formula cells only."""
    con = duckdb.connect()
    with tempfile.TemporaryDirectory() as tmp:
        seed = Path(tmp) / "input_cells.csv"
        write_seed(book, seed)
        con.execute(f"CREATE TABLE b0 AS SELECT * FROM {READ_SEED}", [str(seed)])
    for k, steps in enumerate(compiled.batches, 1):
        con.execute(f"CREATE TABLE b{k} AS {render_batch(steps, f'b{k - 1}')}")
        con.execute(f"DROP TABLE b{k - 1}")
    wanted = {s.key for s in compiled.steps}
    rows = con.execute(f"SELECT * FROM b{len(compiled.batches)}").fetchall()
    return {(s, r, c): (k, n, t) for s, r, c, k, n, t in rows if (s, r, c) in wanted}


def matches(cell: Cell, kind: str, num, txt) -> bool:
    """Excel's saved value vs the SQL result. Until Task 4: any Excel error matches any SQL NULL."""
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
    for step in compiled.steps:
        cell = book.cells[step.key]
        kind, n, t = got[step.key]
        actual = t if kind == "s" else (bool(n) if kind == "b" and n is not None else n)
        out.append(Verdict(step.key, "PASS" if matches(cell, kind, n, t) else "FAIL", cell.value, actual))
    for key, reason in compiled.unresolved.items():
        out.append(Verdict(key, "UNRESOLVED", book.cells[key].value, None, reason))
    return sorted(out, key=lambda v: v.key)
```

- [ ] **Step 7: Replace `src/unexcel/emitter.py`**

```python
"""Write the compiled workbook as a dbt project (DuckDB profile included).

The models run exactly the SQL the verifier ran: seed input_cells -> unexcel_cells/cells_0001 -> ...
-> workbook_cells, which keeps the formula cells only.
"""
from __future__ import annotations

import re
import shutil
import subprocess
from pathlib import Path

from unexcel.compiler import Compiled, Step, render_batch, write_seed
from unexcel.expr import CELLS
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


def model_name(k: int) -> str:
    return f"cells_{k:04d}"


def dbt_sql(steps: list[Step], prev: str) -> str:
    """One batch as a dbt model. Braces from workbook text are escaped so Jinja renders them literally."""
    sql = re.sub(r"[{}]", lambda m: "{{ '" + m.group() + "' }}", render_batch(steps, CELLS))
    return sql.replace(CELLS, "{{ ref('" + prev + "') }}")


def final_sql(last: str) -> str:
    """The formula cells: every row of the last batch that is not an input."""
    return ("SELECT c.* FROM {{ ref('" + last + "') }} AS c\nWHERE NOT EXISTS (\n"
            "  SELECT 1 FROM {{ ref('input_cells') }} AS i\n"
            "  WHERE i.sheet = c.sheet AND i.row_num = c.row_num AND i.col_num = c.col_num\n)")


def write_project(book: Book, compiled: Compiled, out: Path, stem: str) -> None:
    name = project_name(stem)
    batches_dir = out / "models" / "unexcel_cells"
    (out / "seeds").mkdir(parents=True, exist_ok=True)
    batches_dir.mkdir(parents=True, exist_ok=True)
    for old in batches_dir.glob("cells_*.sql"):  # a rerun must not leave batches from a bigger workbook
        old.unlink()
    (out / "dbt_project.yml").write_text(PROJECT.format(name=name), encoding="utf-8")
    (out / "profiles.yml").write_text(PROFILES, encoding="utf-8")
    write_seed(book, out / "seeds" / "input_cells.csv")
    prev = "input_cells"
    for k, steps in enumerate(compiled.batches, 1):
        (batches_dir / f"{model_name(k)}.sql").write_text(dbt_sql(steps, prev) + "\n", encoding="utf-8")
        prev = model_name(k)
    (out / "models" / "workbook_cells.sql").write_text(final_sql(prev) + "\n", encoding="utf-8")
    (out / "models" / "schema.yml").write_text(SCHEMA, encoding="utf-8")


def run_dbt(out: Path) -> tuple[bool, str]:
    """End-to-end check: dbt seeds the inputs and builds the models on DuckDB."""
    exe = shutil.which("dbt")
    if exe is None:
        return False, "dbt is not installed - run: pip install unexcel[dbt]"
    proc = subprocess.run([exe, "build", "--project-dir", str(out), "--profiles-dir", str(out)],
                          capture_output=True, text=True, cwd=out)
    return proc.returncode == 0, (proc.stdout + proc.stderr)[-4000:]
```

- [ ] **Step 8: Update `tests/test_emitter.py`**

Replace the import line `from unexcel.compiler import compile_book, render_model` with `from unexcel.compiler import compile_book, render_batch`, and add `import duckdb` below `import shutil`. Replace `test_writes_dbt_project` and `test_dbt_sql_survives_jinja` with:
```python
def test_writes_dbt_project(tmp_path):
    book = book_of({"S": {"A1": 2, "A2": ("=A1*2", 4)}})
    write_project(book, compile_book(book), tmp_path, "My Book 2026")
    assert "name: my_book_2026" in (tmp_path / "dbt_project.yml").read_text(encoding="utf-8")
    assert (tmp_path / "profiles.yml").exists() and (tmp_path / "models" / "schema.yml").exists()
    seed = (tmp_path / "seeds" / "input_cells.csv").read_text(encoding="utf-8").splitlines()
    assert seed == ["sheet,row_num,col_num,kind,num,txt", "S,1,1,n,2.0,"]
    batch = (tmp_path / "models" / "unexcel_cells" / "cells_0001.sql").read_text(encoding="utf-8")
    assert "{{ ref('input_cells') }}" in batch
    assert "{{ ref('cells_0001') }}" in (tmp_path / "models" / "workbook_cells.sql").read_text(encoding="utf-8")


def test_rerun_removes_stale_batches(tmp_path):
    two = book_of({"S": {"A1": 1, "A2": ("=A1", 1), "A3": ("=SUM(A2:A2)", 1)}})  # the range closes batch 1
    write_project(two, compile_book(two), tmp_path, "b")
    one = book_of({"S": {"A1": 1, "A2": ("=A1", 1)}})
    write_project(one, compile_book(one), tmp_path, "b")
    assert sorted(p.name for p in (tmp_path / "models" / "unexcel_cells").iterdir()) == ["cells_0001.sql"]


def test_dbt_sql_survives_jinja():
    jinja2 = pytest.importorskip("jinja2")
    book = book_of({"S": {"A1": "a", "A2": ('=A1&"{{ oops }} {% x %} {#"', "a{{ oops }} {% x %} {#")}})
    steps = compile_book(book).batches[0]
    rendered = jinja2.Template(dbt_sql(steps, "input_cells")).render(ref=lambda name: name)
    assert rendered == render_batch(steps, "input_cells")
```
Replace the body of `test_dbt_build_passes_on_generated_project` so it also checks what dbt built:
```python
    wb = tmp_path / "commission.xlsx"
    write_xlsx(wb, {"Sales": SALES, "C": {"A1": "Asha", "B1": ("=SUMIFS(Sales!C:C, Sales!A:A, A1)", 120000)}})
    out = tmp_path / "out"
    assert main(["convert", str(wb), "--out", str(out), "--dbt"]) == 0
    con = duckdb.connect(str(out / "unexcel.duckdb"), read_only=True)
    assert con.execute("SELECT sheet, row_num, col_num, kind, num FROM workbook_cells").fetchall() == [
        ("C", 1, 2, "n", 120000.0)]
```

- [ ] **Step 9: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed (the dbt build test runs when `dbt` is on PATH, i.e. with the venv active).

- [ ] **Step 10: Commit**

```bash
git add src/unexcel/expr.py src/unexcel/compiler.py src/unexcel/functions.py src/unexcel/verifier.py src/unexcel/emitter.py tests/test_batches.py tests/test_emitter.py
git commit -m "feat: run cell mode in batches over one cells table; dbt gets one model per batch" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: Excel error kinds end to end

**Files:**
- Replace: `src/unexcel/expr.py`, `src/unexcel/compiler.py`, `src/unexcel/functions.py`, `src/unexcel/verifier.py`
- Modify: `src/unexcel/report.py`
- Test: `tests/test_errors.py` (new); `tests/test_compiler.py`, `tests/test_functions.py`, `tests/test_lookup.py`, `tests/test_verifier.py`, `tests/test_report.py` (one assertion each)

**Interfaces:**
- Consumes: everything from Task 3.
- Produces:
  - `expr.Expr(sql, type, err=NO_ERR)`: `err` is SQL giving an Excel error code or NULL. `expr.NO_ERR = "NULL"` means "cannot fail". New type `"err"` (a value that is always an error). `expr.error(code) -> Expr`, `expr.errs(*err_sqls) -> str` (first error, left to right).
  - `num`/`txt`/`as_bool`/`unify` carry `err` through; `"err"` coerces to any type.
  - `Compiler.done: dict[Key, tuple[str, bool]]` (type, may be an error); `Compiler.open: dict[Key, tuple[str, str, bool]]` (cte, type, may be an error); `Compiler.may_err(key) -> bool`. `KIND["err"] = "e"`.
  - Batch CTEs have columns `v` (value) and `e` (error code or NULL). A cell whose `e` is set is stored with kind `e` and the code in `txt`.
  - `execute` returns `("e", None, "#DIV/0!")` for errors. `matches` requires the same error code.
  - `functions.first_error(c, area, keys) -> str`, `functions.range_sum(c, area) -> Expr` (with `err`).
- Excel rules implemented: first error left to right wins; `x/0` → `#DIV/0!`; `0^negative` → `#DIV/0!`; `0^0` and `negative^fraction` → `#NUM!`; error literals (`=#N/A`) and input error cells flow through; `IF` returns the taken branch's error; `IFERROR` catches; `AND`/`OR`/`ROUND`/`SUM` propagate; `SUM` over a range returns the first error row by row (confirmed against real Excel in Task 6); `VLOOKUP` returns `#N/A` for a missing key and the found cell's own error otherwise.

- [ ] **Step 1: Write the failing tests `tests/test_errors.py`**

```python
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.model import Cell
from unexcel.verifier import execute, matches


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
        "B4": ("=A2^(1/3)", "#NUM!"),
        "B5": ("=#N/A", "#N/A"),
        "B6": ("=B5+1", "#N/A"),
        "B7": ("=B1+B5", "#DIV/0!"),  # the first error, left to right
        "B8": ('=B5&"x"', "#N/A"),
        "B9": ("=-B1", "#DIV/0!"),
    }})
    assert c.unresolved == {}
    assert [got[("S", r, 2)] for r in range(1, 10)] == [
        E("#DIV/0!"), E("#DIV/0!"), E("#NUM!"), E("#NUM!"), E("#N/A"), E("#N/A"), E("#DIV/0!"), E("#N/A"),
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


def test_sumifs_over_error_cells_is_unresolved():
    _, c = run({"S": {"A1": "x", "B1": "#N/A", "C1": ('=SUMIFS(B1:B1, A1:A1, "x")', "#N/A")}})
    assert c.unresolved[("S", 1, 3)] == "SUMIFS over cells that hold errors not supported yet"


def test_error_codes_must_match_exactly():
    cell = Cell("S", 1, 1, "x", "#N/A", "e")
    assert matches(cell, "e", None, "#N/A")
    assert not matches(cell, "e", None, "#DIV/0!")
    assert not matches(cell, "n", None, None)
    assert not matches(Cell("S", 1, 1, "x", 1.0, "n"), "e", None, "#N/A")
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_errors.py -v`
Expected: FAIL. Errors come back as `("n", None, None)`; `=#N/A` is UNRESOLVED ("error literal #N/A"); `matches(cell, "e", None, "#N/A")` is False.

- [ ] **Step 3: Replace `src/unexcel/expr.py`**

```python
"""Typed SQL expressions, Excel-style type coercion and Excel error codes."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from unexcel.model import Key, Unsupported

CELLS = "__unexcel_cells__"  # every cell known before the current batch; replaced when SQL is rendered
NO_ERR = "NULL"  # the err of an expression that cannot fail
Type = Literal["num", "txt", "bool", "blank", "err"]


@dataclass(frozen=True)
class Expr:
    sql: str  # the value; meaningless when err is set
    type: Type
    err: str = NO_ERR  # SQL giving the Excel error code ('#DIV/0!') or NULL


ZERO = Expr("CAST(0 AS DOUBLE)", "num")
BLANK = Expr("NULL", "blank")  # an empty cell; always coerced before use, never rendered raw


def lit_num(v: float) -> Expr:
    return Expr(f"CAST({v!r} AS DOUBLE)", "num")  # bare literals would be DECIMAL in DuckDB


def lit_str(v: str) -> Expr:
    return Expr("'" + v.replace("'", "''") + "'", "txt")


def error(code: str) -> Expr:
    """A value that is always this Excel error."""
    return Expr("NULL", "err", lit_str(code).sql)


def errs(*parts: str) -> str:
    """Excel reports the first error it meets, left to right."""
    live = [p for p in parts if p != NO_ERR]
    if not live:
        return NO_ERR
    return live[0] if len(live) == 1 else f"COALESCE({', '.join(live)})"


def num(e: Expr) -> Expr:
    if e.type == "num":
        return e
    if e.type == "blank":
        return ZERO
    if e.type == "bool":
        return Expr(f"CAST({e.sql} AS DOUBLE)", "num", e.err)
    if e.type == "err":
        return Expr("CAST(NULL AS DOUBLE)", "num", e.err)
    raise Unsupported("text used as a number")


def txt(e: Expr) -> Expr:
    if e.type == "txt":
        return e
    if e.type == "blank":
        return Expr("''", "txt")
    if e.type == "err":
        return Expr("CAST(NULL AS VARCHAR)", "txt", e.err)
    raise Unsupported("number used as text")


def as_bool(e: Expr) -> Expr:
    if e.type == "bool":
        return e
    if e.type == "blank":
        return Expr("FALSE", "bool")
    if e.type == "num":
        return Expr(f"({e.sql} <> 0)", "bool", e.err)
    if e.type == "err":
        return Expr("CAST(NULL AS BOOLEAN)", "bool", e.err)
    raise Unsupported("text used as TRUE/FALSE")


COERCE = {"num": num, "txt": txt, "bool": as_bool}
FLEX = ("blank", "err")  # types that take on the other side's type


def unify(a: Expr, b: Expr) -> tuple[Expr, Expr]:
    """Give two expressions one type (IF branches, IFERROR, comparisons)."""
    if a.type in FLEX and b.type in FLEX:
        return num(a), num(b)
    if a.type in FLEX:
        a = COERCE[b.type](a)
    if b.type in FLEX:
        b = COERCE[a.type](b)
    if a.type != b.type:
        raise Unsupported(f"mixes {a.type} and {b.type}")
    return a, b


def finalize(e: Expr) -> Expr:
    return ZERO if e.type == "blank" else e  # =A1 on an empty cell shows 0 in Excel


def where_cell(key: Key) -> str:
    sheet, row, col = key
    return f"sheet = {lit_str(sheet).sql} AND row_num = {row} AND col_num = {col}"


def where_area(area, alias: str = "") -> str:
    p = f"{alias}." if alias else ""
    return (f"{p}sheet = {lit_str(area.sheet).sql} AND {p}row_num BETWEEN {area.r1} AND {area.r2} "
            f"AND {p}col_num BETWEEN {area.c1} AND {area.c2}")
```

- [ ] **Step 4: Replace `src/unexcel/compiler.py`**

```python
"""Formula AST -> SQL. Cell mode, run in batches.

Every cell known so far (inputs, then the results of earlier batches) lives in one table with the
columns of INPUT_COLUMNS: kind is n/s/b/e, num holds numbers and TRUE/FALSE as 1/0, txt holds text
and Excel error codes. A batch is up to BATCH consecutive formula cells in dependency order. It
becomes one SQL statement: the cells it computes are materialized CTEs (v = value, e = error code
or NULL), appended to the table it read from. One statement per workbook broke DuckDB at ~1k
formula cells; batches keep every statement small.
"""
from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from unexcel.expr import (BLANK, CELLS, NO_ERR, Expr, error, errs, finalize, lit_num, lit_str, num, txt, unify,
                          where_cell)
from unexcel.functions import FUNCS
from unexcel.graph import formula_order
from unexcel.model import Book, Cell, Key, Unsupported, a1
from unexcel.parser import Binary, Bool, Call, ErrLit, Num, Postfix, RefNode, Str, Unary, parse
from unexcel.refs import Area, resolve

INPUT_COLUMNS = "sheet VARCHAR, row_num INTEGER, col_num INTEGER, kind VARCHAR, num DOUBLE, txt VARCHAR"
SEED_HEADER = ["sheet", "row_num", "col_num", "kind", "num", "txt"]
READ_SEED = ("read_csv(?, header = true, quote = '\"', escape = '\"', columns = {'sheet': 'VARCHAR', "
             "'row_num': 'INTEGER', 'col_num': 'INTEGER', 'kind': 'VARCHAR', 'num': 'DOUBLE', 'txt': 'VARCHAR'})")
VOLATILE = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT"}
ARITH = {"+", "-", "*"}
KIND = {"num": "n", "txt": "s", "bool": "b", "err": "e"}
TYPE = {"n": "num", "s": "txt", "b": "bool"}
NO_VALUE = "no saved value in the file - open and save it in Excel"
DEEP = "formula is nested too deeply to translate"
# ponytail: fixed batch size. Measured on DuckDB 1.5.6: 200k flat cells 71 s, a 100k-row running total 50 s.
# Table mode (Plan 2c) is the real speed-up; change this only with a new measurement.
BATCH = 250


@dataclass(frozen=True)
class Step:
    key: Key
    expr: Expr
    cte: str  # its CTE name inside its batch


@dataclass
class Compiled:
    batches: list[list[Step]]
    unresolved: dict[Key, str]
    findings: list[str]

    @property
    def steps(self) -> list[Step]:
        return [s for batch in self.batches for s in batch]


class NewBatch(Exception):
    """A range reads a cell of the open batch. Ranges read the table, so the batch must close first."""


class Compiler:
    def __init__(self, book: Book):
        self.book = book
        self.done: dict[Key, tuple[str, bool]] = {}  # formula cell in a closed batch -> (type, may be an error)
        self.open: dict[Key, tuple[str, str, bool]] = {}  # formula cell in the open batch -> (cte, type, may err)
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
            return error(node.value)
        if isinstance(node, RefNode):
            area = resolve(node.text, sheet, self.book)
            if not area.is_single:
                raise Unsupported(f"range {node.text} used as a single value")
            return self.cell_value(area.key)
        if isinstance(node, Unary):
            x = num(self.expr(node.operand, sheet))
            return Expr(f"(-{x.sql})", "num", x.err)
        if isinstance(node, Postfix):
            x = num(self.expr(node.operand, sheet))
            return Expr(f"({x.sql} / 100)", "num", x.err)
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
        if op in ARITH or op in ("/", "^"):
            x, y = num(a), num(b)
            if op in ARITH:
                return Expr(f"({x.sql} {op} {y.sql})", "num", errs(x.err, y.err))
            if op == "/":  # DuckDB gives inf, Excel #DIV/0!
                return Expr(f"({x.sql} / NULLIF({y.sql}, 0))", "num",
                            errs(x.err, y.err, f"(CASE WHEN {y.sql} = 0 THEN '#DIV/0!' END)"))
            # DuckDB POWER returns inf or nan where Excel returns an error (verified on 1.5.6)
            bad = (f"(CASE WHEN {x.sql} = 0 AND {y.sql} < 0 THEN '#DIV/0!' "
                   f"WHEN {x.sql} = 0 AND {y.sql} = 0 THEN '#NUM!' "
                   f"WHEN {x.sql} < 0 AND {y.sql} <> FLOOR({y.sql}) THEN '#NUM!' END)")
            return Expr(f"POWER({x.sql}, {y.sql})", "num", errs(x.err, y.err, bad))
        if op == "&":
            x, y = txt(a), txt(b)
            return Expr(f"({x.sql} || {y.sql})", "txt", errs(x.err, y.err))
        a, b = unify(a, b)
        if a.type == "txt":  # Excel text comparison ignores case
            return Expr(f"(LOWER({a.sql}) {op} LOWER({b.sql}))", "bool", errs(a.err, b.err))
        return Expr(f"({a.sql} {op} {b.sql})", "bool", errs(a.err, b.err))

    def cell_value(self, key: Key) -> Expr:
        cell = self.book.cells.get(key)
        if cell is None or (cell.formula is None and cell.kind == "blank"):
            return BLANK
        if key in self.open:
            cte, typ, may_err = self.open[key]
            return Expr(f"(SELECT v FROM {cte})", typ, f"(SELECT e FROM {cte})" if may_err else NO_ERR)
        if cell.formula is None:
            if cell.kind == "e":
                return error(cell.value)
            typ, may_err = TYPE[cell.kind], False
        elif key in self.done:
            typ, may_err = self.done[key]
        else:
            raise Unsupported(f"depends on unresolved cell {a1(key)}")
        where = f"FROM {CELLS} WHERE {where_cell(key)}"
        err = f"(SELECT CASE WHEN kind = 'e' THEN txt END {where})" if may_err else NO_ERR
        if typ == "num":
            return Expr(f"(SELECT num {where})", "num", err)
        if typ == "txt":
            return Expr(f"(SELECT COALESCE(txt, '') {where})", "txt", err)  # the seed CSV reads '' back as NULL
        if typ == "bool":
            return Expr(f"((SELECT num {where}) <> 0)", "bool", err)
        return Expr("NULL", "err", err)

    def ready(self, area: Area) -> list[Key]:
        """The non-empty cells of a range that SQL will read from the table. All must be computed already."""
        keys = area.present(self.book)
        for key in keys:
            if self.book.cells[key].formula is None:
                continue
            if key in self.open:
                raise NewBatch
            if key not in self.done:
                raise Unsupported(f"depends on unresolved cell {a1(key)}")
        return keys

    def kind_of(self, key: Key) -> str:
        """Value kind (n/s/b/e) a cell has in the table when it isn't an error. Call ready() first."""
        cell = self.book.cells[key]
        return KIND[self.done[key][0]] if cell.formula is not None else cell.kind

    def may_err(self, key: Key) -> bool:
        """Whether the table can hold an error for this cell. Call ready() first."""
        cell = self.book.cells[key]
        return self.done[key][1] if cell.formula is not None else cell.kind == "e"


def _close(comp: Compiler, batches: list[list[Step]], current: list[Step]) -> None:
    """Seal the open batch: later cells read its results from the table."""
    if current:
        batches.append(list(current))
        comp.done.update((s.key, (s.expr.type, s.expr.err != NO_ERR)) for s in current)
        comp.open.clear()
        current.clear()


def compile_book(book: Book) -> Compiled:
    order, unresolved = formula_order(book)
    comp = Compiler(book)
    batches: list[list[Step]] = []
    current: list[Step] = []
    for key in order:
        cell = book.cells[key]
        if cell.kind == "blank":
            unresolved[key] = NO_VALUE
            continue
        if len(current) >= BATCH:
            _close(comp, batches, current)
        try:
            try:
                e = comp.compile_cell(cell)
            except NewBatch:
                _close(comp, batches, current)
                e = comp.compile_cell(cell)  # the open batch is empty now, so this can't raise NewBatch
        except Unsupported as err:
            unresolved[key] = str(err)
            continue
        except RecursionError:  # Python's stack, not DuckDB's: DuckDB took 900 nesting levels in tests
            unresolved[key] = DEEP
            continue
        step = Step(key, e, f"c{len(current) + 1}")
        comp.open[key] = (step.cte, e.type, e.err != NO_ERR)
        current.append(step)
    _close(comp, batches, current)
    return Compiled(batches, unresolved, comp.findings)


def render_batch(steps: list[Step], prev: str) -> str:
    """One batch as a SELECT: every row of `prev` plus one row per cell this batch computes."""
    ctes = ",\n".join(f"{s.cte} AS MATERIALIZED (SELECT {s.expr.sql} AS v, CAST({s.expr.err} AS VARCHAR) AS e)"
                      for s in steps)
    rows = []
    for s in steps:
        sheet, row, col = s.key
        typ = s.expr.type
        v, e = f"(SELECT v FROM {s.cte})", f"(SELECT e FROM {s.cte})"
        kind = f"CASE WHEN {e} IS NULL THEN '{KIND[typ]}' ELSE 'e' END"
        num_sql = ("CAST(NULL AS DOUBLE)" if typ in ("txt", "err")
                   else f"CASE WHEN {e} IS NULL THEN CAST({v} AS DOUBLE) END")
        txt_sql = f"COALESCE({e}, CAST({v} AS VARCHAR))" if typ == "txt" else e
        rows.append(f"SELECT {lit_str(sheet).sql}, {row}, {col}, {kind}, {num_sql}, {txt_sql}")
    sql = f"WITH\n{ctes}\nSELECT * FROM {CELLS}\nUNION ALL\n" + "\nUNION ALL\n".join(rows)
    return sql.replace(CELLS, prev)


def input_rows(book: Book) -> list[tuple]:
    """Every non-formula cell as a row of the cells table (see INPUT_COLUMNS)."""
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


def write_seed(book: Book, path: Path) -> None:
    """The input cells as CSV: the dbt seed, and what the verifier loads (0.1 s for 200k rows)."""
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(SEED_HEADER)
        for row in input_rows(book):
            w.writerow(["" if x is None else x for x in row])
```

- [ ] **Step 5: Replace `src/unexcel/functions.py`**

```python
"""Excel function -> SQL. Each entry: fn(compiler, args, sheet) -> Expr.

Ranges are read from the cells table (inputs plus the formula cells of earlier batches).
Compiler.ready() checks every formula cell in a range was computed in an earlier batch.
Every function returns its Excel error in Expr.err; see expr.errs for the order rule.
"""
from __future__ import annotations

from unexcel.expr import CELLS, NO_ERR, ZERO, Expr, as_bool, errs, lit_str, num, unify, where_area
from unexcel.model import Unsupported, a1
from unexcel.parser import Bool, Num, RefNode, Str
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


def note_text_numbers(c, keys):
    """Excel skips text in SUM-like ranges, even text that looks like a number. Tell the user."""
    for key in keys:
        cell = c.book.cells[key]
        if cell.formula is None and cell.kind == "s" and _looks_numeric(cell.value):
            c.findings.append(f"{a1(key)}: number stored as text ({cell.value!r}) is skipped by Excel")


def first_error(c, area: Area, keys) -> str:
    """SQL for the first error in a range, row by row, or NO_ERR when no cell in it can hold one."""
    if not any(c.may_err(k) for k in keys):
        return NO_ERR
    return f"(SELECT txt FROM {CELLS} WHERE {where_area(area)} AND kind = 'e' ORDER BY row_num, col_num LIMIT 1)"


def range_sum(c, area: Area) -> Expr:
    """SUM over a range: text, TRUE/FALSE and blanks are skipped; the first error wins."""
    keys = c.ready(area)
    note_text_numbers(c, keys)
    return Expr(f"(SELECT COALESCE(SUM(num), CAST(0 AS DOUBLE)) FROM {CELLS} WHERE {where_area(area)} AND kind = 'n')",
                "num", first_error(c, area, keys))


@register("SUM")
def sum_(c, args, sheet):
    parts = [range_sum(c, resolve(arg.text, sheet, c.book)) if isinstance(arg, RefNode) else num(c.expr(arg, sheet))
             for arg in args]
    if not parts:
        return ZERO
    return Expr("(" + " + ".join(p.sql for p in parts) + ")", "num", errs(*(p.err for p in parts)))


@register("IF")
def if_(c, args, sheet):
    arity("IF", args, 2, 3)
    cond = as_bool(c.expr(args[0], sheet))
    yes = c.expr(args[1], sheet)
    no = c.expr(args[2], sheet) if len(args) == 3 else Expr("FALSE", "bool")
    yes, no = unify(yes, no)
    taken = NO_ERR if yes.err == no.err == NO_ERR else f"(CASE WHEN {cond.sql} THEN {yes.err} ELSE {no.err} END)"
    return Expr(f"(CASE WHEN {cond.sql} THEN {yes.sql} ELSE {no.sql} END)", yes.type, errs(cond.err, taken))


@register("ROUND")
def round_(c, args, sheet):
    arity("ROUND", args, 2, 2)
    x, digits = num(c.expr(args[0], sheet)), num(c.expr(args[1], sheet))
    # DuckDB ROUND on DOUBLE rounds half away from zero like Excel (2.675 -> 2.68, verified on 1.5.6)
    return Expr(f"ROUND({x.sql}, CAST(TRUNC({digits.sql}) AS INTEGER))", "num", errs(x.err, digits.err))


def _logic(name, joiner):
    def fn(c, args, sheet):
        if not args:
            raise Unsupported(f"{name} needs arguments")
        parts = [as_bool(c.expr(a, sheet)) for a in args]
        return Expr("(" + f" {joiner} ".join(p.sql for p in parts) + ")", "bool", errs(*(p.err for p in parts)))

    return fn


FUNCS["AND"] = _logic("AND", "AND")
FUNCS["OR"] = _logic("OR", "OR")


@register("IFERROR")
def iferror(c, args, sheet):
    arity("IFERROR", args, 2, 2)
    value, fallback = unify(c.expr(args[0], sheet), c.expr(args[1], sheet))
    if value.err == NO_ERR:
        return value
    ok = f"(({value.err}) IS NULL)"
    err = NO_ERR if fallback.err == NO_ERR else f"(CASE WHEN {ok} THEN NULL ELSE {fallback.err} END)"
    return Expr(f"(CASE WHEN {ok} THEN {value.sql} ELSE {fallback.sql} END)", value.type, err)


def _column(c, node, sheet, fname) -> tuple[Area, list]:
    if not isinstance(node, RefNode):
        raise Unsupported(f"{fname} needs a cell range")
    area = resolve(node.text, sheet, c.book)
    return area, c.ready(area)


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
    total, keys = _column(c, args[0], sheet, "SUMIFS")
    if total.c1 != total.c2:
        raise Unsupported("SUMIFS sum range wider than one column")
    keys = list(keys)
    joins, conds = [], []
    for i in range(1, len(args), 2):
        area, area_keys = _column(c, args[i], sheet, "SUMIFS")
        keys += area_keys
        if area.c1 != area.c2 or area.r2 - area.r1 != total.r2 - total.r1:
            raise Unsupported("SUMIFS ranges must be single columns of the same height")
        node = args[i + 1]
        if isinstance(node, Str) and node.value and (node.value[0] in "<>=" or any(ch in node.value for ch in "*?~")):
            raise Unsupported("SUMIFS comparison/wildcard criteria not supported yet")
        crit = c.expr(node, sheet)
        if crit.err != NO_ERR:
            raise Unsupported("SUMIFS criteria that can be an error not supported yet")
        k = f"k{len(joins) + 1}"
        joins.append(f"JOIN {CELLS} {k} ON {k}.sheet = {lit_str(area.sheet).sql} AND {k}.col_num = {area.c1} "
                     f"AND {k}.row_num = s.row_num + {area.r1 - total.r1}")
        conds.append(_match(k, crit))
    if any(c.may_err(k) for k in keys):
        raise Unsupported("SUMIFS over cells that hold errors not supported yet")
    note_text_numbers(c, [k for k in keys if k[2] == total.c1 and k[0] == total.sheet])
    where = " AND ".join([where_area(total, "s"), "s.kind = 'n'", *conds])
    return Expr(f"(SELECT COALESCE(SUM(s.num), CAST(0 AS DOUBLE)) FROM {CELLS} s {' '.join(joins)} WHERE {where})",
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
    table, keys = _column(c, args[1], sheet, "VLOOKUP")
    col = table.c1 + int(args[2].value) - 1
    if not table.c1 <= col <= table.c2:
        raise Unsupported("VLOOKUP column index outside the table")
    kinds = [c.kind_of(k) for k in keys if k[2] == col]
    values = [k for k in kinds if k != "e"]  # an error in the result column is returned as that error
    if len(values) > 1 and values[0] == "s" and set(values[1:]) == {"n"}:
        values = values[1:]  # text header above a numeric column
    if set(values) <= {"n"}:
        column, typ, empty = "num", "num", "CAST(0 AS DOUBLE)"
    elif set(values) == {"s"}:
        column, typ, empty = "txt", "txt", "''"
    else:
        raise Unsupported("VLOOKUP result column mixes types")
    lookup = c.expr(args[0], sheet)
    first_col = Area(table.sheet, table.r1, table.c1, table.r2, table.c1)
    src = (f"FROM {CELLS} k LEFT JOIN {CELLS} r ON r.sheet = k.sheet AND r.row_num = k.row_num "
           f"AND r.col_num = {col} WHERE {where_area(first_col, 'k')} AND {_match('k', lookup)} "
           f"ORDER BY k.row_num LIMIT 1")
    found = f"(SELECT CASE WHEN r.kind = 'e' THEN r.txt ELSE '' END {src})"  # NULL = no row matched
    err = f"(CASE WHEN {found} IS NULL THEN '#N/A' WHEN {found} = '' THEN NULL ELSE {found} END)"
    return Expr(f"(SELECT COALESCE(r.{column}, {empty}) {src})", typ, errs(lookup.err, err))
```

- [ ] **Step 6: Replace `src/unexcel/verifier.py`**

```python
"""Run the compiled batches on DuckDB and compare every cell with Excel's saved value."""
from __future__ import annotations

import math
import tempfile
from pathlib import Path

import duckdb

from unexcel.compiler import READ_SEED, Compiled, render_batch, write_seed
from unexcel.model import Book, Cell, Key, Verdict


def execute(book: Book, compiled: Compiled) -> dict[Key, tuple]:
    """The same statements as the dbt project: seed -> b1 -> b2 -> ... Returns formula cells only."""
    con = duckdb.connect()
    with tempfile.TemporaryDirectory() as tmp:
        seed = Path(tmp) / "input_cells.csv"
        write_seed(book, seed)
        con.execute(f"CREATE TABLE b0 AS SELECT * FROM {READ_SEED}", [str(seed)])
    for k, steps in enumerate(compiled.batches, 1):
        con.execute(f"CREATE TABLE b{k} AS {render_batch(steps, f'b{k - 1}')}")
        con.execute(f"DROP TABLE b{k - 1}")
    wanted = {s.key for s in compiled.steps}
    rows = con.execute(f"SELECT * FROM b{len(compiled.batches)}").fetchall()
    return {(s, r, c): (k, n, t) for s, r, c, k, n, t in rows if (s, r, c) in wanted}


def matches(cell: Cell, kind: str, num, txt) -> bool:
    """Excel's saved value vs the SQL result. An Excel error matches only the same error code."""
    if cell.kind == "e":
        return kind == "e" and txt == cell.value
    if kind == "e" or (num is None and txt is None):
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
    for step in compiled.steps:
        cell = book.cells[step.key]
        kind, n, t = got[step.key]
        actual = t if kind in ("s", "e") else (bool(n) if kind == "b" and n is not None else n)
        out.append(Verdict(step.key, "PASS" if matches(cell, kind, n, t) else "FAIL", cell.value, actual))
    for key, reason in compiled.unresolved.items():
        out.append(Verdict(key, "UNRESOLVED", book.cells[key].value, None, reason))
    return sorted(out, key=lambda v: v.key)
```

- [ ] **Step 7: Edit `src/unexcel/report.py`**

Errors now arrive as their codes (`#DIV/0!`), so `None` means SQL gave no value at all. Replace
```python
    return "(error)" if value is None else str(value).replace("|", "\\|")
```
with
```python
    return "(no value)" if value is None else str(value).replace("|", "\\|")
```

- [ ] **Step 8: Update the Plan 1 assertions that expected "error = NULL"**

- `tests/test_compiler.py`: replace `assert got[("S", 3, 1)] == ("n", None, None)  # error, not inf` with `assert got[("S", 3, 1)] == ("e", None, "#DIV/0!")  # an error, not inf`.
- `tests/test_functions.py`: in `test_sum_propagates_errors_and_handles_empty_ranges` replace `assert got[("S", 1, 2)] == ("n", None, None)` with `assert got[("S", 1, 2)] == ("e", None, "#DIV/0!")`.
- `tests/test_lookup.py`: replace `assert got[("X", 2, 2)] == ("n", None, None)` with `assert got[("X", 2, 2)] == ("e", None, "#N/A")`.
- `tests/test_verifier.py`: replace `assert matches(Cell("S", 1, 1, "x", "#N/A", "e"), "n", None, None)` with `assert matches(Cell("S", 1, 1, "x", "#N/A", "e"), "e", None, "#N/A")`.
- `tests/test_report.py`: replace `assert "| S!A4 | 2.0 | (error) |" in md` with `assert "| S!A4 | 2.0 | (no value) |" in md`.

- [ ] **Step 9: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 10: Commit**

```bash
git add src/unexcel tests
git commit -m "feat: Excel error kinds flow through operators and functions and must match exactly" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: Computed answer key when the file has no saved values

**Files:**
- Create: `src/unexcel/compute.py`
- Replace: `src/unexcel/cli.py`, `src/unexcel/report.py`
- Modify: `pyproject.toml`, `README.md`
- Test: `tests/test_compute.py`

**Interfaces:**
- Consumes: `compile_book`, `NO_VALUE`, `verify`, `write_project`, `run_dbt`, `load`.
- Produces:
  - `compute.fill_missing(book, path) -> tuple[Book, set[Key], str]`: a Book whose formula cells with no saved value get the value `formulas` computes, the set of keys filled, and (when nothing could be computed) the reason to append to `NO_VALUE`. Never touches a cell Excel saved.
  - `report.write_report(verdicts, findings, out, source, computed=frozenset())`: `report.json` cells gain `"answer": "excel-saved" | "computed"`; `REPORT.md` says how many answers were computed and marks them `(computed)` in the failed table.
  - Extra `compute = ["formulas>=1.2"]`.

- [ ] **Step 1: Write the failing tests `tests/test_compute.py`** (Review Focus 5)

```python
import json
import sys

import pytest
from xlsxkit import SALES, write_xlsx

from unexcel.cli import main
from unexcel.compute import fill_missing
from unexcel.reader import load


def test_without_the_formulas_library_cells_stay_unresolved_with_a_hint(tmp_path, monkeypatch):
    monkeypatch.setitem(sys.modules, "formulas", None)  # makes `import formulas` raise ImportError
    wb = tmp_path / "script.xlsx"
    write_xlsx(wb, {"S": {"A1": 2, "A2": ("=A1*2", None)}})
    assert main(["convert", str(wb)]) == 0
    md = (tmp_path / "script_dbt" / "REPORT.md").read_text(encoding="utf-8")
    assert ("| S!A2 | no saved value in the file - open and save it in Excel "
            "(or install unexcel[compute] to compute it) |") in md


def test_sheet_names_with_apostrophes_are_not_computed(tmp_path):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"Bob's": {"A1": 2, "A2": ("=A1*2", None)}})
    _, computed, why = fill_missing(load(wb), wb)
    assert computed == set() and "apostrophe" in why


def test_saved_values_are_never_replaced(tmp_path):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2, "A2": ("=A1*2", 5)}})
    book = load(wb)
    assert fill_missing(book, wb) == (book, set(), "")


def test_computes_missing_answers_and_labels_them(tmp_path):
    pytest.importorskip("formulas")
    wb = tmp_path / "script.xlsx"
    write_xlsx(wb, {"Sales": SALES, "C": {
        "A1": "Asha",
        "B1": ("=SUMIFS(Sales!C:C, Sales!A:A, A1)", None),
        "B2": ("=1/0", None),
        "B3": ('=A1&"!"', None),
        "B4": ("=B1>5", None),
    }})
    assert main(["convert", str(wb)]) == 0
    out = tmp_path / "script_dbt"
    assert "4 formula cells had no saved value" in (out / "REPORT.md").read_text(encoding="utf-8")
    cells = json.loads((out / "report.json").read_text(encoding="utf-8"))["cells"]
    assert {c["cell"]: (c["status"], c["answer"]) for c in cells} == {
        f"C!B{r}": ("PASS", "computed") for r in range(1, 5)}
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_compute.py -v`
Expected: FAIL, `ModuleNotFoundError: No module named 'unexcel.compute'`.

- [ ] **Step 3: Write `src/unexcel/compute.py`**

```python
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
    if isinstance(x, str):
        return x, "s"
    if type(x).__name__ == "XlError":
        return str(x), "e"
    return None, "blank"  # formulas' empty token: leave the cell unresolved


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
    return Book(cells, book.sheets, book.names, book.max_row), computed, ""
```

- [ ] **Step 4: Replace `src/unexcel/report.py`**

```python
"""REPORT.md for people, report.json for tools."""
from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

from unexcel.model import Key, Verdict, a1


def _md(value) -> str:
    return "(no value)" if value is None else str(value).replace("|", "\\|")


def write_report(verdicts: list[Verdict], findings: list[str], out: Path, source: str,
                 computed: frozenset[Key] | set[Key] = frozenset()) -> Counter:
    counts = Counter({"PASS": 0, "FAIL": 0, "UNRESOLVED": 0})
    counts.update(v.status for v in verdicts)
    findings = list(dict.fromkeys(findings))

    data = {
        "source": source,
        "counts": dict(counts),
        "findings": findings,
        "cells": [{"cell": a1(v.key), "status": v.status, "expected": v.expected, "actual": v.actual,
                   "reason": v.reason, "answer": "computed" if v.key in computed else "excel-saved"}
                  for v in verdicts],
    }
    (out / "report.json").write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    lines = [f"# unexcel report: {source}", "", "| Status | Cells |", "|---|---|"]
    lines += [f"| {status} | {counts[status]} |" for status in ("PASS", "FAIL", "UNRESOLVED")]
    if computed:
        lines += ["", f"{len(computed)} formula cells had no saved value. Their expected values were computed "
                      "by the `formulas` library, not by Excel, and are marked (computed)."]
    if findings:
        lines += ["", "## Findings in the workbook", ""] + [f"- {f}" for f in findings]
    failed = [v for v in verdicts if v.status == "FAIL"]
    if failed:
        lines += ["", "## Failed cells", "", "| Cell | Excel | SQL |", "|---|---|---|"]
        lines += [f"| {a1(v.key)} | {_md(v.expected)}{' (computed)' if v.key in computed else ''} | {_md(v.actual)} |"
                  for v in failed]
    unresolved = [v for v in verdicts if v.status == "UNRESOLVED"]
    if unresolved:
        lines += ["", "## Unresolved cells", "", "| Cell | Reason |", "|---|---|"]
        lines += [f"| {a1(v.key)} | {_md(v.reason)} |" for v in unresolved]
    (out / "REPORT.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return counts
```

- [ ] **Step 5: Replace `src/unexcel/cli.py`**

```python
"""unexcel convert WORKBOOK [--out DIR] [--dbt]"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from unexcel.compiler import NO_VALUE, compile_book
from unexcel.compute import fill_missing
from unexcel.emitter import run_dbt, write_project
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
    convert.add_argument("--dbt", action="store_true", help="also run `dbt build` on the result")
    args = parser.parse_args(argv)

    try:
        book = load(args.workbook)
    except UnreadableWorkbook as err:
        print(f"error: {err}", file=sys.stderr)
        return 2
    book, computed, why = fill_missing(book, args.workbook)
    compiled = compile_book(book)
    if why:
        compiled.unresolved.update({k: f"{r} ({why})" for k, r in compiled.unresolved.items() if r == NO_VALUE})
    verdicts = verify(book, compiled)
    out = args.out or args.workbook.with_name(args.workbook.stem + "_dbt")
    write_project(book, compiled, out, args.workbook.stem)
    counts = write_report(verdicts, compiled.findings, out, args.workbook.name, computed)
    if computed:
        print(f"{len(computed)} expected values computed by the formulas library (the file had none saved)")
    print(f"{counts['PASS']} pass, {counts['FAIL']} fail, {counts['UNRESOLVED']} unresolved -> {out / 'REPORT.md'}")
    code = 1 if counts["FAIL"] else 0
    if args.dbt:
        ok, log = run_dbt(out)
        print("dbt build: ok" if ok else f"dbt build failed:\n{log}")
        code = code if ok else 1
    return code
```

- [ ] **Step 6: Add the extra to `pyproject.toml`**

In `[project.optional-dependencies]` add a line after `dbt = [...]`:
```toml
compute = ["formulas>=1.2"]
```
Append to `README.md`:
```markdown

Optional extras: `pip install unexcel[dbt]` runs `dbt build` on the result (`--dbt`); `pip install unexcel[compute]` computes the expected values of workbooks that were never saved in Excel (labelled `computed` in the report).
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `python -m pip install -e ".[dev,dbt,compute]"` then `python -m pytest tests -v`
Expected: all passed. `test_computes_missing_answers_and_labels_them` runs because `formulas` is installed (it takes several seconds: `formulas` expands `C:C` to a million rows).

- [ ] **Step 8: Commit**

```bash
git add src/unexcel/compute.py src/unexcel/cli.py src/unexcel/report.py pyproject.toml README.md tests/test_compute.py
git commit -m "feat: computed answer key via formulas when the file has no saved values" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: Real-Excel ground truth for error codes

**Files:**
- Create: `examples/make_errors.py`
- Modify: `tests/test_excel_saved.py` (append), `.gitignore`

**Interfaces:**
- Consumes: `load`, `compile_book`, `verify`.
- Produces: `examples/errors.xlsx` (saved by real Excel, committed). Confirms the error rules Task 4 assumed, especially "SUM returns the first error row by row".

- [ ] **Step 1: Create `examples/make_errors.py`**

```python
"""Writes examples/errors_unsaved.xlsx: one cell per Excel error rule unexcel relies on.

Open it in real Microsoft Excel, then File > Save As examples/errors.xlsx.
Excel stores its own answers in the file; tests/test_excel_saved.py checks unexcel against them.
"""
from pathlib import Path

import openpyxl

CELLS = {
    "A1": 0, "A2": -8, "A3": "South",
    "B1": "=1/A1", "B2": "=A1^-1", "B3": "=A1^0", "B4": "=A2^(1/3)", "B5": "=#N/A",
    "B6": "=B5+1", "B7": "=B1+B5", "B8": "=IFERROR(1/A1,-1)", "B9": "=IF(1/A1>1,1,2)",
    # SUM reports the first error row by row: E1 (#DIV/0!) comes before D2 (#N/A)
    "D1": 1, "E1": "=1/0", "D2": "=#N/A", "E2": 2, "F1": "=SUM(D1:E2)",
    # VLOOKUP: a missing key is #N/A; a matched row returns the error its cell holds
    "G1": "West", "H1": 0.05, "G2": "South", "H2": "=1/0",
    "I1": '=VLOOKUP("North",G1:H2,2,FALSE)', "I2": "=VLOOKUP(A3,G1:H2,2,FALSE)",
}

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Errors"
for ref, value in CELLS.items():
    ws[ref] = value
wb.save(Path(__file__).parent / "errors_unsaved.xlsx")
```

- [ ] **Step 2: Append to `tests/test_excel_saved.py`**

```python
ERRORS_EXAMPLE = Path(__file__).parent.parent / "examples" / "errors.xlsx"


@pytest.mark.skipif(not ERRORS_EXAMPLE.exists(), reason="save examples/errors.xlsx from real Excel (PLAN-2a Task 6)")
def test_real_excel_error_codes_match():
    book = load(ERRORS_EXAMPLE)
    verdicts = verify(book, compile_book(book))
    assert [v for v in verdicts if v.status != "PASS"] == []
    assert len(verdicts) == 15  # B1:B9, D2, E1, F1, H2, I1, I2
```

Append `examples/errors_unsaved.xlsx` to `.gitignore`.

- [ ] **Step 3: Run the tests**

Run: `python examples/make_errors.py` then `python -m pytest tests -v`
Expected: all passed; `test_real_excel_error_codes_match` is skipped until Step 4.

- [ ] **Step 4: MANUAL, needs the user's Excel.** Open `examples/errors_unsaved.xlsx` in Microsoft Excel and save it as `examples/errors.xlsx`. Then:

Run: `python -m pytest tests/test_excel_saved.py -v`
Expected: passed. If `F1` FAILs with `#N/A`, Excel's SUM reports errors column by column, not row by row: change `ORDER BY row_num, col_num` to `ORDER BY col_num, row_num` in `functions.first_error` and the expectation in `tests/test_errors.py::test_sum_returns_the_first_error_row_by_row`. Any other FAIL: fix the compiler, not the expectation.

- [ ] **Step 5: Commit**

```bash
git add examples/make_errors.py examples/errors.xlsx tests/test_excel_saved.py .gitignore
git commit -m "test: real-Excel ground truth for error codes" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: 200k-formula performance check

**Files:**
- Create: `tests/test_perf.py`

**Interfaces:**
- Consumes: `write_xlsx`, `load`, `compile_book`, `verify`, `write_project`.
- Produces: a timed, opt-in test (`UNEXCEL_PERF=1`) that converts a generated 200,002-formula workbook and prints the time per stage. CI runs it in Plan 4.

- [ ] **Step 1: Write `tests/test_perf.py`**

```python
"""200k-formula workbook, timed (DESIGN.md §2.1 and §7). Slow, so opt-in:

    UNEXCEL_PERF=1 python -m pytest tests/test_perf.py -s
"""
import os
import time

import pytest
from xlsxkit import write_xlsx

from unexcel.compiler import compile_book
from unexcel.emitter import write_project
from unexcel.reader import load
from unexcel.verifier import verify

pytestmark = pytest.mark.skipif(not os.environ.get("UNEXCEL_PERF"), reason="slow: set UNEXCEL_PERF=1")
ROWS = 50_000  # four formula columns -> 200,000 formula cells, plus two summary cells


def make(path):
    """A finance-shaped sheet: amount, amount x rate, running total, threshold, difference."""
    data, total, d_sum = {}, 0.0, 0.0
    for r in range(1, ROWS + 1):
        amount = float(r % 997)
        b = amount * 1.1
        total = b if r == 1 else total + b  # same additions, same order as SQL: exact match
        d = b if b > 100 else 0.0
        d_sum += d
        data[f"A{r}"] = amount
        data[f"B{r}"] = (f"=A{r}*1.1", b)
        data[f"C{r}"] = ("=B1" if r == 1 else f"=C{r - 1}+B{r}", total)
        data[f"D{r}"] = (f"=IF(B{r}>100,B{r},0)", d)
        data[f"E{r}"] = (f"=D{r}-A{r}", d - amount)
    summary = {"A1": ("=SUM(Data!D:D)", d_sum), "A2": (f"=Data!C{ROWS}", total)}
    write_xlsx(path, {"Data": data, "Summary": summary})


def test_200k_formulas_under_five_minutes(tmp_path):
    wb = tmp_path / "big.xlsx"
    make(wb)
    marks = [("start", time.perf_counter())]
    book = load(wb)
    marks.append(("read", time.perf_counter()))
    compiled = compile_book(book)
    marks.append(("compile", time.perf_counter()))
    verdicts = verify(book, compiled)
    marks.append(("verify", time.perf_counter()))
    write_project(book, compiled, tmp_path / "out", "big")
    marks.append(("emit", time.perf_counter()))
    for (_, t0), (name, t1) in zip(marks, marks[1:]):
        print(f"{name:8s} {t1 - t0:6.1f} s")
    total_s = marks[-1][1] - marks[0][1]
    print(f"{'total':8s} {total_s:6.1f} s  ({len(compiled.batches)} batches)")
    assert len(verdicts) == 4 * ROWS + 2
    assert {v.status for v in verdicts} == {"PASS"}
    assert total_s < 300  # DESIGN.md §2.1: under 5 minutes on a laptop
```

- [ ] **Step 2: Run it**

Run: `UNEXCEL_PERF=1 python -m pytest tests/test_perf.py -s -v` (Windows PowerShell: `$env:UNEXCEL_PERF=1; python -m pytest tests/test_perf.py -s -v`)
Expected: passed, printing the per-stage times. Measured prototypes put `verify` near 70 s and `read` near 2 s. Paste the printed times into the PR description. If `total` is over 300 s, the test fails; record which stage dominates in the PR and stop: the fix (table mode) is Plan 2c, not a tweak here.

- [ ] **Step 3: Run the default suite to confirm it stays fast**

Run: `python -m pytest tests -v`
Expected: all passed, `test_perf.py` skipped.

- [ ] **Step 4: Commit**

```bash
git add tests/test_perf.py
git commit -m "test: opt-in 200k-formula performance check" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## Self-review (done while writing)

- **Spec coverage, 2a scope:** DESIGN §2.1 (200k formulas, under 5 min, measured) → T3 batching, T2 index, T7 measurement; §2.3 match rule incl. errors compared explicitly → T4; §2.4 computed answer key labelled `computed` → T5; §6 "Excel errors: SQL NULL + error-kind column, compared by kind" → T4 (error code in `txt`, kind `e`); §6 "No cached values → compute with `formulas`, label computed" → T5; §7 perf check → T7, real-Excel ground truth → T6. Function library (§2.2 ~80 functions) and table mode/regions (§4) are Plans 2b and 2c, listed at the top.
- **Plan 1 bugs found and fixed here:** one-statement SQL crashes at ~1k cells (T3); `RecursionError` crashes the run (T1); whole-column refs walk every row (T2); `POWER` returns `nan`/`inf` silently (T4); `""` text read back as NULL (T3).
- **Placeholders:** none. Every code step has full code; files changed heavily are given whole.
- **Name consistency:** `CELLS` replaces `INPUT` everywhere from T3; `Step(key, expr, cte)`, `Compiled.batches` / `.steps`, `render_batch(steps, prev)`, `write_seed`, `READ_SEED`, `NO_VALUE`, `DEEP` defined in T3 (`DEEP` first in T1) and used in T4/T5/T7; `Compiler.done` is `type` in T3 and `(type, may_err)` from T4, and T4 replaces every reader of it (`cell_value`, `kind_of`, `may_err`, `_close`); `note_text_numbers(c, keys)` takes keys from T3 on; `write_report(..., computed)` defaults so T4 callers are unchanged.
- **Verified against real libraries (2026-10-11, DuckDB 1.5.6, openpyxl 3.1.5, formulas 1.3.4):** depth error / segfault in Plan 1's single statement; batched chain timings; CSV load speed and exact round-trip; `read_csv(?, header = true, columns = …)` with a parameter inside `CREATE TABLE AS`, header-only file → 0 rows, `""` → NULL; `POWER` edge values; `COUNT(*) FILTER`, `NOT EXISTS` anti-join, `ORDER BY … LIMIT 1` first error; `formulas` result types (`str`, `bool`, `int`/`float`, `XlError`, empty `Token`), key format `'[book.xlsx]SHEET'!A1`, the apostrophe bug.
- **Not verified, confirmed by Task 6:** Excel's SUM reports the first error row by row.
