# Plan 2d: Table Mode, Slice 1 (Row-wise Columns) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A column that repeats one formula down its rows compiles to one readable, set-based SQL statement instead of one CTE per cell, with identical verdicts.

**Architecture:** `regions.py` finds blocks of pasted-down formula columns. The dependency graph treats each block as one node. `region.py` compiles a block with the existing `Compiler` (same-row references become column references) into a wide readable model plus a mechanical unpivot into the long `cells` table. Anything a region cannot express stays in cell mode, in the same project.

**Tech Stack:** Python 3.10+, DuckDB, openpyxl (its formula tokenizer), pytest. No new runtime dependencies.

**Spec:** `docs/superpowers/specs/2026-10-14-plan-2d-table-mode-design.md` (read it first). Plan 2c is merged; this plan builds on it (`Compiler.ready()` summaries, `RangeIndex`, append-in-place verifier, dbt checkpoints).

## Global Constraints

- Runtime dependencies stay `openpyxl` and `duckdb` only (AGENTS.md section 7).
- A false PASS is the worst bug: when unsure, a column stays in cell mode or is UNRESOLVED. Every formula cell still gets PASS, FAIL or UNRESOLVED against Excel's saved value.
- Verdicts must be identical with and without regions: `compile_book(book, regions=True)` and `compile_book(book, regions=False)` give the same status and value for every cell.
- No DuckDB exception may escape from a supported function; cell values never leave the machine.
- Tests first, watch them fail. Console output stays ASCII-only.
- Python is `.venv/Scripts/python.exe` on Windows. Prefix commands that need dbt with `PATH="$PWD/.venv/Scripts:$PATH"`.
- Commit messages end with `Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>`. Do not push, merge or open a PR without the user's go-ahead.
- Slice 1 only: row-wise columns. No previous-row references, growing ranges, `SUMIFS`/lookups as joins, blanks, errors or mixed kinds in input columns (all stay in cell mode).

## Review Focus

Inputs the spec implies but no obvious task exercises. Each line has its test in the task named.

1. An input column read by a region has a blank cell or an error cell in the region's rows: the reading column must stay in cell mode (Excel's blank-as-zero and error semantics are not reproduced by the pivot). (Task 5, Task 6)
2. An absolute reference that points INTO the region's own rows (`$A$3` inside a region over rows 2..9) and a mixed reference (`B$2`, `$A2`): the first is one scalar for every row, the second is row-wise only when its row part is relative. (Task 1, Task 5)
3. A region that creates a cycle the cell graph does not have (the region reads a cell that reads the region): it must dissolve to cell mode with no cell lost. (Task 2, Task 3)
4. Relative cross-sheet references (`Other!A2`) must keep the column in cell mode; absolute cross-sheet references (`Other!$A$1`) are fine. (Task 1, Task 5)
5. Duplicate, empty or numeric header text above a region, and a workbook with no regions at all: names fall back to `col_<letter>` and nothing else changes. (Task 5, Task 8)

---

### Task 1: Reference classification and R1C1 keys

**Files:**
- Create: `src/unexcel/r1c1.py`
- Test: `tests/test_r1c1.py`

**Interfaces:**
- Produces:
  - `ROWREF = "__ROWREF_{}__"`: placeholder text for a same-row reference to column `n`.
  - `classify(text: str, row: int) -> tuple[str, int]`: `("same_row", col)` for one relative-row cell of the same sheet whose row equals `row`; `("abs", 0)` when the reference reads the same cells on every row (all row parts anchored with `$`, whole columns, names); `("bad", 0)` for anything that moves with the row otherwise (previous/next row, relative ranges, relative cells on other sheets, whole rows).
  - `r1c1_key(formula: str, row: int, col: int) -> str`: the formula with each reference written relative to its own cell; equal keys mean the same formula pasted down.
  - `map_refs(node, f)`: the AST with every `RefNode` replaced by `f(node)`.

- [ ] **Step 1: Write the failing test** (`tests/test_r1c1.py`)

```python
from unexcel.parser import RefNode, parse, refs
from unexcel.r1c1 import classify, map_refs, r1c1_key


def test_same_row_references():
    assert classify("A2", 2) == ("same_row", 1)
    assert classify("$A2", 2) == ("same_row", 1)  # anchored column, relative row: still row-wise
    assert classify("C2", 2) == ("same_row", 3)


def test_references_that_read_the_same_cells_on_every_row():  # Review Focus 2
    assert classify("A$2", 2) == ("abs", 0)
    assert classify("$A$2", 2) == ("abs", 0)
    assert classify("$A$2", 5) == ("abs", 0)  # anchored, even when it points into the region's own rows
    assert classify("$A$1:$A$9", 3) == ("abs", 0)
    assert classify("A:A", 2) == ("abs", 0)
    assert classify("Rate", 2) == ("abs", 0)  # a name
    assert classify("Other!$A$1", 2) == ("abs", 0)  # Review Focus 4


def test_references_that_move_with_the_row_are_bad():  # Review Focus 4
    assert classify("A1", 2)[0] == "bad"  # previous row
    assert classify("A3", 2)[0] == "bad"  # next row
    assert classify("Other!A2", 2)[0] == "bad"  # relative on another sheet
    assert classify("A2:B2", 2)[0] == "bad"  # relative range
    assert classify("$A$2:A9", 2)[0] == "bad"  # growing range
    assert classify("2:2", 2)[0] == "bad"  # whole row


def test_r1c1_key_is_equal_for_a_formula_pasted_down():
    assert r1c1_key("A2*1.1", 2, 2) == r1c1_key("A3*1.1", 3, 2)
    assert r1c1_key("A2*$J$1", 2, 2) == r1c1_key("A9*$J$1", 9, 2)
    assert r1c1_key('IF(A2>1,"x","y")', 2, 3) == r1c1_key('IF(A7>1,"x","y")', 7, 3)


def test_r1c1_key_differs_when_the_formula_differs():
    assert r1c1_key("A3*1.1", 3, 2) != r1c1_key("A3*1.1", 4, 2)  # a different row of the same text
    assert r1c1_key("A2*$J$1", 2, 2) != r1c1_key("A2*$J$2", 2, 2)
    assert r1c1_key("A2*1.1", 2, 2) != r1c1_key("A2*1.2", 2, 2)
    assert r1c1_key("A2*1.1", 2, 2) != r1c1_key("A2*1.1", 2, 3)  # relative column offset differs


def test_map_refs_replaces_every_reference():
    ast = parse("A2*B2+SUM($C$1:$C$9)")
    out = map_refs(ast, lambda n: RefNode("X") if n.text in ("A2", "B2") else n)
    assert list(refs(out)) == ["X", "X", "$C$1:$C$9"]
```

- [ ] **Step 2: Run it, expect failure**

Run: `.venv/Scripts/python.exe -m pytest tests/test_r1c1.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'unexcel.r1c1'`.

- [ ] **Step 3: Write `src/unexcel/r1c1.py`**

```python
"""Reference classification and R1C1 keys: which formulas are 'the same formula pasted down a column'."""
from __future__ import annotations

import re

from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils.cell import column_index_from_string

from unexcel.parser import Binary, Call, Postfix, RefNode, Unary

CELL = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$")
ROW_ONLY = re.compile(r"^\$?[0-9]+$")
ROWREF = "__ROWREF_{}__"


def _split(text: str) -> tuple[str, str]:
    if "!" in text:
        sheet, ref = text.rsplit("!", 1)
        return sheet + "!", ref
    return "", text


def classify(text: str, row: int) -> tuple[str, int]:
    """How a reference behaves when its formula is copied down.
    ('same_row', col): one cell of the same sheet and the same row, row part not anchored.
    ('abs', 0): reads the same cells on every row (all row parts anchored, whole columns, names).
    ('bad', 0): anything else that moves with the row (previous/next row, relative ranges, other sheets, whole rows)."""
    prefix, ref = _split(text)
    parts = ref.split(":")
    relative = []
    for part in parts:
        m = CELL.match(part)
        if m:
            relative.append((not m.group(3), column_index_from_string(m.group(2).upper()), int(m.group(4))))
        elif ROW_ONLY.match(part):
            return "bad", 0  # whole-row reference
        # whole columns (A:A) and names read the same cells on every row
    if not any(is_relative for is_relative, _, _ in relative):
        return "abs", 0
    if len(parts) == 1 and not prefix and relative[0][2] == row:
        return "same_row", relative[0][1]
    return "bad", 0


def _norm(part: str, row: int, col: int) -> str:
    m = CELL.match(part)
    if not m:
        return part
    c = column_index_from_string(m.group(2).upper())
    cs = f"C{c}" if m.group(1) else f"C[{c - col}]"
    rs = f"R{m.group(4)}" if m.group(3) else f"R[{int(m.group(4)) - row}]"
    return rs + cs


def r1c1_key(formula: str, row: int, col: int) -> str:
    """The formula with every reference written relative to its own cell: equal keys = same formula pasted down."""
    out = []
    for tok in Tokenizer("=" + formula).items:
        if tok.type == Token.OPERAND and tok.subtype == Token.RANGE:
            prefix, ref = _split(tok.value)
            out.append(prefix + ":".join(_norm(p, row, col) for p in ref.split(":")))
        else:
            out.append(tok.value)
    return "".join(out)


def map_refs(node, f):
    """The tree with every RefNode replaced by f(node)."""
    if isinstance(node, RefNode):
        return f(node)
    if isinstance(node, (Unary, Postfix)):
        return type(node)(node.op, map_refs(node.operand, f))
    if isinstance(node, Binary):
        return Binary(node.op, map_refs(node.left, f), map_refs(node.right, f))
    if isinstance(node, Call):
        return Call(node.name, tuple(None if a is None else map_refs(a, f) for a in node.args))
    return node
```

- [ ] **Step 4: Run it, expect pass**

Run: `.venv/Scripts/python.exe -m pytest tests/test_r1c1.py -q`
Expected: `6 passed`.

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/r1c1.py tests/test_r1c1.py
git commit -m "feat: classify references and build R1C1 keys for pasted-down formulas"
```

---

### Task 2: Region detection

**Files:**
- Create: `src/unexcel/regions.py`
- Test: `tests/test_regions.py`

**Interfaces:**
- Consumes: `classify`, `r1c1_key`, `map_refs`, `ROWREF` (Task 1); `parse`, `refs`, `RefNode`; `resolve`; `Book._rows_by_col`.
- Produces:
  - `MIN_ROWS = 3`.
  - `Region` dataclass: `sheet, r1, r2, cols: list[int]` (sorted), `template: dict[int, node]` (AST of the first row's formula per column, same-row references replaced by `RefNode(ROWREF.format(source_col))`), `reads: dict[int, set[int]]` (source columns read in the same row, per column); `keys() -> list[Key]`; `first_key -> Key` = `(sheet, r1, min(cols))`.
  - `RegionScan(regions: list[Region], rejected: list[str])`.
  - `find_regions(book) -> RegionScan`.
  - `column_order(cols, reads) -> tuple[list[int], list[int]]`: `(order, stuck)`, dependency order with the smallest ready column first; `stuck` are columns in a same-row cycle.
  - `order_with_regions(book, regions) -> tuple[list[Key], dict[Key, str], list[Region]]`: ordering plus the regions that survived (implemented in Task 3 once the graph supports groups; this task creates the stub that raises `NotImplementedError`).

- [ ] **Step 1: Write the failing tests** (`tests/test_regions.py`)

```python
from xlsxkit import book_of

from unexcel.regions import MIN_ROWS, column_order, find_regions


def sheet(n=6):
    spec = {"A1": "amount", "B1": "tax"}
    for r in range(2, n + 2):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*1.1", 0.0)  # row-wise
        spec[f"C{r}"] = (f"=SUM($A$2:A{r})", 0.0)  # growing range: stays in cell mode
        spec[f"D{r}"] = (f"=A{r - 1}+1", 0.0)  # previous row: stays in cell mode
    return book_of({"S": spec})


def spans(scan):
    return [(r.sheet, r.r1, r.r2, r.cols) for r in scan.regions]


def test_finds_only_the_row_wise_column_and_says_why_the_rest_stayed():
    scan = find_regions(sheet())
    assert spans(scan) == [("S", 2, 7, [2])]
    notes = " ".join(scan.rejected)
    assert "S!C2:C7" in notes and "S!D2:D7" in notes


def test_a_run_shorter_than_min_rows_is_not_a_region():
    assert MIN_ROWS == 3
    book = book_of({"S": {"A1": 1.0, "A2": 2.0, "B1": ("=A1*2", 0.0), "B2": ("=A2*2", 0.0)}})
    assert find_regions(book).regions == []


def test_columns_with_the_same_rows_form_one_region():
    spec = {}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=B{r}+1", 0.0)
    scan = find_regions(book_of({"S": spec}))
    assert spans(scan) == [("S", 2, 5, [2, 3])]
    region = scan.regions[0]
    assert region.reads == {2: {1}, 3: {2}}
    assert region.first_key == ("S", 2, 2)
    assert region.keys()[0] == ("S", 2, 2) and len(region.keys()) == 8


def test_different_row_spans_are_separate_regions():
    spec = {}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
    for r in range(2, 8):
        spec[f"C{r}"] = ("=$A$2+B$2", 0.0)  # the same formula on every row, rows 2..7
    scan = find_regions(book_of({"S": spec}))
    assert spans(scan) == [("S", 2, 5, [2]), ("S", 2, 7, [3])]


def test_a_range_over_the_regions_own_formula_cells_keeps_the_column_out():
    spec = {}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=SUM($B$2:$B$5)+A{r}", 0.0)  # absolute range over column B's formula cells
    scan = find_regions(book_of({"S": spec}))
    assert spans(scan) == [("S", 2, 5, [2])]


def test_a_same_row_cycle_between_columns_keeps_both_out():
    spec = {}
    for r in range(2, 6):
        spec[f"B{r}"] = (f"=C{r}+1", 0.0)
        spec[f"C{r}"] = (f"=B{r}+1", 0.0)
        spec[f"D{r}"] = (f"=A{r}*2", 0.0)
        spec[f"A{r}"] = float(r)
    assert spans(find_regions(book_of({"S": spec}))) == [("S", 2, 5, [4])]


def test_array_formulas_and_formulas_without_a_saved_value_break_runs():
    spec = {f"A{r}": float(r) for r in range(2, 8)}
    spec.update({f"B{r}": (f"=A{r}*2", 0.0) for r in range(2, 8)})
    spec["B5"] = ("=A5*2", None)  # no saved value: kind blank
    scan = find_regions(book_of({"S": spec}))
    assert spans(scan) == [("S", 2, 4, [2])]  # rows 6..7 are too few for a second run


def test_column_order_is_dependency_order_with_the_smallest_ready_column_first():
    reads = {2: {1}, 3: {2}, 4: {1}, 5: {3, 4}}
    assert column_order([2, 3, 4, 5], reads) == ([2, 3, 4, 5], [])
    assert column_order([3, 2], {3: {2}, 2: {3}}) == ([], [2, 3])
```

- [ ] **Step 2: Run, expect failure**

Run: `.venv/Scripts/python.exe -m pytest tests/test_regions.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'unexcel.regions'`.

- [ ] **Step 3: Write `src/unexcel/regions.py`**

```python
"""Find row-wise regions: blocks of columns that repeat one formula down the same rows.

A region column is a run of at least MIN_ROWS consecutive formula cells with the same R1C1 key. Columns whose runs
cover the same rows form one region. A column joins only if every reference in its formula is a same-row cell, or
reads the same cells on every row; anything that moves with the row (a previous row, a growing range) stays in cell
mode. The region compiler decides later whether the SQL can really be built.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl.utils.cell import get_column_letter

from unexcel.model import Book, Key, Unsupported
from unexcel.parser import RefNode, parse, refs
from unexcel.r1c1 import ROWREF, classify, map_refs, r1c1_key
from unexcel.refs import resolve

MIN_ROWS = 3


@dataclass
class Region:
    sheet: str
    r1: int
    r2: int
    cols: list[int]
    template: dict[int, object] = field(default_factory=dict)  # col -> AST, same-row refs as __ROWREF_n__
    reads: dict[int, set] = field(default_factory=dict)  # col -> source columns read in the same row

    def keys(self) -> list[Key]:
        return [(self.sheet, r, c) for c in self.cols for r in range(self.r1, self.r2 + 1)]

    @property
    def first_key(self) -> Key:
        return (self.sheet, self.r1, min(self.cols))


@dataclass
class RegionScan:
    regions: list[Region]
    rejected: list[str]  # why pasted-down columns did not become regions: the map for the next slice


def column_order(cols, reads) -> tuple[list[int], list[int]]:
    """Columns in dependency order (smallest ready column first, so the SQL follows the sheet); and the stuck ones."""
    left, order = set(cols), []
    while left:
        ready = sorted(c for c in left if not (reads[c] & left))
        if not ready:
            break
        order.append(ready[0])
        left.discard(ready[0])
    return order, sorted(left)


def _runs(book: Book):
    """(sheet, col, r1, r2) for every maximal run of one pasted-down formula, at least MIN_ROWS rows."""
    for (sheet, col), rows in sorted(book._rows_by_col.items()):
        run, key0 = [], None
        for row in rows + [None]:
            cell = book.cells[(sheet, row, col)] if row is not None else None
            ok = cell is not None and cell.formula is not None and not cell.array and cell.kind != "blank"
            key = r1c1_key(cell.formula, row, col) if ok else None
            if ok and run and row == run[-1] + 1 and key == key0:
                run.append(row)
                continue
            if len(run) >= MIN_ROWS:
                yield sheet, col, run[0], run[-1]
            run, key0 = ([row], key) if ok else ([], None)


def _place(sheet: str, col: int, r1: int, r2: int) -> str:
    letter = get_column_letter(col)
    return f"{sheet}!{letter}{r1}:{letter}{r2}"


def find_regions(book: Book) -> RegionScan:
    groups: dict[tuple, list[int]] = {}
    for sheet, col, r1, r2 in _runs(book):
        groups.setdefault((sheet, r1, r2), []).append(col)
    regions, rejected = [], []
    for (sheet, r1, r2), cols in groups.items():
        template, reads, keep = {}, {}, []
        for col in cols:
            cell = book.cells[(sheet, r1, col)]
            try:
                ast = parse(cell.formula)
                found, why = set(), None
                for text in refs(ast):
                    kind, source = classify(text, r1)
                    if kind == "bad":
                        why = f"reads {text}, which moves with the row (previous row, growing or relative range, other sheet)"
                    elif kind == "same_row":
                        found.add(source)
                    else:  # the same cells on every row: must not be the region's own formula cells
                        area = resolve(text, sheet, book)
                        if area.sheet == sheet and area.r1 <= r2 and area.r2 >= r1 and any(
                                area.c1 <= c <= area.c2 and book.cells.get((sheet, r, c)) is not None
                                and book.cells[(sheet, r, c)].formula is not None
                                for c in cols for r in range(max(area.r1, r1), min(area.r2, r2) + 1)):
                            why = f"reads {text}, which includes formula cells of its own rows"
            except Unsupported as err:
                rejected.append(f"{_place(sheet, col, r1, r2)} stays in cell mode: {err}")
                continue
            if why:
                rejected.append(f"{_place(sheet, col, r1, r2)} stays in cell mode: {why}")
                continue
            template[col] = map_refs(ast, lambda n: RefNode(ROWREF.format(classify(n.text, r1)[1]))
                                     if classify(n.text, r1)[0] == "same_row" else n)
            reads[col] = found
            keep.append(col)
        _, stuck = column_order(keep, reads)
        for col in stuck:
            rejected.append(f"{_place(sheet, col, r1, r2)} stays in cell mode: its row-wise reads form a cycle")
        keep = [c for c in keep if c not in stuck]
        if keep:
            regions.append(Region(sheet, r1, r2, sorted(keep), {c: template[c] for c in keep},
                                  {c: reads[c] for c in keep}))
    return RegionScan(regions, rejected)


def order_with_regions(book: Book, regions: list[Region]):
    """Dependency order with each region as one node (implemented with the graph in the next task)."""
    raise NotImplementedError
```

- [ ] **Step 4: Run, expect pass**

Run: `.venv/Scripts/python.exe -m pytest tests/test_regions.py -q`
Expected: `8 passed`.

- [ ] **Step 5: Measure the key cost**

Run: `.venv/Scripts/python.exe -c "import time; from unexcel.r1c1 import r1c1_key; t=time.perf_counter(); [r1c1_key(f'IF(B{r}>100,B{r},0)', r, 4) for r in range(1, 100001)]; print(round(time.perf_counter()-t, 2), 's per 100k formulas')"`
Expected: about 5 s per 100k formulas (measured 4.7 s). Record the number in the commit body. If it is above 15 s, stop and report: the key function needs a faster implementation before 200k workbooks are practical.

- [ ] **Step 6: Commit**

```bash
git add src/unexcel/regions.py tests/test_regions.py
git commit -m "feat: find row-wise regions (pasted-down formula columns) and say why others stay in cell mode"
```

---

### Task 3: Dependency graph with groups

**Files:**
- Modify: `src/unexcel/graph.py` (`formula_order` gains `groups`; new `order_with_groups`)
- Modify: `src/unexcel/regions.py` (`order_with_regions` implemented)
- Test: `tests/test_graph.py` (append)

**Interfaces:**
- Consumes: `Region.keys()` (Task 2).
- Produces:
  - `formula_order(book, groups=()) -> (order, unresolved)`: `groups` is a list of key lists. Each group is one node: all its cells are released together and appear consecutively in `order` (sorted row-major). A read of another member of the same group is not a wait; a cell reading itself still is a cycle.
  - `order_with_groups(book, groups) -> (order, unresolved, kept_groups)`: orders with the groups; if that reports a `CYCLE` for any cell that the ungrouped ordering does not, the groups containing such cells are dissolved and ordering is repeated. Terminates because the set only shrinks.
  - `order_with_regions(book, regions) -> (order, unresolved, kept_regions)`.

- [ ] **Step 1: Append the failing tests** to `tests/test_graph.py`

```python
from unexcel.graph import order_with_groups


def _col(c, rows):
    return [("S", r, c) for r in rows]


def test_a_group_is_released_together_and_stays_consecutive():
    spec = {"Z1": ("=A1*2", 0.0)}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=B{r}+Z1", 0.0)  # reads a cell outside the group
    spec["A1"] = 1.0
    book = book_of({"S": spec})
    order, bad = formula_order(book, groups=[_col(2, range(2, 6)) + _col(3, range(2, 6))])
    assert not bad
    assert order[0] == ("S", 1, 26)  # Z1 first: the group waits for it
    assert order[1:] == sorted(_col(2, range(2, 6)) + _col(3, range(2, 6)))  # then the group, row-major


def test_reading_another_member_of_the_group_is_not_a_wait_but_self_reading_is_a_cycle():
    spec = {f"A{r}": float(r) for r in range(2, 5)}
    spec.update({f"B{r}": (f"=A{r}*2", 0.0) for r in range(2, 5)})
    spec.update({f"C{r}": (f"=B{r}+1", 0.0) for r in range(2, 5)})
    order, bad = formula_order(book_of({"S": spec}), groups=[_col(2, range(2, 5)) + _col(3, range(2, 5))])
    assert len(order) == 6 and not bad
    order, bad = formula_order(book_of({"S": {"A1": ("=A1+1", 0.0)}}), groups=[[("S", 1, 1)]])
    assert order == [] and set(bad) == {("S", 1, 1)}


def test_a_group_that_creates_a_cycle_the_cells_do_not_have_is_dissolved():  # Review Focus 3
    spec = {}
    for r in range(2, 5):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=Z1+B{r}", 0.0)
    spec["Z1"] = ("=B3+1", 0.0)  # B3 -> Z1 -> C2..C4 is fine for cells, a cycle if B and C are one node
    book = book_of({"S": spec})
    group = _col(2, range(2, 5)) + _col(3, range(2, 5))
    grouped_order, grouped_bad = formula_order(book, groups=[group])
    assert grouped_bad  # the false cycle
    order, bad, kept = order_with_groups(book, [group])
    assert bad == {} and kept == [] and len(order) == 7
    position = {k: i for i, k in enumerate(order)}
    assert position[("S", 3, 2)] < position[("S", 1, 26)] < position[("S", 2, 3)]


def test_a_real_cycle_is_still_reported_and_the_group_is_kept_out_of_it():
    spec = {"A1": ("=B1", 0.0), "B1": ("=A1", 0.0)}
    for r in range(2, 5):
        spec[f"C{r}"] = (f"={r}+1", 0.0)
    order, bad, kept = order_with_groups(book_of({"S": spec}), [])
    assert set(bad) == {("S", 1, 1), ("S", 1, 2)}


def test_groups_agree_with_cell_order_on_random_books():
    rng = random.Random(11)
    kept_total = dissolved_total = 0
    for i in range(600):
        book, groups = _random_book_and_groups(rng)
        order, bad, kept = order_with_groups(book, groups)
        base_order, base_bad = formula_order(book)
        assert set(bad) == set(base_bad), i  # the same cycles as cell mode
        formulas = {k for k, c in book.cells.items() if c.formula is not None}
        assert set(order) | set(bad) == formulas and len(order) == len(set(order))
        position = {k: j for j, k in enumerate(order)}
        for key in order:
            for text in refs(parse(book.cells[key].formula)):
                for dep in resolve(text, "S", book).present(book):
                    if dep in formulas and dep in position and dep != key:
                        assert any(key in g and dep in g for g in kept) or position[dep] < position[key], (i, key, dep)
        for g in kept:
            idx = sorted(position[k] for k in g if k in position)
            assert not idx or idx == list(range(idx[0], idx[0] + len(idx))), i
        kept_total += len(kept)
        dissolved_total += len(groups) - len(kept)
    assert kept_total > 20 and dissolved_total > 20  # both paths were exercised (about 65 and 210 expected)
    for _ in range(200):
        book, _unused = _random_book_and_groups(rng)
        assert formula_order(book, groups=[]) == formula_order(book)


def _internally_ok(book, group):
    """Like a real region: no range covers a member, and the reads among members are acyclic."""
    members = set(group)
    edges = {k: set() for k in group}
    for k in group:
        for text in refs(parse(book.cells[k].formula)):
            area = resolve(text, "S", book)
            if area.is_single:
                if area.key == k:
                    return False
                if area.key in members:
                    edges[k].add(area.key)
            elif any(m in members for m in area.present(book)):
                return False
    done = set()
    while len(done) < len(group):
        ready = [k for k in group if k not in done and not (edges[k] - done)]
        if not ready:
            return False
        done.update(ready)
    return True


def _random_book_and_groups(rng):
    n = rng.randint(4, 14)
    spec = {}
    for r in range(1, n + 1):
        for c in "ABCD":
            if rng.random() < 0.25:
                spec[f"{c}{r}"] = float(r)
            else:
                a, b = rng.randint(1, n), rng.randint(1, n)
                c2, c3 = rng.choice("ABCD"), rng.choice("ABCD")
                spec[f"{c}{r}"] = (rng.choice([f"={c2}{r}+1", f"={c2}{r}*{c3}{a}",
                                               f"=SUM({c2}{min(a, b)}:{c3}{max(a, b)})", f"={c2}{a}+1"]), 0.0)
    book = book_of({"S": spec})
    formulas = sorted(k for k, c in book.cells.items() if c.formula is not None)
    groups, used = [], set()
    for _ in range(rng.randint(0, 3)):
        col = rng.randint(1, 4)
        rows = sorted(r for (_s, r, c) in formulas if c == col)
        if len(rows) < 3:
            continue
        a = rng.randint(0, len(rows) - 3)
        b = rng.randint(a + 2, len(rows) - 1)
        run = [("S", r, col) for r in rows[a:b + 1] if ("S", r, col) not in used]
        if len(run) >= 2:
            groups.append(run)
            used.update(run)
            col2 = rng.randint(1, 4)
            if col2 != col and rng.random() < 0.5:
                run2 = [("S", r, col2) for r in rows[a:b + 1] if ("S", r, col2) in set(formulas) and ("S", r, col2) not in used]
                if len(run2) >= 2:
                    groups[-1] = groups[-1] + run2
                    used.update(run2)
    return book, [g for g in groups if _internally_ok(book, g)]
```

(The file's existing imports already provide `book_of`, `formula_order`, `random`, `parse`, `refs`, `resolve`.)

- [ ] **Step 2: Run, expect failure**

Run: `.venv/Scripts/python.exe -m pytest tests/test_graph.py -q`
Expected: the new tests FAIL (`formula_order() got an unexpected keyword argument 'groups'`, `ImportError: cannot import name 'order_with_groups'`).

- [ ] **Step 3: Replace `src/unexcel/graph.py`** (keeps the Plan 2c algorithm; every cell is a node unless it is in a group)

```python
"""Order formula cells so every cell comes after the cells it reads.

Ranges are not expanded into cell sets (that is quadratic for a running SUM($B$1:Bn)); they are O(n log n) when the
cells they read are ordered in ascending order (see the ponytail comment in the main loop for the bottom-up case). Per
column, a range is an interval of that column's formula cells, and a Fenwick tree of "not yet ordered" counts says
whether the interval is empty and which cell to wait for next. Cells are released level by level: a node's level is one
more than the last node it waited for, so every cell that needs only inputs comes first, then the cells that need those.

A group of cells (a region) is one node: it waits for everything its members read from outside the group, and when it is
released all its members are ordered at once, consecutively, row-major.
"""
from __future__ import annotations

import heapq
from bisect import bisect_left, bisect_right
from collections import defaultdict

from unexcel.fenwick import Fenwick
from unexcel.model import Book, Key, Unsupported
from unexcel.parser import parse, refs
from unexcel.refs import resolve

CYCLE = "circular reference (or depends on one)"

Column = tuple[str, int]  # (sheet, col)


def formula_order(book: Book, groups=()) -> tuple[list[Key], dict[Key, str]]:
    formulas = {k for k, c in book.cells.items() if c.formula is not None}
    grouped = {k for g in groups for k in g}
    members = [sorted(g) for g in groups] + [[k] for k in sorted(formulas - grouped)]
    node_of = {k: i for i, m in enumerate(members) for k in m}
    frows: dict[Column, list[int]] = defaultdict(list)  # rows holding a formula, per column, sorted
    for sheet, row, col in sorted(formulas):
        frows[(sheet, col)].append(row)
    fcols: dict[str, list[int]] = defaultdict(list)  # columns holding a formula, per sheet, sorted
    for sheet, col in sorted(frows):
        fcols[sheet].append(col)
    unordered = {column: Fenwick([1] * len(rows)) for column, rows in frows.items()}

    singles: dict[Key, set[Key]] = {}  # cell -> formula cells it reads one by one
    slices: dict[Key, list[tuple[Column, int, int]]] = {}  # cell -> (column, first, past-last) index into frows
    for key in formulas:
        cell = book.cells[key]
        one: set[Key] = set()
        many: list[tuple[Column, int, int]] = []
        try:
            for text in refs(parse(cell.formula)):
                area = resolve(text, cell.sheet, book)
                if area.is_single:
                    # a read of another member of the same group is not a wait; a cell reading itself still is
                    if area.key in formulas and (node_of[area.key] != node_of[key] or area.key == key):
                        one.add(area.key)
                    continue
                cols = fcols.get(area.sheet, [])
                for col in cols[bisect_left(cols, area.c1):bisect_right(cols, area.c2)]:
                    rows = frows[(area.sheet, col)]
                    i, j = bisect_left(rows, area.r1), bisect_right(rows, area.r2)
                    if j > i:
                        many.append(((area.sheet, col), i, j))
        except (Unsupported, RecursionError):
            pass  # the compiler reports the real reason for this cell
        singles[key], slices[key] = one, many

    users = defaultdict(list)  # cell -> nodes waiting for it, once per edge
    for key, reads in singles.items():
        for dep in reads:
            users[dep].append(node_of[key])
    blocked_on: dict[tuple[Column, int], list] = defaultdict(list)  # (column, index) -> slices waiting for that cell
    waiting = [0] * len(members)
    for key in formulas:
        node = node_of[key]
        waiting[node] += len(singles[key]) + len(slices[key])
        for column, i, j in slices[key]:
            blocked_on[(column, j - 1)].append((node, column, i, j))
    level = [0] * len(members)
    ready = [(0, members[n][0], n) for n in range(len(members)) if waiting[n] == 0]
    heapq.heapify(ready)
    order = []

    def release(user: int, after: int) -> None:
        level[user] = max(level[user], level[after] + 1)
        waiting[user] -= 1
        if waiting[user] == 0:
            heapq.heappush(ready, (level[user], members[user][0], user))

    while ready:
        _, _, node = heapq.heappop(ready)
        for key in members[node]:
            order.append(key)
            column = (key[0], key[2])
            pos = bisect_left(frows[column], key[1])
            unordered[column].add(pos, -1)
            for user in users[key]:
                release(user, node)
            for user, ucol, i, j in blocked_on.pop((column, pos), ()):
                fw = unordered[ucol]
                if fw.range(i, j) == 0:
                    release(user, node)
                else:  # wait for the last cell of the range that is still unordered
                    # ponytail: bottom-up chain (B_r = B_{r+1}+1 with C_r = SUM($B$1:B_r)) moves every waiting slice to
                    # the previous cell one step at a time, so ordering is O(n^2 log n) there (4000 rows: 28.8 s vs 7.9 s
                    # for the old graph). Upgrade: register each slice on canonical Fenwick/segment nodes with counters,
                    # or move whole waiter lists in one step.
                    blocked_on[(ucol, fw.kth(fw.prefix(j)))].append((user, ucol, i, j))
    return order, {k: CYCLE for k in sorted(formulas - set(order))}


def order_with_groups(book: Book, groups: list[list[Key]]):
    """Order with groups as nodes. A group that creates a cycle the cells do not have is dissolved and ordering
    repeated; the set of groups only shrinks, so this ends. Returns (order, unresolved, kept groups)."""
    groups = [list(g) for g in groups]
    base_bad = None
    while True:
        order, bad = formula_order(book, groups)
        if not bad or not groups:
            return order, bad, groups
        if base_bad is None:
            base_bad = set(formula_order(book)[1])
        extra = set(bad) - base_bad
        if not extra:
            return order, bad, groups
        groups = [g for g in groups if not (set(g) & extra)]
```

- [ ] **Step 4: Implement `order_with_regions`** in `src/unexcel/regions.py`: replace the stub with

```python
def order_with_regions(book: Book, regions: list[Region]):
    """Dependency order with each region as one node: (order, unresolved, regions that survived)."""
    from unexcel.graph import order_with_groups

    order, bad, kept = order_with_groups(book, [r.keys() for r in regions])
    kept_sets = [set(g) for g in kept]
    return order, bad, [r for r in regions if set(r.keys()) in kept_sets]
```

- [ ] **Step 5: Run the graph tests and the whole suite**

Run: `.venv/Scripts/python.exe -m pytest tests/test_graph.py tests/test_regions.py -q` then `PATH="$PWD/.venv/Scripts:$PATH" .venv/Scripts/python.exe -m pytest tests -q --ignore=tests/test_perf.py`
Expected: all pass. The earlier graph tests (order, cycles, speed) must pass unchanged: with no groups the algorithm is the Plan 2c one.

- [ ] **Step 6: Commit**

```bash
git add src/unexcel/graph.py src/unexcel/regions.py tests/test_graph.py
git commit -m "feat: dependency graph orders groups of cells as one node and dissolves false cycles"
```

---

### Task 4: Compiler hooks (shared row builder, same-row references, batch dispatch)

**Files:**
- Modify: `src/unexcel/compiler.py`
- Test: `tests/test_compiler.py` (append), `tests/test_batches.py` (append)

**Interfaces:**
- Produces:
  - `cell_row(typ, v, e, t, k) -> (kind_sql, num_sql, txt_sql)`: the kind/num/txt SQL of one result row from SQL for the value `v`, error `e`, text `t` and kind `k`. `render_batch` uses it; its output must stay byte-identical. When `e == NO_ERR` the CASE wrappers are dropped.
  - `Compiler.rowrefs: dict[str, Expr] | None = None`: when set, a `RefNode` whose text is a key of it compiles to that `Expr` (used by the region compiler for `__ROWREF_n__` placeholders). A function that resolves such a placeholder itself fails with `Unsupported`, which keeps the column in cell mode.
  - `_finish(comp, batches, batch)`: appends a finished batch (a list of `Step`, or a region batch) and registers its cells as done in the compiler and the range index.
  - `batch_rows(batch, prev) -> str`: SQL yielding only the batch's new rows (a cell batch or a region batch); used by the verifier and emitter.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_compiler.py`:

```python
from unexcel.compiler import cell_row
from unexcel.expr import Expr
from unexcel.parser import Binary, Num, RefNode


def test_cell_row_drops_the_error_cases_when_the_cell_cannot_fail():
    assert cell_row("num", "x", "NULL", "NULL", "NULL") == ("'n'", "CAST(x AS DOUBLE)", "NULL")
    assert cell_row("txt", "x", "NULL", "NULL", "NULL") == ("'s'", "CAST(NULL AS DOUBLE)", "CAST(x AS VARCHAR)")
    assert cell_row("bool", "x", "NULL", "NULL", "NULL") == ("'b'", "CAST(x AS DOUBLE)", "NULL")
    kind, num_sql, txt_sql = cell_row("num", "x", "e", "NULL", "NULL")
    assert "CASE WHEN e IS NULL" in kind and "CASE WHEN e IS NULL" in num_sql and txt_sql == "e"


def test_rowrefs_compile_placeholders_to_the_given_expressions():
    comp = Compiler(book_of({"S": {"A1": 1.0}}))
    comp.rowrefs = {"__ROWREF_1__": Expr("amount", "num")}
    e = comp.expr(Binary("*", RefNode("__ROWREF_1__"), Num(2.0)), "S")
    assert e.sql == "(amount * CAST(2.0 AS DOUBLE))" and e.type == "num"
    comp.rowrefs = None
    with pytest.raises(Unsupported):
        comp.expr(RefNode("__ROWREF_1__"), "S")  # without the hook the placeholder is an unknown name
```

Append to `tests/test_batches.py`:

```python
from unexcel.compiler import batch_rows


def test_batch_rows_of_a_cell_batch_is_its_own_rows_sql():
    book = book_of({"S": {"A1": 2, "A2": ("=A1*2", 4)}})
    steps = compile_book(book).batches[0]
    assert batch_rows(steps, "cells") == render_batch(steps, "cells", own_only=True)
```

- [ ] **Step 2: Run, expect failure**

Run: `.venv/Scripts/python.exe -m pytest tests/test_compiler.py tests/test_batches.py -q`
Expected: FAIL with `ImportError: cannot import name 'cell_row'` (and `batch_rows`).

- [ ] **Step 3: Edit `src/unexcel/compiler.py`**

(a) Add `cell_row` above `render_batch` and use it inside `render_batch`:

```python
def cell_row(typ: str, v: str, e: str, t: str, k: str) -> tuple[str, str, str]:
    """The kind, num and txt SQL of one result row, from SQL for its value v, error e, text t and kind k."""
    clean = e == NO_ERR  # the cell cannot hold an error: no CASE needed
    if typ == "any":  # the kind is only known when the SQL runs
        if clean:
            return k, f"CASE WHEN {k} <> 's' THEN CAST({v} AS DOUBLE) END", f"CASE WHEN {k} = 's' THEN {t} END"
        return (f"CASE WHEN {e} IS NULL THEN {k} ELSE 'e' END",
                f"CASE WHEN {e} IS NULL AND {k} <> 's' THEN CAST({v} AS DOUBLE) END",
                f"CASE WHEN {e} IS NOT NULL THEN {e} WHEN {k} = 's' THEN {t} END")
    if clean:
        num_sql = "CAST(NULL AS DOUBLE)" if typ in ("txt", "err") else f"CAST({v} AS DOUBLE)"
        return f"'{KIND[typ]}'", num_sql, (f"CAST({v} AS VARCHAR)" if typ == "txt" else "NULL")
    num_sql = ("CAST(NULL AS DOUBLE)" if typ in ("txt", "err")
               else f"CASE WHEN {e} IS NULL THEN CAST({v} AS DOUBLE) END")
    txt_sql = f"COALESCE({e}, CAST({v} AS VARCHAR))" if typ == "txt" else e
    return f"CASE WHEN {e} IS NULL THEN '{KIND[typ]}' ELSE 'e' END", num_sql, txt_sql
```

and in `render_batch` replace the block from `if typ == "any":` through the `txt_sql = ...` assignments with:

```python
        kind, num_sql, txt_sql = cell_row(typ, v, e, t, k)
```
(keep `v, e, t, k = (f"(SELECT {c} FROM {s.cte})" for c in "vetk")` above it and the `rows.append(...)` below it unchanged; `e` here is always a subquery string, so the output is byte-identical to before).

(b) In `Compiler.__init__` add `self.rowrefs: dict[str, Expr] | None = None`. In `_expr`, as the first lines of the `RefNode` branch:

```python
        if isinstance(node, RefNode):
            if self.rowrefs is not None and node.text in self.rowrefs:
                return self.rowrefs[node.text]
            area = resolve(node.text, sheet, self.book)
```

(c) Replace `_close` with:

```python
def _finish(comp: Compiler, batches: list, batch) -> None:
    """A batch is finished: later cells read its results from the table. `batch` is a list of Steps or a region batch."""
    batches.append(batch)
    comp.done.update((s.key, (s.expr.type, s.expr.err != NO_ERR)) for s in batch)
    for s in batch:
        comp.index.complete(s.key, KIND[s.expr.type], s.expr.err != NO_ERR)


def _close(comp: Compiler, batches: list[list[Step]], current: list[Step]) -> None:
    """Seal the open batch: later cells read its results from the table."""
    if current:
        _finish(comp, batches, list(current))
        comp.open.clear()
        current.clear()
```

(d) Add `batch_rows` after `render_batch`:

```python
def batch_rows(batch, prev: str) -> str:
    """SQL yielding only the new rows of a batch: a list of Steps (cell batch) or a region batch."""
    if hasattr(batch, "render_rows"):
        return batch.render_rows(prev)
    return render_batch(batch, prev, own_only=True)
```

- [ ] **Step 4: Run the whole suite**

Run: `PATH="$PWD/.venv/Scripts:$PATH" .venv/Scripts/python.exe -m pytest tests -q --ignore=tests/test_perf.py`
Expected: all pass (byte-identical `render_batch` keeps every emitter and verifier test green).

- [ ] **Step 5: Commit**

```bash
git add -u
git commit -m "refactor: shared cell_row builder, same-row reference hook and batch dispatch in the compiler"
```

---

### Task 5: Compile a region

**Files:**
- Create: `src/unexcel/region.py`
- Test: `tests/test_region.py`

**Interfaces:**
- Consumes: `Region`, `column_order` (Task 2); `Compiler.rowrefs`, `cell_row`, `Step` (Task 4); `ROWREF`.
- Produces:
  - `names_for(book, region, cols) -> dict[int, str]`: SQL column names: the unique text above the region (lower-cased, `[^0-9a-z]` to `_`), else `col_<letter>` (also when empty, starting with a digit, duplicated, or equal to `row_num`/`sheet`).
  - `source_info(comp, region, col) -> (type, may_err)` or `Unsupported`: a column read but not computed by the region must be dense over the region's rows and single-kind (all numbers, all text or all booleans, no errors); formula cells must already be in `comp.done` with one type among `num/txt/bool`.
  - `RegionBatch(list)`: the `Step` of every region cell, plus `model: str`, `wide_sql: str`, `rows: list[str]`, `unpivot(source) -> str`, `render_rows(prev) -> str`.
  - `compile_region(comp, region) -> (RegionBatch | None, rejected: dict[int, str])`: compiles the columns it can (per column; a column that fails, or reads a rejected column, goes into `rejected` with its reason); `None` when no column compiled. The compiler must have no open batch.
  - `region_note(region, batch, rejected) -> list[str]`: report lines.

- [ ] **Step 1: Write the failing tests** (`tests/test_region.py`)

```python
import tempfile
from pathlib import Path

import duckdb
import pytest
from xlsxkit import book_of

import unexcel.compiler as compiler
from unexcel.compiler import READ_SEED, Compiler, compile_book, write_seed
from unexcel.model import Unsupported
from unexcel.region import compile_region, names_for, region_note
from unexcel.regions import find_regions
from unexcel.verifier import execute


def small_book():
    spec = {"A1": "amount", "B1": "tax", "C1": "big"}
    for r in range(2, 5):
        spec[f"A{r}"] = float(r * 40)
        spec[f"B{r}"] = (f"=A{r}*1.1", 0.0)
        spec[f"C{r}"] = (f"=IF(B{r}>100,B{r},0)", 0.0)
    return book_of({"S": spec})


GOLDEN_WIDE = """WITH
src AS (
  SELECT row_num,
    MAX(num) FILTER (WHERE col_num = 1) AS amount
  FROM __unexcel_cells__ WHERE sheet = 'S' AND row_num BETWEEN 2 AND 4 AND col_num IN (1) GROUP BY row_num),
tax AS (
  SELECT *, (amount * CAST(1.1 AS DOUBLE)) AS tax
  FROM src),
big AS (
  SELECT *, (CASE WHEN (tax > CAST(100.0 AS DOUBLE)) THEN tax ELSE CAST(0.0 AS DOUBLE) END) AS big
  FROM tax)
SELECT row_num, tax, big
FROM big"""


def test_the_wide_model_reads_as_plain_sql():  # the readability contract, reviewed by the user
    book = small_book()
    batch, rejected = compile_region(Compiler(book), find_regions(book).regions[0])
    assert rejected == {} and batch.wide_sql == GOLDEN_WIDE
    assert batch.model == "region_s_2_4"
    assert batch.unpivot("wide") == (
        "SELECT 'S', row_num, 2, 'n', CAST(tax AS DOUBLE), NULL FROM wide\n"
        "UNION ALL\n"
        "SELECT 'S', row_num, 3, 'n', CAST(big AS DOUBLE), NULL FROM wide")
    assert len(batch) == 6 and batch[0].key == ("S", 2, 2)


def run_region(book, region):
    """The region's rows next to what cell mode computes for the same cells."""
    comp = Compiler(book)
    batch, rejected = compile_region(comp, region)
    con = duckdb.connect()
    with tempfile.TemporaryDirectory() as tmp:
        seed = Path(tmp) / "s.csv"
        write_seed(book, seed)
        con.execute(f"CREATE TABLE cells AS SELECT * FROM {READ_SEED}", [str(seed)])
    if batch is not None:
        con.execute("INSERT INTO cells " + batch.render_rows("cells"))
    got = {(s, r, c): (k, n, t) for s, r, c, k, n, t in con.execute("SELECT * FROM cells").fetchall()}
    return got, execute(book, compile_book(book, regions=False)), batch, rejected


def same(a, b):
    return (a is not None and b is not None and a[0] == b[0] and a[2] == b[2]
            and (a[1] == b[1] or (a[1] is not None and b[1] is not None and abs(a[1] - b[1]) <= 1e-9 * max(1, abs(a[1])))))


def test_a_mixed_region_computes_what_cell_mode_computes():
    spec = {"A1": "amount", "B1": "tax", "C1": "net", "D1": "ratio", "E1": "divisor", "F1": "label", "G1": "flag",
            "H1": "mixed", "I1": "scaled", "J1": "rate", "K1": 0.5}
    for r in range(2, 14):
        spec[f"A{r}"] = float(r * 17 % 53)
        spec[f"E{r}"] = float(r % 3)  # zeros: #DIV/0!
        spec[f"B{r}"] = (f"=A{r}*1.1", 0.0)
        spec[f"C{r}"] = (f"=IF(B{r}>30,B{r},0)", 0.0)
        spec[f"D{r}"] = (f"=A{r}/E{r}", 0.0)
        spec[f"F{r}"] = (f'=IF(A{r}>20,"hi","lo")&"!"', "x")
        spec[f"G{r}"] = (f"=B{r}>30", True)
        spec[f"H{r}"] = (f'=IF(A{r}>25,A{r},"n/a")', 0.0)
        spec[f"I{r}"] = (f"=A{r}*$K$1+ROUND(B{r},1)", 0.0)
    book = book_of({"S": spec})
    region = find_regions(book).regions[0]
    got, cell_mode, batch, rejected = run_region(book, region)
    assert rejected == {} and len(batch) == len(region.keys())
    for key in region.keys():
        assert same(got[key], cell_mode[key]), (key, got[key], cell_mode[key])


def test_an_absolute_reference_into_the_regions_own_rows_is_one_scalar():  # Review Focus 2
    spec = {"A1": "a"}
    for r in range(2, 7):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*$A$3+A$2", 0.0)  # $A$3 and A$2 sit inside rows 2..6 but are not row-wise
    book = book_of({"S": spec})
    got, cell_mode, batch, rejected = run_region(book, find_regions(book).regions[0])
    assert rejected == {}
    assert all(same(got[("S", r, 2)], cell_mode[("S", r, 2)]) for r in range(2, 7))
    assert got[("S", 4, 2)][1] == 4.0 * 3.0 + 2.0


def test_an_input_column_with_a_blank_or_an_error_keeps_the_reading_column_in_cell_mode():  # Review Focus 1
    spec = {}
    for r in range(2, 7):
        spec[f"B{r}"] = float(r)
        spec[f"C{r}"] = (f"=B{r}*2", 0.0)
        spec[f"D{r}"] = (f"=A{r}*2", 0.0)
        spec[f"E{r}"] = (f"=F{r}+1", 0.0)
        spec[f"F{r}"] = float(r)
    spec.update({f"A{r}": float(r) for r in (2, 3, 5, 6)})  # A4 is empty
    spec["F5"] = "#N/A"  # an error cell in column F
    book = book_of({"S": spec})
    region = find_regions(book).regions[0]
    got, cell_mode, batch, rejected = run_region(book, region)
    assert set(rejected) == {4, 5}  # D reads A (blank row), E reads F (an error)
    assert "empty cell" in rejected[4] and "holds e" in rejected[5]
    assert all(same(got[k], cell_mode[k]) for k in batch_keys(batch))  # column C is region-computed and equal


def batch_keys(batch):
    return [s.key for s in batch]


def test_a_relative_cross_sheet_reference_is_not_a_region_but_an_absolute_one_is():  # Review Focus 4
    spec, other = {}, {"A1": 3.0}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*Other!$A$1", 0.0)
        spec[f"C{r}"] = (f"=A{r}*Other!A{r}", 0.0)
    book = book_of({"S": spec, "Other": other})
    scan = find_regions(book)
    assert [(r.cols) for r in scan.regions] == [[2]]
    got, cell_mode, batch, rejected = run_region(book, scan.regions[0])
    assert all(same(got[k], cell_mode[k]) for k in batch_keys(batch))


def test_header_names_fall_back_to_the_column_letter():  # Review Focus 5
    spec = {"A1": "x y", "B1": "x y", "C1": 5.0, "D1": "1st", "E1": "", "F1": "ok"}
    for r in range(2, 6):
        for c in "ABCDEF":
            spec[f"{c}{r}"] = float(r)
    book = book_of({"S": spec})
    from unexcel.regions import Region
    fake = Region("S", 2, 5, [2], {}, {})  # only the row above (row 1) and the sheet matter for naming
    assert names_for(book, fake, [1, 2, 3, 4, 5, 6]) == {
        1: "col_a", 2: "col_b",  # duplicate header text
        3: "col_c",  # a number, not text
        4: "col_d",  # starts with a digit
        5: "col_e",  # empty
        6: "ok"}


def test_a_column_that_cannot_compile_is_rejected_with_a_reason_and_its_dependants_follow():
    spec = {}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=MAX(A{r},1)", 0.0)  # a function that reads a reference itself: not row-wise in slice 1
        spec[f"C{r}"] = (f"=B{r}+1", 0.0)
        spec[f"D{r}"] = (f"=A{r}*2", 0.0)
    book = book_of({"S": spec})
    region = find_regions(book).regions[0]
    batch, rejected = compile_region(Compiler(book), region)
    assert set(rejected) == {2, 3}
    assert "stays in cell mode" in rejected[3]
    assert [s.key[2] for s in batch] == [4] * 4
    notes = region_note(region, batch, rejected)
    assert any("region_s_2_5" in n for n in notes) and any("column B" in n for n in notes)
```

- [ ] **Step 2: Run, expect failure**

Run: `.venv/Scripts/python.exe -m pytest tests/test_region.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'unexcel.region'`.

- [ ] **Step 3: Write `src/unexcel/region.py`** and make one small edit so the tests can run: change the signature in `src/unexcel/compiler.py` to `def compile_book(book: Book, regions: bool = True) -> Compiled:` (the flag is ignored until Task 6; the new tests call `compile_book(book, regions=False)`).

```python
"""Compile a Region into one set-based statement.

A region becomes a wide model (a short pivot of the columns it reads, then one CTE per formula column holding the
readable expression) and a mechanical unpivot of that model into long `cells` rows, which is what later batches read
and what the verifier compares. The expressions come from the same Compiler as cell mode, with a different reader for
same-row references, so every Excel rule carries over unchanged.
"""
from __future__ import annotations

import re

from openpyxl.utils.cell import get_column_letter

from unexcel.compiler import DEEP, Compiler, Step, a1, cell_row
from unexcel.expr import CELLS, NO_ERR, Expr, any_of, finalize, lit_str
from unexcel.model import Book, Unsupported
from unexcel.r1c1 import ROWREF
from unexcel.regions import Region, column_order

WIDE = "__unexcel_wide__"  # stands for the relation the unpivot reads: the wide CTE or the wide dbt model


def names_for(book: Book, region: Region, cols: list[int]) -> dict[int, str]:
    """SQL column names: the unique text above the region, else col_<letter>."""
    raw = {}
    for c in cols:
        cell = book.cells.get((region.sheet, region.r1 - 1, c))
        if cell is not None and cell.formula is None and cell.kind == "s":
            raw[c] = re.sub(r"[^0-9a-zA-Z]+", "_", str(cell.value)).strip("_").lower()
    out = {}
    for c in cols:
        name = raw.get(c, "")
        dup = sum(1 for x in raw.values() if x == name) > 1
        bad = not name or name[0].isdigit() or dup or name in ("row_num", "sheet", "src")
        out[c] = f"col_{get_column_letter(c).lower()}" if bad else name
    return out


def source_info(comp: Compiler, region: Region, col: int) -> tuple[str, bool]:
    """(type, may_err) of a column the region reads but does not compute; Unsupported if it is not plain."""
    types, may = set(), False
    for r in range(region.r1, region.r2 + 1):
        key = (region.sheet, r, col)
        cell = comp.book.cells.get(key)
        if cell is None:
            raise Unsupported(f"column {get_column_letter(col)} has an empty cell in the region's rows")
        if cell.formula is None:
            if cell.kind not in ("n", "s", "b"):
                raise Unsupported(f"column {get_column_letter(col)} holds {cell.kind} cells")
            types.add({"n": "num", "s": "txt", "b": "bool"}[cell.kind])
        else:
            if key not in comp.done:
                raise Unsupported(f"depends on unresolved cell {a1(key)}")
            typ, err = comp.done[key]
            types.add(typ)
            may = may or err
    if len(types) != 1 or next(iter(types)) not in ("num", "txt", "bool"):
        raise Unsupported(f"column {get_column_letter(col)} mixes value kinds")
    return types.pop(), may


class RegionBatch(list):
    """The Step of every cell of a region, plus the SQL that computes them all at once."""

    def __init__(self, steps, model: str, wide_sql: str, rows: list[str]):
        super().__init__(steps)
        self.model, self.wide_sql, self.rows = model, wide_sql, rows

    def unpivot(self, source: str) -> str:
        return "\nUNION ALL\n".join(row.replace(WIDE, source) for row in self.rows)

    def render_rows(self, prev: str) -> str:
        """Only the new long rows, reading the cells table `prev` (what the verifier appends)."""
        return (f"WITH wide AS (\n{self.wide_sql}\n)\n" + self.unpivot("wide")).replace(CELLS, prev)


def model_name(region: Region) -> str:
    slug = re.sub(r"[^0-9a-zA-Z]+", "_", region.sheet).strip("_").lower() or "sheet"
    return f"region_{slug}_{region.r1}_{region.r2}"


def compile_region(comp: Compiler, region: Region):
    """-> (RegionBatch or None, {column: reason it stays in cell mode}). The compiler must have no open batch."""
    book = comp.book
    letter = get_column_letter
    sources = sorted({s for c in region.cols for s in region.reads[c]} - set(region.cols))
    names = names_for(book, region, sorted(set(region.cols) | set(sources)))
    bad_source, src_types = {}, {}
    for s in sources:
        try:
            src_types[s] = source_info(comp, region, s)
        except Unsupported as err:
            bad_source[s] = str(err)
    rowrefs = {ROWREF.format(s): Expr(names[s], t, f"{names[s]}__err" if may else NO_ERR)
               for s, (t, may) in src_types.items()}
    rejected: dict[int, str] = {}
    layers, out_exprs, prev = [], {}, "src"
    comp.rowrefs = rowrefs
    try:
        for c in column_order(region.cols, region.reads)[0]:
            blocked = sorted(region.reads[c] & (set(rejected) | set(bad_source)))
            if blocked:
                b = blocked[0]
                rejected[c] = bad_source.get(b) or f"reads column {letter(b)}, which stays in cell mode"
                continue
            try:
                e = finalize(comp.expr(region.template[c], region.sheet))
            except Unsupported as err:
                rejected[c] = str(err)
                continue
            except RecursionError:
                rejected[c] = DEEP
                continue
            n = names[c]
            extra = [f"{e.sql} AS {n}"]
            if e.err != NO_ERR:
                extra.append(f"CAST({e.err} AS VARCHAR) AS {n}__err")
            if e.type == "any":
                extra += [f"CAST({e.txt} AS VARCHAR) AS {n}__t", f"CAST({e.kind} AS VARCHAR) AS {n}__k"]
            layers.append(f"{n} AS (\n  SELECT *, " + ", ".join(extra) + f"\n  FROM {prev})")
            prev = n
            out_exprs[c] = e
            err = f"{n}__err" if e.err != NO_ERR else NO_ERR
            if e.type == "any":
                rowrefs[ROWREF.format(c)] = any_of(f"{n}__k", n, f"{n}__t", err)
            elif e.type == "err":
                rowrefs[ROWREF.format(c)] = Expr("NULL", "err", err)
            else:
                rowrefs[ROWREF.format(c)] = Expr(n, e.type, err)
    finally:
        comp.rowrefs = None
    if not out_exprs:
        return None, rejected

    used = sorted({s for c in out_exprs for s in region.reads[c] if s in src_types})
    sheet_lit = lit_str(region.sheet).sql
    if used:
        pivot = []
        for s in used:
            typ, may = src_types[s]
            agg = {"num": "MAX(num)", "txt": "MAX(COALESCE(txt, ''))", "bool": "MAX(num) <> 0"}[typ]
            pivot.append(f"{agg} FILTER (WHERE col_num = {s}) AS {names[s]}")
            if may:
                pivot.append(f"MAX(CASE WHEN kind = 'e' THEN txt END) FILTER (WHERE col_num = {s}) AS {names[s]}__err")
        src = ("src AS (\n  SELECT row_num,\n    " + ",\n    ".join(pivot) +
               f"\n  FROM {CELLS} WHERE sheet = {sheet_lit} AND row_num BETWEEN {region.r1} AND {region.r2} "
               f"AND col_num IN ({', '.join(str(s) for s in used)}) GROUP BY row_num)")
    else:
        src = f"src AS (SELECT row_num FROM range({region.r1}, {region.r2 + 1}) AS t(row_num))"
    cols = [c for c in region.cols if c in out_exprs]
    select = ["row_num"]
    rows = []
    for c in cols:
        e, n = out_exprs[c], names[c]
        select.append(n)
        if e.err != NO_ERR:
            select.append(f"{n}__err")
        if e.type == "any":
            select += [f"{n}__t", f"{n}__k"]
        v = n if e.type != "err" else "CAST(NULL AS DOUBLE)"
        kind, num_sql, txt_sql = cell_row(e.type, v, f"{n}__err" if e.err != NO_ERR else NO_ERR,
                                          f"{n}__t" if e.type == "any" else "NULL",
                                          f"{n}__k" if e.type == "any" else "NULL")
        rows.append(f"SELECT {sheet_lit}, row_num, {c}, {kind}, {num_sql}, {txt_sql} FROM {WIDE}")
    wide = "WITH\n" + ",\n".join([src] + layers) + f"\nSELECT {', '.join(select)}\nFROM {prev}"
    steps = [Step((region.sheet, r, c), out_exprs[c], "") for c in cols for r in range(region.r1, region.r2 + 1)]
    return RegionBatch(steps, model_name(region), wide, rows), rejected


def region_note(region: Region, batch, rejected: dict[int, str]) -> list[str]:
    """Report lines: what became a model and what stayed in cell mode, with the reasons."""
    letters = lambda cols: ", ".join(get_column_letter(c) for c in cols)  # noqa: E731
    lines = []
    if batch is not None:
        done = [c for c in region.cols if c not in rejected]
        lines.append(f"{region.sheet} rows {region.r1}-{region.r2}, columns {letters(done)} -> model {batch.model}")
    for c in sorted(rejected):
        lines.append(f"{region.sheet} column {get_column_letter(c)} rows {region.r1}-{region.r2} stays in cell mode: "
                     f"{rejected[c]}")
    return lines
```

Notes for the implementer: `DEEP` and `a1` must be importable from `unexcel.compiler` (`DEEP` is defined there; `a1` is imported there from `unexcel.model`). `Expr.err` strings for a column that reads another column's error use `NO_ERR` when the source cannot hold one, so a plain `A * 1.1` stays plain SQL.

- [ ] **Step 4: Run, expect pass**

Run: `.venv/Scripts/python.exe -m pytest tests/test_region.py -q`
Expected: all pass. If the golden text differs only by whitespace, fix the implementation to match the golden (the golden is the readability contract), not the other way round.

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/region.py tests/test_region.py src/unexcel/compiler.py
git commit -m "feat: compile a region into a wide readable model and a long-row unpivot"
```

---

### Task 6: Use regions in `compile_book`, the verifier, and prove verdicts are unchanged

**Files:**
- Modify: `src/unexcel/compiler.py` (`Compiled`, `compile_book`)
- Modify: `src/unexcel/verifier.py` (`execute`)
- Modify: `tests/test_batches.py` and any other test that asserts cell-mode batch structure (pass `regions=False`)
- Test: `tests/test_regions_differential.py` (create)

**Interfaces:**
- Consumes: everything above.
- Produces:
  - `Compiled.notes: list[str]` (default `[]`): report lines from `region_note` plus the rejected-run notes from `find_regions`.
  - `compile_book(book, regions=True) -> Compiled`. With `regions=False` the result is exactly Plan 2c's.
  - `execute()` appends each batch with `batch_rows`.

- [ ] **Step 1: Write the differential test** (`tests/test_regions_differential.py`)

```python
import random

from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute

TEMPLATES = [
    "={a}{r}*1.1", "=IF({x}{r}>10,{x}{r},0)", "={x}{r}-{y}{r}", "={a}{r}/{b}{r}", '={c}{r}&"!"',
    "={x}{r}>{y}{r}", '=IF({a}{r}>5,{a}{r},"n/a")', "={x}{r}*$H$1", "=SUM($A$2:$A${last})*{x}{r}",
    "=ROUND({x}{r},1)", "=IFERROR({a}{r}/{b}{r},0)", "=SUM($A$2:A{r})", "={x}{p}+1",
    '=IF({x}{r}>3,"big","small")', "={x}{r}&{c}{r}", "=ABS({x}{r})", "=MAX({a}{r},{b}{r})",
]


def make_book(rng):
    n = rng.randint(3, 14)
    last = n + 1
    spec = {"A1": "a", "B1": "b", "C1": "c", "H1": 2.0}
    for r in range(2, last + 1):
        if not (r == 3 and rng.random() < 0.15):  # sometimes a blank in A
            spec[f"A{r}"] = float(rng.randint(0, 9))
        spec[f"B{r}"] = float(rng.randint(0, 9))
        spec[f"C{r}"] = rng.choice(["x", "yy", "3", ""])
    if rng.random() < 0.15:
        spec["B3"] = "#N/A"
    letters = ["A", "B", "C"]  # columns a template may read: inputs, then the formula columns built so far
    for col in "DEFGIJK"[:rng.randint(2, 6)]:
        template = rng.choice(TEMPLATES)
        x, y = rng.choice(letters), rng.choice(letters)  # fixed for the whole column: it is one formula pasted down
        for r in range(2, last + 1):
            spec[f"{col}{r}"] = (template.format(a="A", b="B", c="C", x=x, y=y, r=r, p=r - 1, last=last), 0.0)
        letters.append(col)
    return book_of({"S": spec})


def same(a, b):
    return (a[0] == b[0] and a[2] == b[2]
            and (a[1] == b[1] or (a[1] is not None and b[1] is not None and abs(a[1] - b[1]) <= 1e-9 * max(1, abs(a[1])))))


def test_regions_never_change_a_status_or_a_value():
    rng = random.Random(2026)
    used = 0
    for seed in range(400):
        book = make_book(rng)
        with_regions = compile_book(book, regions=True)
        without = compile_book(book, regions=False)
        assert with_regions.unresolved == without.unresolved, seed
        a, b = execute(book, with_regions), execute(book, without)
        assert a.keys() == b.keys(), seed
        for key in a:
            assert same(a[key], b[key]), (seed, key, a[key], b[key])
        used += any(hasattr(batch, "wide_sql") for batch in with_regions.batches)
    assert used > 100  # regions were really exercised, not silently skipped
```

The generator is seeded (`random.Random(2026)`), so a failure is reproducible: the assertion messages carry the seed. Texts in column C (including `""` and `"3"`) and an occasional `#N/A` in B and a blank in A are deliberate: they push columns onto the rejection paths (blank, error, mixed kinds), where regions must hand the cells back to cell mode with identical results.

- [ ] **Step 2: Run, expect failure**

Run: `.venv/Scripts/python.exe -m pytest tests/test_regions_differential.py -q`
Expected: FAIL (`used > 100` assertion: regions are not used yet), after Task 5's signature stub.

- [ ] **Step 3: Edit `src/unexcel/compiler.py`**

`Compiled` gains `notes: list[str] = field(default_factory=list)` (import `field` from `dataclasses`). Replace `compile_book` with:

```python
def compile_book(book: Book, regions: bool = True) -> Compiled:
    from unexcel.region import compile_region, region_note
    from unexcel.regions import RegionScan, column_order, find_regions, order_with_regions

    scan = find_regions(book) if regions else RegionScan([], [])
    order, unresolved, kept = order_with_regions(book, scan.regions)
    comp = Compiler(book)
    batches: list = []
    current: list[Step] = []
    weight = 0
    first = {r.first_key: r for r in kept}
    inside = {k for r in kept for k in r.keys()} - set(first)
    notes = list(scan.rejected)

    def add(key) -> None:
        nonlocal weight
        cell = book.cells[key]
        if cell.kind == "blank":
            unresolved[key] = NO_VALUE
            return
        if current and (len(current) >= BATCH or weight >= BATCH_WEIGHT):
            _close(comp, batches, current)
            weight = 0
        try:
            try:
                e = comp.compile_cell(cell)
            except NewBatch:
                _close(comp, batches, current)
                weight = 0
                e = comp.compile_cell(cell)  # the open batch is empty now, so this can't raise NewBatch
        except Unsupported as err:
            unresolved[key] = str(err)
            return
        except RecursionError:  # Python's stack, not DuckDB's: DuckDB took 900 nesting levels in tests
            unresolved[key] = DEEP
            return
        step = Step(key, e, f"c{len(current) + 1}")
        comp.open[key] = (step.cte, e.type, e.err != NO_ERR)
        current.append(step)
        weight += step_weight(e)

    for key in order:
        if key in inside:
            continue
        region = first.get(key)
        if region is None:
            add(key)
            continue
        _close(comp, batches, current)  # the region reads finished cells from the table
        weight = 0
        batch, rejected = compile_region(comp, region)
        if batch is not None:
            _finish(comp, batches, batch)
        notes += region_note(region, batch, rejected)
        for col in column_order(region.cols, region.reads)[0]:
            if col in rejected:
                for row in range(region.r1, region.r2 + 1):
                    add((region.sheet, row, col))
    _close(comp, batches, current)
    return Compiled(batches, unresolved, comp.findings, notes)
```

Keep `Compiled.steps` as is (it flattens `batches`; a region batch is a list of Steps).

- [ ] **Step 4: Edit `src/unexcel/verifier.py`**: import `batch_rows` instead of `render_batch`, and in `execute`:

```python
    for batch in compiled.batches:
        con.execute(f"INSERT INTO cells {batch_rows(batch, 'cells')}")
```

- [ ] **Step 5: Run the whole suite, then fix tests that assert cell-mode structure**

Run: `PATH="$PWD/.venv/Scripts:$PATH" .venv/Scripts/python.exe -m pytest tests -q --ignore=tests/test_perf.py`
Expected after steps 3-4: the differential test and the new region tests pass. Some existing tests may fail because a pasted-down column is now a region (known: `test_batches.py::test_thousands_of_cells_no_longer_crash` and `test_running_total...` shapes; also anything counting `len(compiled.batches)`). For each failing test decide: if it asserts the cell-mode batching mechanics, pass `regions=False` to its `compile_book(...)` call and say so in the report; if it asserts a verdict, the code is wrong: STOP and report BLOCKED with the output. Never weaken a verdict assertion.

- [ ] **Step 6: Real Excel check**

Run: `.venv/Scripts/python.exe -m pytest tests/test_excel_saved.py -v`
Expected: 3 passed (verdicts unchanged on the real-Excel workbooks).

- [ ] **Step 7: Commit**

```bash
git add -u && git add tests/test_regions_differential.py
git commit -m "feat: compile row-wise regions as set-based batches; verdicts identical with and without them"
```

---

### Task 7: dbt models for regions

**Files:**
- Modify: `src/unexcel/emitter.py`
- Test: `tests/test_emitter.py` (append)

**Interfaces:**
- Consumes: `RegionBatch.model`, `.wide_sql`, `.unpivot(source)`; `sources_for(k)`, `model_name(k)`, `_union` (Plan 2c).
- Produces: for a region batch number `k`, two models in `models/unexcel_cells/`: `<RegionBatch.model>.sql` (wide, readable) and `cells_000k.sql` (the long unpivot reading it through `ref`), so Plan 2c's `sources_for`/checkpoints are unchanged. `schema.yml` gets one entry per region model. A shared `_escape(sql)` replaces the inline brace escaping.

- [ ] **Step 1: Write the failing tests** (append to `tests/test_emitter.py`)

```python
def region_book(n=8):
    spec = {"A1": "amount", "B1": "tax"}
    for r in range(2, n + 2):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*1.1", 0.0)
        spec[f"C{r}"] = (f"=SUM($A$2:A{r})", 0.0)  # growing range: a cell-mode batch after the region
    return book_of({"S": spec})


def test_a_region_writes_a_wide_model_and_a_long_model(tmp_path):
    book = region_book()
    compiled = compile_book(book)
    assert any(hasattr(b, "wide_sql") for b in compiled.batches)
    write_project(book, compiled, tmp_path, "r")
    names = {p.name for p in (tmp_path / "models" / "unexcel_cells").iterdir()}
    assert "region_s_2_9.sql" in names and "cells_0001.sql" in names
    wide = (tmp_path / "models" / "unexcel_cells" / "region_s_2_9.sql").read_text(encoding="utf-8")
    assert wide.startswith("WITH\nall_cells AS (") and "AS tax" in wide and "{{ ref('input_cells') }}" in wide
    long = (tmp_path / "models" / "unexcel_cells" / "cells_0001.sql").read_text(encoding="utf-8")
    assert "{{ ref('region_s_2_9') }}" in long
    schema = (tmp_path / "models" / "schema.yml").read_text(encoding="utf-8")
    assert "region_s_2_9" in schema


def test_region_models_equal_the_verifier(tmp_path):
    book = region_book()
    compiled = compile_book(book)
    write_project(book, compiled, tmp_path, "r")
    got = {(s, r, c): (k, n, t) for s, r, c, k, n, t in _run_models(tmp_path)}
    expected = execute(book, compiled)
    assert got.keys() == expected.keys()
    for key, value in expected.items():
        assert got[key][0] == value[0] and got[key][2] == value[2]
        assert (got[key][1] is None and value[1] is None) or abs(got[key][1] - value[1]) < 1e-9


def test_region_text_with_braces_survives_jinja(tmp_path):
    jinja2 = pytest.importorskip("jinja2")
    spec = {"A1": "a"}
    for r in range(2, 6):
        spec[f"A{r}"] = "x"
        spec[f"B{r}"] = (f'=A{r}&"{{{{ oops }}}} {{% x %}}"', "x")
    book = book_of({"S": spec})
    compiled = compile_book(book)
    batch = next(b for b in compiled.batches if hasattr(b, "wide_sql"))
    from unexcel.emitter import dbt_region_wide
    rendered = jinja2.Template(dbt_region_wide(batch, ["input_cells"])).render(ref=lambda name: name)
    assert "{{ oops }} {% x %}" in rendered
```

The existing `_run_models` helper must create each region's wide model before the long model that reads it: in its loop over `names`, before creating `cells_k`, create every `region_*` model that `cells_k`'s SQL refs (regex `ref\('(region_\w+)'\)`), reading `models/unexcel_cells/<name>.sql` with the same ref rewriting. Update the helper accordingly (it is test code in the same file).

- [ ] **Step 2: Run, expect failure**

Run: `.venv/Scripts/python.exe -m pytest tests/test_emitter.py -q`
Expected: FAIL (`ImportError: cannot import name 'dbt_region_wide'` / missing files).

- [ ] **Step 3: Edit `src/unexcel/emitter.py`**

Add:

```python
def _escape(sql: str) -> str:
    """Braces from workbook text are escaped so Jinja renders them literally."""
    return re.sub(r"[{}]", lambda m: "{{ '" + m.group() + "' }}", sql)


def dbt_region_wide(batch, sources: list[str]) -> str:
    """The wide, readable model of a region: it reads every earlier cell through the CTE all_cells."""
    body = _escape(batch.wide_sql.replace(CELLS, "all_cells"))
    return body.replace("WITH\n", f"WITH\nall_cells AS (\n{_union(sources)}\n),\n", 1)


def dbt_region_cells(batch) -> str:
    """The long rows of a region: a mechanical unpivot of its wide model."""
    marker = "__unexcel_wide__"
    return _escape(batch.unpivot(marker)).replace(marker, "{{ ref('" + batch.model + "') }}")
```
(import `CELLS` from `unexcel.expr` again), change `dbt_sql` to use `_escape(render_batch(steps, "all_cells", own_only=True))`, and in `write_project`'s batch loop:

```python
    for k, batch in enumerate(compiled.batches, 1):
        if hasattr(batch, "wide_sql"):
            (batches_dir / f"{batch.model}.sql").write_text(dbt_region_wide(batch, sources_for(k)) + "\n", encoding="utf-8")
            (batches_dir / f"{model_name(k)}.sql").write_text(dbt_region_cells(batch) + "\n", encoding="utf-8")
        else:
            (batches_dir / f"{model_name(k)}.sql").write_text(dbt_sql(batch, sources_for(k)) + "\n", encoding="utf-8")
        ...  # checkpoint block unchanged
```
and build `schema.yml` as `SCHEMA` plus, for each region batch:

```python
        f'  - name: {batch.model}\n    description: "Region: rows of one sheet that repeat one formula down their rows, '
        f'computed in one statement. Wide and readable; its cells_NNNN twin holds the same values as long rows."\n'
```

- [ ] **Step 4: Run the emitter tests, then the whole suite with dbt on PATH**

Run: `PATH="$PWD/.venv/Scripts:$PATH" .venv/Scripts/python.exe -m pytest tests -q --ignore=tests/test_perf.py`
Expected: all pass, including `test_dbt_build_passes_on_generated_project` (it must run, not skip). Report whether it ran.

- [ ] **Step 5: Real dbt on a region workbook**

Run: `PATH="$PWD/.venv/Scripts:$PATH" .venv/Scripts/python.exe tests/bench_shapes.py flat 2000 --dbt`
Expected: `dbt_ok=True`, `pass=2000/2000`. Record the line (the Plan 2c layout needed hundreds of models for 20k flat cells; a region needs a handful).

- [ ] **Step 6: Commit**

```bash
git add -u
git commit -m "feat: dbt gets a wide readable model per region and a long twin for the batch chain"
```

---

### Task 8: Report, CLI wiring, close-out

**Files:**
- Modify: `src/unexcel/report.py` (`notes` argument and section), `src/unexcel/cli.py` (pass `compiled.notes`)
- Modify: `PROGRESS.md`, `AGENTS.md` (status block), the spec (status + results)
- Test: `tests/test_report.py` (append), `tests/test_cli.py` (append)

**Interfaces:**
- Produces: `write_report(verdicts, findings, out, source, computed=frozenset(), notes=())`: `REPORT.md` gets a "## Regions" section when `notes` is non-empty; `report.json` gets `"regions": [...]`.

- [ ] **Step 1: Failing tests**

`tests/test_report.py` (append):

```python
def test_regions_section_lists_the_notes(tmp_path):
    from unexcel.report import write_report
    write_report([], [], tmp_path, "b.xlsx", notes=["Data rows 2-9, columns B -> model region_data_2_9",
                                                    "Data column C rows 2-9 stays in cell mode: reads A1"])
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "## Regions" in md and "region_data_2_9" in md and "stays in cell mode" in md
    import json
    assert json.loads((tmp_path / "report.json").read_text(encoding="utf-8"))["regions"][0].startswith("Data rows")


def test_no_regions_section_without_notes(tmp_path):
    from unexcel.report import write_report
    write_report([], [], tmp_path, "b.xlsx")
    assert "## Regions" not in (tmp_path / "REPORT.md").read_text(encoding="utf-8")  # Review Focus 5
```

`tests/test_cli.py` (append): a test that converts a small workbook with a row-wise column through `main([...])` and asserts `REPORT.md` contains `## Regions` and exit code 0 (use the file's existing helper style: `write_xlsx` into `tmp_path`, `main(["convert", str(wb), "--out", str(out)])`).

- [ ] **Step 2: Implement** `write_report(..., notes=())` (add the section after the counts table and `"regions": list(notes)` to the JSON) and in `cli.py`: `write_report(verdicts, compiled.findings, out, args.workbook.name, computed, compiled.notes)`.

- [ ] **Step 3: Run the whole suite and the real-Excel tests**

Run: `PATH="$PWD/.venv/Scripts:$PATH" .venv/Scripts/python.exe -m pytest tests -q --ignore=tests/test_perf.py` and `.venv/Scripts/python.exe -m pytest tests/test_excel_saved.py -v`
Expected: all pass; the 3 saved workbooks still PASS every cell.

- [ ] **Step 4: Benchmarks against the spec's criteria** (one at a time, sequentially; record every line)

```bash
.venv/Scripts/python.exe tests/bench_shapes.py flat 20000
.venv/Scripts/python.exe tests/bench_shapes.py running_input 4000
.venv/Scripts/python.exe tests/bench_shapes.py running_formula 4000
.venv/Scripts/python.exe tests/bench_shapes.py vlookup 1000
PATH="$PWD/.venv/Scripts:$PATH" .venv/Scripts/python.exe tests/bench_shapes.py flat 20000 --dbt
UNEXCEL_PERF=1 .venv/Scripts/python.exe -m pytest tests/test_perf.py -s -q   # alone, last, in the background
```
Compare with the Plan 2c numbers (flat 20k verify 28.3 s / 741 batches; dbt 225.9 s / 213 MB; 200k total 390.3 s) and the spec's criteria 1-5; state MET / NOT MET with the number for each. Shapes that do not become regions must not be slower than after Plan 2c (criterion 4). Any miss is reported, not hidden; targets are not changed.

- [ ] **Step 5: Docs.** `PROGRESS.md`: replace the "Where we stopped" block with the real state and the table of results, move "table mode" from next-plans to done, keep the honest known gaps (what slice 2 should cover: the report's "stays in cell mode" reasons are the map). `AGENTS.md`: replace only the first status block with a pointer to `PROGRESS.md` and today's date. The spec: set the status line to implemented and append the before/after table.

- [ ] **Step 6: Commit**

```bash
git add -u
git commit -m "feat: report lists regions and what stayed in cell mode; Plan 2d results and docs"
```

- [ ] **Step 7: Stop and report.** Do not push or open a PR; ask the user.

---

## Self-review (done while writing)

- **Spec coverage:** detection and admission rules (Tasks 1-2), group ordering and cycle dissolution (Task 3), same compiler with a different reader and the wide/long two-layer output (Tasks 4-5), verifier and dbt (Tasks 6-7), report section (Task 8), differential test and golden readability test (Tasks 5-6), benchmarks and criteria (Task 8).
- **Deviation from the spec text:** the spec says a range over the region's rows is rejected; the plan narrows this to "a range that includes the region's own formula cells" (inputs-only ranges are safe and common, and the compiler's own `ready()` check already protects against a not-yet-computed cell). The spec's sentence about ranges is satisfied for every case that matters for correctness.
- **Placeholders:** none; the one hash-based choice in the differential generator is explained and has an equivalent simpler form.
- **Type consistency:** `Region.keys()/first_key/reads/template`, `RegionScan.regions/rejected`, `column_order`, `order_with_groups`, `order_with_regions`, `cell_row`, `batch_rows`, `RegionBatch.model/wide_sql/rows/unpivot/render_rows`, `compile_region -> (batch|None, rejected)`, `region_note`, `Compiled.notes`, `dbt_region_wide/dbt_region_cells`, and `write_report(..., notes)` are used with the same names and signatures in every task that mentions them.
- **Known risks carried from the spec:** relative/absolute misclassification (mitigated by `classify` and the Task 5 tests plus the differential test); grouped-ordering bugs (the Plan 2c range index still guards every read, plus the dissolve loop and the random-book oracle test); key cost (measured in Task 2 step 5).
