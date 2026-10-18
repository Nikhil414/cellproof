# Plan 2c: Speed Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `unexcel convert` fast on running totals and per-row lookups, and make `--dbt` scale, without changing any verdict.

**Architecture:** Cell mode stays. Four changes: (1) a Fenwick-tree based dependency graph that orders cells level by level and never expands ranges into cell sets; (2) a per-column `RangeIndex` so `Compiler.ready()` and its callers cost O(log n) per range; (3) batches close by SQL weight instead of cell count, and the verifier appends rows in place instead of copying the table; (4) dbt batch models hold only their own rows, with a checkpoint model every 64 batches.

**Tech Stack:** Python 3.10+, DuckDB, openpyxl, pytest. No new runtime dependencies.

**Spec:** `docs/superpowers/specs/2026-10-14-plan-2c-speed-design.md` (read it first: it holds the measured baseline and the success criteria).

## Global Constraints

- Runtime dependencies stay `openpyxl` and `duckdb` only (AGENTS.md section 7).
- A false PASS is the worst bug: when unsure, produce an honest FAIL or UNRESOLVED. The SQL a cell compiles to must not change; verdicts must be identical to `main`.
- No DuckDB exception may escape from a supported function; cell values never leave the machine.
- Write tests first, watch them fail. Console output stays ASCII-only (Windows consoles).
- Measure before claiming speed: every number in a commit message or doc comes from a run on this machine.
- Python is `.venv/Scripts/python.exe` on Windows (`.venv/bin/python` on macOS). Run the suite with `python -m pytest tests -q`.
- Commit messages end with `Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>`.
- Do not push, merge or open a PR without the user's go-ahead.

## Review Focus

Inputs and conditions the spec implies but no obvious task exercises. Each line has its test in the task named.

1. A workbook with no formula cells at all (zero batches) must still verify and emit a dbt project. (Task 6 verifier, Task 8 dbt)
2. One cell whose SQL is heavier than the whole batch budget must still run, alone, and never loop forever. (Task 7)
3. Exactly 64, 65 and 130 batches: the dbt checkpoint boundaries, where an off-by-one would read a cell twice or miss one. (Task 8)
4. A range that contains an unresolved cell (a cycle member, or an unsupported function) must still say "depends on unresolved cell X", and a range that contains the cell itself must still be a cycle. (Task 3, Task 5)
5. A number stored as text inside a range read by SUM must still produce its finding in the report, once. (Task 5)

---

### Task 1: Benchmark tool and baseline

**Files:**
- Create: `tests/bench_shapes.py` (a script, not collected by pytest)

**Interfaces:**
- Produces: `python tests/bench_shapes.py SHAPE ROWS [--dbt] [--batch-weight N]` prints one line `shape n=.. read=.. compile=.. verify=.. batches=.. pass=../..`. `SHAPE` is one of `running_formula`, `running_input`, `vlookup`, `flat`. `make(shape, n)` returns `{sheet: {A1: spec}}` for `xlsxkit.write_xlsx`.

- [ ] **Step 1: Write the tool**

```python
"""Timing tool for the slow sheet shapes. Not a test: run it by hand.

    python tests/bench_shapes.py SHAPE ROWS [--dbt] [--batch-weight N]

SHAPE is running_formula (SUM over a formula column), running_input (SUM over an input column),
vlookup (a VLOOKUP into a 1000-row table on every row) or flat (A*1.1 on every row).
"""
import argparse
import tempfile
import time
from pathlib import Path

from xlsxkit import write_xlsx

import unexcel.compiler as compiler
from unexcel.emitter import run_dbt, write_project
from unexcel.reader import load
from unexcel.verifier import verify


def make(shape: str, n: int) -> dict:
    d = {}
    if shape == "running_formula":
        total = 0.0
        for r in range(1, n + 1):
            b = (r % 97) * 1.5
            total += b
            d[f"A{r}"] = float(r % 97)
            d[f"B{r}"] = (f"=A{r}*1.5", b)
            d[f"C{r}"] = (f"=SUM($B$1:B{r})", total)
        return {"Data": d}
    if shape == "running_input":
        total = 0.0
        for r in range(1, n + 1):
            total += float(r % 97)
            d[f"A{r}"] = float(r % 97)
            d[f"C{r}"] = (f"=SUM($A$1:A{r})", total)
        return {"Data": d}
    if shape == "vlookup":
        table = {}
        for r in range(1, 1001):
            table[f"A{r}"] = float(r)
            table[f"B{r}"] = float(r * 2)
        for r in range(1, n + 1):
            k = float(r % 1000 + 1)
            d[f"A{r}"] = k
            d[f"B{r}"] = (f"=VLOOKUP(A{r},Tbl!$A$1:$B$1000,2,FALSE)", k * 2)
        return {"Data": d, "Tbl": table}
    if shape == "flat":
        for r in range(1, n + 1):
            d[f"A{r}"] = float(r)
            d[f"B{r}"] = (f"=A{r}*1.1", r * 1.1)
        return {"Data": d}
    raise SystemExit(f"unknown shape {shape!r}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("shape")
    ap.add_argument("rows", type=int)
    ap.add_argument("--dbt", action="store_true", help="also write the dbt project, run dbt build, report its size")
    ap.add_argument("--batch-weight", type=int, help="override compiler.BATCH_WEIGHT (once it exists)")
    args = ap.parse_args()
    if args.batch_weight:
        compiler.BATCH_WEIGHT = args.batch_weight
    tmp = Path(tempfile.mkdtemp())
    wb = tmp / "bench.xlsx"
    write_xlsx(wb, make(args.shape, args.rows))
    t0 = time.perf_counter()
    book = load(wb)
    t1 = time.perf_counter()
    compiled = compiler.compile_book(book)
    t2 = time.perf_counter()
    verdicts = verify(book, compiled)
    t3 = time.perf_counter()
    passed = sum(v.status == "PASS" for v in verdicts)
    line = (f"{args.shape} n={args.rows} read={t1 - t0:.1f} compile={t2 - t1:.1f} verify={t3 - t2:.1f} "
            f"batches={len(compiled.batches)} pass={passed}/{len(verdicts)}")
    if args.dbt:
        out = tmp / "out"
        t4 = time.perf_counter()
        write_project(book, compiled, out, "bench")
        ok, _ = run_dbt(out)
        db = out / "unexcel.duckdb"
        line += f" dbt_ok={ok} dbt={time.perf_counter() - t4:.1f}s db_mb={db.stat().st_size / 1e6:.1f}"
    print(line, flush=True)


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Run the baseline on `main` behaviour (this branch has no code change yet) and keep the output**

Run (each line can take minutes; the last two are the slow ones, run them last or in the background):
```bash
python tests/bench_shapes.py running_formula 2000
python tests/bench_shapes.py running_input 4000
python tests/bench_shapes.py flat 20000
python tests/bench_shapes.py vlookup 1000
python tests/bench_shapes.py running_input 2000 --dbt
```
Expected: every line ends `pass=N/N`. Paste the five lines into the commit message body. These are the "before" numbers for Task 9. (Measured on 2026-10-14 on this machine: running_formula 2000 compile 5.2 s verify 12 s, 2001 batches; running_input 4000 verify about 32-59 s; flat 20000 verify about 160 s; vlookup 1000 verify about 85 s.)

- [ ] **Step 3: Start the 200k baseline in the background**

Run: `UNEXCEL_PERF=1 python -m pytest tests/test_perf.py -s` with `run_in_background`. It may exceed its 300 s assertion on this machine (the 255 s figure was measured on a Mac); that is a baseline, not a regression. Note the printed read/compile/verify/emit/total lines in the commit message when it finishes.

- [ ] **Step 4: Commit**

```bash
git add tests/bench_shapes.py
git commit -m "test: benchmark tool for the slow sheet shapes, with the pre-change baseline"
```

---

### Task 2: Fenwick tree

**Files:**
- Create: `src/unexcel/fenwick.py`
- Test: `tests/test_fenwick.py`

**Interfaces:**
- Produces: `Fenwick(values: list[int])` with `add(i, delta)`, `prefix(i)` (sum of positions `0..i-1`), `range(i, j)` (sum of positions `i..j-1`), `kth(k)` (0-based position of the k-th unit, k >= 1, counts non-negative). Values may be arbitrary Python ints for `add`/`prefix`/`range`; `kth` needs non-negative counts.

- [ ] **Step 1: Write the failing test**

```python
import random

from unexcel.fenwick import Fenwick


def test_prefix_range_and_add_match_a_plain_list():
    rng = random.Random(1)
    for _ in range(200):
        n = rng.randint(0, 40)
        values = [rng.randint(0, 5) for _ in range(n)]
        fw = Fenwick(values)
        for _ in range(30):
            if n and rng.random() < 0.4:
                i, d = rng.randrange(n), rng.randint(0, 3)
                values[i] += d
                fw.add(i, d)
            i = rng.randint(0, n)
            j = rng.randint(i, n)
            assert fw.prefix(i) == sum(values[:i])
            assert fw.range(i, j) == sum(values[i:j])


def test_kth_finds_the_position_of_the_kth_unit():
    rng = random.Random(2)
    for _ in range(200):
        n = rng.randint(1, 40)
        values = [rng.randint(0, 3) for _ in range(n)]
        fw = Fenwick(values)
        units = [i for i, v in enumerate(values) for _ in range(v)]
        for k, pos in enumerate(units, 1):
            assert fw.kth(k) == pos


def test_big_packed_integers_work_for_prefix_sums():
    fw = Fenwick([1 << 100, 1 << 100, 0])
    fw.add(2, 1 << 90)
    assert fw.range(0, 3) == (2 << 100) + (1 << 90)
```

- [ ] **Step 2: Run it, expect failure**

Run: `python -m pytest tests/test_fenwick.py -q`
Expected: FAIL with `ModuleNotFoundError: No module named 'unexcel.fenwick'`.

- [ ] **Step 3: Write `src/unexcel/fenwick.py`**

```python
"""Prefix sums over positions 0..n-1 with point updates and k-th lookup, all O(log n)."""
from __future__ import annotations


class Fenwick:
    def __init__(self, values: list[int]):
        n = len(values)
        self.n = n
        t = [0] + list(values)
        for i in range(1, n + 1):  # O(n) build
            j = i + (i & -i)
            if j <= n:
                t[j] += t[i]
        self.t = t
        self.top = 1 << n.bit_length()

    def add(self, i: int, delta: int) -> None:
        """Add delta to position i."""
        i += 1
        t, n = self.t, self.n
        while i <= n:
            t[i] += delta
            i += i & -i

    def prefix(self, i: int) -> int:
        """Sum of positions 0..i-1."""
        s, t = 0, self.t
        while i > 0:
            s += t[i]
            i -= i & -i
        return s

    def range(self, i: int, j: int) -> int:
        """Sum of positions i..j-1."""
        return self.prefix(j) - self.prefix(i)

    def kth(self, k: int) -> int:
        """0-based position of the k-th unit (k >= 1) when every value is a non-negative count."""
        pos, t, n = 0, self.t, self.n
        step = self.top
        while step:
            nxt = pos + step
            if nxt <= n and t[nxt] < k:
                pos = nxt
                k -= t[nxt]
            step >>= 1
        return pos
```

- [ ] **Step 4: Run it, expect pass**

Run: `python -m pytest tests/test_fenwick.py -q`
Expected: `3 passed`.

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/fenwick.py tests/test_fenwick.py
git commit -m "feat: Fenwick tree for range counts in O(log n)"
```

---

### Task 3: Dependency graph without range expansion

**Files:**
- Modify: `src/unexcel/graph.py` (whole file)
- Test: `tests/test_graph.py` (append)

**Interfaces:**
- Consumes: `Fenwick` (Task 2); `unexcel.parser.parse/refs`; `unexcel.refs.resolve`; `Area.is_single`, `Area.key`.
- Produces: `formula_order(book) -> (order, unresolved)` with the same contract as today (every cell after the formula cells it reads; cells on or behind a cycle go to `unresolved` with `CYCLE`), plus: cells are released level by level, so cells that read only inputs come first, then cells that read those, and so on.

- [ ] **Step 1: Append the failing tests to `tests/test_graph.py`**

```python
import random

from unexcel.parser import parse, refs
from unexcel.refs import resolve


def _old_unresolved(book):
    """The previous algorithm, kept as an oracle: expands every range into a set of cells."""
    import heapq
    from collections import defaultdict
    formulas = {k for k, c in book.cells.items() if c.formula is not None}
    deps = {}
    for key in formulas:
        found = set()
        try:
            for text in refs(parse(book.cells[key].formula)):
                found.update(resolve(text, key[0], book).present(book))
        except Exception:
            pass
        deps[key] = found & formulas
    users = defaultdict(set)
    for key, ds in deps.items():
        for d in ds:
            users[d].add(key)
    waiting = {k: len(ds) for k, ds in deps.items()}
    ready = [k for k, n in waiting.items() if n == 0]
    heapq.heapify(ready)
    done = set()
    while ready:
        key = heapq.heappop(ready)
        done.add(key)
        for user in users[key]:
            waiting[user] -= 1
            if waiting[user] == 0:
                heapq.heappush(ready, user)
    return formulas - done


def _random_book(rng):
    n = rng.randint(3, 25)
    spec = {}
    for r in range(1, n + 1):
        for c in "ABC":
            x = rng.random()
            if x < 0.3:
                spec[f"{c}{r}"] = float(r)
            elif x < 0.85:
                a, b = rng.randint(1, n), rng.randint(1, n)
                lo, hi = min(a, b), max(a, b)
                c2, c3 = rng.choice("ABC"), rng.choice("ABC")
                spec[f"{c}{r}"] = (rng.choice([f"={c2}{a}+1", f"=SUM({c2}{lo}:{c3}{hi})", f"=SUM({c2}:{c2})",
                                               f"={c2}{a}*{c3}{b}"]), 0.0)
    return book_of({"S": spec})


def test_order_and_cycles_match_the_old_algorithm_on_random_books():
    rng = random.Random(7)
    for _ in range(1500):
        book = _random_book(rng)
        order, bad = formula_order(book)
        formulas = {k for k, c in book.cells.items() if c.formula is not None}
        assert set(bad) == _old_unresolved(book)
        assert set(order) | set(bad) == formulas and not set(order) & set(bad) and len(set(order)) == len(order)
        position = {k: i for i, k in enumerate(order)}
        for key in order:
            for text in refs(parse(book.cells[key].formula)):
                for dep in resolve(text, "S", book).present(book):
                    if dep in formulas:
                        assert position[dep] < position[key]


def test_a_cell_inside_its_own_range_is_a_cycle():  # Review Focus 4
    book = book_of({"S": {"A1": 1, "A2": ("=SUM(A1:A3)", 0), "A3": 2}})
    assert formula_order(book) == ([], {("S", 2, 1): CYCLE})


def test_a_range_over_a_cycle_member_is_unresolved_too():  # Review Focus 4
    book = book_of({"S": {"A1": ("=A2", 0), "A2": ("=A1", 0), "B1": ("=SUM(A1:A2)", 0)}})
    assert formula_order(book)[1].keys() == {("S", 1, 1), ("S", 2, 1), ("S", 1, 2)}


def test_cells_that_read_only_inputs_come_before_cells_that_read_formulas():
    n = 50  # interleaved in key order today: B1, C1, B2, C2 ...
    spec = {}
    for r in range(1, n + 1):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=SUM($B$1:B{r})", 0.0)
    order, bad = formula_order(book_of({"S": spec}))
    assert not bad
    cols = [key[2] for key in order]
    assert cols == [2] * n + [3] * n


def test_running_total_of_a_long_column_is_fast():
    import time
    n = 20000
    spec = {}
    for r in range(1, n + 1):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=SUM($B$1:B{r})", 0.0)
    book = book_of({"S": spec})
    t0 = time.perf_counter()
    order, bad = formula_order(book)
    assert len(order) == 2 * n and not bad
    assert time.perf_counter() - t0 < 10  # measured 1.6 s; the old graph took minutes and gigabytes here
```

- [ ] **Step 2: Run, expect failures**

Run: `python -m pytest tests/test_graph.py -q`
Expected: `test_cells_that_read_only_inputs_come_before...` FAILS (interleaved order) and `test_running_total_of_a_long_column_is_fast` FAILS or takes very long; the oracle test passes (old graph).

- [ ] **Step 3: Replace `src/unexcel/graph.py`**

```python
"""Order formula cells so every cell comes after the cells it reads.

Ranges are not expanded into cell sets (that is quadratic for a running SUM($B$1:Bn)). Per column, a range is an
interval of that column's formula cells, and a Fenwick tree of "not yet ordered" counts says whether the interval is
empty and which cell to wait for next. Cells are released level by level: a cell's level is one more than the last cell
it waited for, so every cell that needs only inputs comes first, then the cells that need those, and so on.
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


def formula_order(book: Book) -> tuple[list[Key], dict[Key, str]]:
    formulas = {k for k, c in book.cells.items() if c.formula is not None}
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
                    if area.key in formulas:
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

    users = defaultdict(list)
    for key, reads in singles.items():
        for dep in reads:
            users[dep].append(key)
    blocked_on: dict[tuple[Column, int], list] = defaultdict(list)  # (column, index) -> slices waiting for that cell
    waiting = {}
    for key in formulas:
        waiting[key] = len(singles[key]) + len(slices[key])
        for column, i, j in slices[key]:
            blocked_on[(column, j - 1)].append((key, column, i, j))
    level = dict.fromkeys(formulas, 0)
    ready = [(0, k) for k, n in waiting.items() if n == 0]
    heapq.heapify(ready)
    order = []

    def release(user: Key, after: Key) -> None:
        level[user] = max(level[user], level[after] + 1)
        waiting[user] -= 1
        if waiting[user] == 0:
            heapq.heappush(ready, (level[user], user))

    while ready:
        _, key = heapq.heappop(ready)
        order.append(key)
        column = (key[0], key[2])
        pos = bisect_left(frows[column], key[1])
        unordered[column].add(pos, -1)
        for user in users[key]:
            release(user, key)
        for user, ucol, i, j in blocked_on.pop((column, pos), ()):
            fw = unordered[ucol]
            if fw.range(i, j) == 0:
                release(user, key)
            else:  # wait for the last cell of the range that is still unordered
                blocked_on[(ucol, fw.kth(fw.prefix(j)))].append((user, ucol, i, j))
    return order, {k: CYCLE for k in sorted(formulas - set(order))}
```

- [ ] **Step 4: Run the graph tests and the whole suite**

Run: `python -m pytest tests/test_graph.py -q` then `python -m pytest tests -q`
Expected: graph tests all pass. In the full suite `tests/test_batches.py::test_range_over_the_open_batch_closes_it_first` must still pass unchanged (B1,B2 then C1,C2). If any other test fails because it assumed key-order interleaving, fix the test's expectation only if the verdicts are still identical; otherwise stop and report.

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/graph.py tests/test_graph.py
git commit -m "perf: dependency graph without range expansion, released level by level"
```

---

### Task 4: Range index

**Files:**
- Create: `src/unexcel/rangeindex.py`
- Test: `tests/test_rangeindex.py`

**Interfaces:**
- Consumes: `Fenwick` (Task 2); `Book._rows_by_col` (existing cached property: `(sheet, col) -> sorted rows that hold a cell`).
- Produces:
  - `RangeIndex(book)`.
  - `index.complete(key, kind, may_err)`: `kind` is one of `n s b e a` (value kind of a finished formula cell, `a` = decided at run time).
  - `index.summary(sheet, r1, c1, r2, c2) -> Summary(present, formulas, total)`.
  - `index.text_cells(sheet, r1, c1, r2, c2) -> list[Key]`: input text cells, row-major.
  - `field(total, name) -> int` with `name` in `FIELDS = ("n","s","b","e","a","blank","merr","fdone")`.

- [ ] **Step 1: Write the failing test**

```python
import random

from xlsxkit import book_of

from unexcel.rangeindex import FIELDS, RangeIndex, field


def test_summary_matches_a_brute_force_count():
    rng = random.Random(3)
    for it in range(1500):
        n = rng.randint(2, 20)
        spec = {}
        for r in range(1, n + 1):
            for c in "ABCD":
                x = rng.random()
                if x < 0.2:
                    spec[f"{c}{r}"] = float(r)
                elif x < 0.35:
                    spec[f"{c}{r}"] = "txt"
                elif x < 0.4:
                    spec[f"{c}{r}"] = "#N/A"
                elif x < 0.45:
                    spec[f"{c}{r}"] = True
                elif x < 0.5:
                    spec[f"{c}{r}"] = None
                elif x < 0.85:
                    spec[f"{c}{r}"] = ("=1", 1.0)
        book = book_of({"S": spec, "T": {"A1": 1.0}})
        idx = RangeIndex(book)
        done = {}
        for key, cell in book.cells.items():
            if cell.formula is not None and rng.random() < 0.6:
                kind, err = rng.choice([("n", False), ("s", False), ("b", False), ("e", True), ("a", True), ("n", True)])
                idx.complete(key, kind, err)
                done[key] = (kind, err)
        for _ in range(25):
            r1, c1 = rng.randint(1, n), rng.randint(1, 4)
            r2, c2 = rng.randint(r1, n), rng.randint(c1, 4)
            s = idx.summary("S", r1, c1, r2, c2)
            keys = book.present("S", r1, c1, r2, c2)
            want = dict.fromkeys(FIELDS, 0)
            formulas = 0
            for k in keys:
                cell = book.cells[k]
                if cell.formula is None:
                    want[cell.kind] += 1
                    want["merr"] += cell.kind == "e"
                else:
                    formulas += 1
                    if k in done:
                        kind, err = done[k]
                        want[kind] += 1
                        want["merr"] += err
                        want["fdone"] += 1
            assert (s.present, s.formulas) == (len(keys), formulas), it
            assert {f: field(s.total, f) for f in FIELDS} == want, it
            texts = [k for k in keys if book.cells[k].formula is None and book.cells[k].kind == "s"]
            assert idx.text_cells("S", r1, c1, r2, c2) == texts


def test_a_sheet_with_no_cells_in_the_range_is_empty():
    idx = RangeIndex(book_of({"S": {"A1": 1.0}, "T": {}}))
    assert idx.summary("T", 1, 1, 10, 10).present == 0
    assert idx.summary("S", 5, 5, 9, 9).present == 0
```

- [ ] **Step 2: Run, expect failure**

Run: `python -m pytest tests/test_rangeindex.py -q`
Expected: FAIL `ModuleNotFoundError: No module named 'unexcel.rangeindex'`.

- [ ] **Step 3: Write `src/unexcel/rangeindex.py`**

```python
"""What the compiler needs to know about a rectangle of cells, in O(columns x log rows) instead of O(cells).

Per (sheet, column) one Fenwick tree holds a packed counter for every present cell: how many cells of each kind
(n s b e a blank), how many can hold an error, and how many formula cells are finished. Input cells are counted
when the index is built; a formula cell is counted by complete() when its batch closes.
"""
from __future__ import annotations

from bisect import bisect_left, bisect_right
from collections import defaultdict
from dataclasses import dataclass

from unexcel.fenwick import Fenwick
from unexcel.model import Book, Key

FIELDS = ("n", "s", "b", "e", "a", "blank", "merr", "fdone")
BITS = 21  # a column has at most 1,048,576 cells < 2**21
SHIFT = {name: BITS * i for i, name in enumerate(FIELDS)}
MASK = (1 << BITS) - 1
Column = tuple[str, int]


def field(total: int, name: str) -> int:
    return (total >> SHIFT[name]) & MASK


@dataclass(frozen=True)
class Summary:
    present: int  # non-empty cells in the rectangle
    formulas: int  # of those, formula cells
    total: int  # packed counters, read with field()


class RangeIndex:
    def __init__(self, book: Book):
        self.rows = book._rows_by_col  # (sheet, col) -> sorted rows that hold a cell
        self.cols: dict[str, list[int]] = defaultdict(list)  # sheet -> sorted columns that hold a cell
        for sheet, col in sorted(self.rows):
            self.cols[sheet].append(col)
        self.fen: dict[Column, Fenwick] = {}
        self.formula_prefix: dict[Column, list[int]] = {}  # formula cells among the first i cells of the column
        self.text_rows: dict[Column, list[int]] = {}  # rows of input text cells
        for (sheet, col), rows in self.rows.items():
            values, prefix, text = [], [0], []
            for row in rows:
                cell = book.cells[(sheet, row, col)]
                if cell.formula is not None:
                    values.append(0)
                    prefix.append(prefix[-1] + 1)
                    continue
                values.append((1 << SHIFT[cell.kind]) + ((1 << SHIFT["merr"]) if cell.kind == "e" else 0))
                prefix.append(prefix[-1])
                if cell.kind == "s":
                    text.append(row)
            self.fen[(sheet, col)] = Fenwick(values)
            self.formula_prefix[(sheet, col)] = prefix
            self.text_rows[(sheet, col)] = text

    def complete(self, key: Key, kind: str, may_err: bool) -> None:
        """A formula cell is finished: kind is its value kind (n s b e a), may_err whether it can hold an error."""
        sheet, row, col = key
        pos = bisect_left(self.rows[(sheet, col)], row)
        add = (1 << SHIFT[kind]) + (1 << SHIFT["fdone"]) + ((1 << SHIFT["merr"]) if may_err else 0)
        self.fen[(sheet, col)].add(pos, add)

    def _slices(self, sheet: str, r1: int, c1: int, r2: int, c2: int):
        cols = self.cols.get(sheet, [])
        for col in cols[bisect_left(cols, c1):bisect_right(cols, c2)]:
            rows = self.rows[(sheet, col)]
            yield (sheet, col), bisect_left(rows, r1), bisect_right(rows, r2)

    def summary(self, sheet: str, r1: int, c1: int, r2: int, c2: int) -> Summary:
        present = formulas = total = 0
        for column, i, j in self._slices(sheet, r1, c1, r2, c2):
            if j > i:
                present += j - i
                prefix = self.formula_prefix[column]
                formulas += prefix[j] - prefix[i]
                total += self.fen[column].range(i, j)
        return Summary(present, formulas, total)

    def text_cells(self, sheet: str, r1: int, c1: int, r2: int, c2: int) -> list[Key]:
        """Input text cells in the rectangle, row-major."""
        found = []
        for (s, col), _, _ in self._slices(sheet, r1, c1, r2, c2):
            rows = self.text_rows[(s, col)]
            found.extend((s, r, col) for r in rows[bisect_left(rows, r1):bisect_right(rows, r2)])
        return sorted(found)
```

- [ ] **Step 4: Run, expect pass**

Run: `python -m pytest tests/test_rangeindex.py -q`
Expected: `2 passed`.

- [ ] **Step 5: Commit**

```bash
git add src/unexcel/rangeindex.py tests/test_rangeindex.py
git commit -m "feat: RangeIndex, per-column Fenwick counters for range summaries"
```

---

### Task 5: `Compiler.ready()` returns a summary

**Files:**
- Modify: `src/unexcel/compiler.py` (imports, `Compiler.__init__`, `ready`, `_close`, new `Keys` class)
- Modify: `src/unexcel/functions.py:38-58` (`note_text_numbers`, `first_error`)
- Modify: `src/unexcel/fn/aggregate.py:17-45` (`Rng`, `_rng`)
- Modify: `src/unexcel/fn/criteria.py:70-83, 134-142`
- Modify: `src/unexcel/fn/lookup.py:17-33`
- Modify: `src/unexcel/fn/text.py:34-42`
- Test: `tests/test_compiler.py` (append) and `tests/test_shapes.py` (create)

**Interfaces:**
- Consumes: `RangeIndex`, `field` (Task 4).
- Produces:
  - `Compiler.index: RangeIndex`.
  - `Compiler.ready(area) -> Keys`; raises `NewBatch` if a formula cell of the range is in the open batch, `Unsupported("depends on unresolved cell <A1 ref>")` if one is neither open nor done (same first-offender-in-row-major-order rule as before).
  - `Keys`: `len(keys)` = non-empty cells; `keys.kinds` = `frozenset` drawn from `{"n","s","b","e","a","blank"}` (the kinds `Compiler.kind_of` returned per cell); `keys.may_err: bool`; `iter(keys)` lists the cells (O(range), only for callers that need the cells); `keys.text_cells()` = input text cells in the range, row-major.
  - `Compiler.kind_of` and `Compiler.may_err(key)` stay (still used for SUBTOTAL's skip list).

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_compiler.py`:

```python
import pytest

from unexcel.compiler import Compiler, NewBatch
from unexcel.model import Unsupported
from unexcel.refs import Area


def test_ready_reports_open_unresolved_and_done_cells():  # Review Focus 4
    book = book_of({"S": {"A1": ("=1", 1), "A2": 5.0}})
    comp = Compiler(book)
    area = Area("S", 1, 1, 2, 1)
    with pytest.raises(Unsupported, match="depends on unresolved cell S!A1"):
        comp.ready(area)
    comp.open[("S", 1, 1)] = ("c1", "num", False)
    with pytest.raises(NewBatch):
        comp.ready(area)
    comp.open.clear()
    comp.done[("S", 1, 1)] = ("num", True)
    comp.index.complete(("S", 1, 1), "n", True)
    keys = comp.ready(area)
    assert len(keys) == 2 and keys.kinds == {"n"} and keys.may_err
    assert list(keys) == [("S", 1, 1), ("S", 2, 1)]


def test_keys_of_an_empty_range():
    comp = Compiler(book_of({"S": {"A1": 1.0}}))
    keys = comp.ready(Area("S", 5, 1, 9, 1))
    assert len(keys) == 0 and keys.kinds == frozenset() and not keys.may_err
```

Create `tests/test_shapes.py`:

```python
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import verify


def _running_formula(n):
    spec, total = {}, 0.0
    for r in range(1, n + 1):
        b = (r % 97) * 1.5
        total += b
        spec[f"A{r}"] = float(r % 97)
        spec[f"B{r}"] = (f"=A{r}*1.5", b)
        spec[f"C{r}"] = (f"=SUM($B$1:B{r})", total)
    return book_of({"S": spec})


def test_running_sum_over_a_formula_column_does_not_make_a_batch_per_row():
    n = 300
    book = _running_formula(n)
    compiled = compile_book(book)
    assert len(compiled.batches) < n // 4  # 301 batches before Plan 2c
    assert {v.status for v in verify(book, compiled)} == {"PASS"}


def test_number_stored_as_text_in_a_summed_range_is_reported():  # Review Focus 5
    book = book_of({"S": {"A1": "15,000", "A2": 5.0, "B1": ("=SUM(A1:A2)", 5.0)}})
    compiled = compile_book(book)
    assert any("S!A1" in f and "number stored as text" in f for f in compiled.findings)
    assert {v.status for v in verify(book, compiled)} == {"PASS"}
```

- [ ] **Step 2: Run, expect failure**

Run: `python -m pytest tests/test_compiler.py tests/test_shapes.py -q`
Expected: the two `ready` tests FAIL (`AttributeError: 'Compiler' object has no attribute 'index'`); `test_running_sum_over_a_formula_column...` passes only after Task 7 (it needs weight batching), so expect it to FAIL now with `301 < 75`; the text-number test passes already.

- [ ] **Step 3: Edit `src/unexcel/compiler.py`**

Add to the imports:
```python
from unexcel.rangeindex import RangeIndex, field
```

Add the `Keys` class right after `class NewBatch`:

```python
class Keys:
    """The non-empty cells of a range, summarised in O(log n). Iterating lists them (O(range): avoid on big ranges)."""

    def __init__(self, comp: "Compiler", area: Area, present: int, total: int):
        self._comp, self.area, self._n = comp, area, present
        self.kinds = frozenset(name for name in ("n", "s", "b", "e", "a", "blank") if field(total, name))
        self.may_err = field(total, "merr") > 0

    def __len__(self) -> int:
        return self._n

    def __iter__(self):
        return iter(self.area.present(self._comp.book))

    def text_cells(self) -> list:
        """Input text cells in the range, row-major (the only cells SUM-like functions must warn about)."""
        a = self.area
        return self._comp.index.text_cells(a.sheet, a.r1, a.c1, a.r2, a.c2)
```

In `Compiler.__init__` add (after `self.findings`): `self.index = RangeIndex(book)`.

Replace `Compiler.ready` with:

```python
    def ready(self, area: Area) -> Keys:
        """The non-empty cells of a range that SQL will read from the table. All must be computed already."""
        s = self.index.summary(area.sheet, area.r1, area.c1, area.r2, area.c2)
        if field(s.total, "fdone") < s.formulas:
            self._blocked(area)
        return Keys(self, area, s.present, s.total)

    def _blocked(self, area: Area) -> None:
        """Some formula cell of the range is unfinished: raise for the first one, in row-major order."""
        for key in area.present(self.book):
            if self.book.cells[key].formula is None:
                continue
            if key in self.open:
                raise NewBatch
            if key not in self.done:
                raise Unsupported(f"depends on unresolved cell {a1(key)}")
        raise AssertionError("range index is out of step with the compiler")
```

In `_close`, after `comp.done.update(...)` add:
```python
        for s in current:
            comp.index.complete(s.key, KIND[s.expr.type], s.expr.err != NO_ERR)
```

- [ ] **Step 4: Edit the callers** (each replaces a loop over every key by the summary)

`src/unexcel/functions.py`:
```python
def note_text_numbers(c, keys):
    """Excel skips text in SUM-like ranges, even text that looks like a number. Tell the user."""
    for key in keys.text_cells():
        cell = c.book.cells[key]
        if _looks_numeric(cell.value):
            c.findings.append(f"{a1(key)}: number stored as text ({cell.value!r}) is skipped by Excel")


def first_error(c, area: Area, keys) -> str:
    """SQL for the first error in a range, row by row, or NO_ERR when no cell in it can hold one."""
    if not keys.may_err:
        return NO_ERR
    return f"(SELECT txt FROM {CELLS} WHERE {where_area(area)} AND kind = 'e' ORDER BY row_num, col_num LIMIT 1)"
```

`src/unexcel/fn/aggregate.py` (`Rng` keeps `keys`, now a `Keys`; the error test uses the summary unless cells are skipped):
```python
@dataclass(frozen=True)
class Rng:
    area: Area
    keys: object  # compiler.Keys
    skip: tuple = ()  # cells left out (SUBTOTAL leaves out other SUBTOTALs)
    ...
    def err(self, c) -> str:
        if self.skip:
            if not any(c.may_err(k) for k in self.keys if k not in self.skip):
                return NO_ERR
        elif not self.keys.may_err:
            return NO_ERR
        return (f"(SELECT txt FROM {CELLS} WHERE {self.where()} AND kind = 'e' "
                f"ORDER BY row_num, col_num LIMIT 1)")


def _rng(c, node, sheet: str, skip_subtotals: bool = False) -> Rng:
    area = resolve(node.text, sheet, c.book)
    keys = c.ready(area)
    note_text_numbers(c, keys)
    skip = ()
    if skip_subtotals:
        skip = tuple(k for k in keys if "SUBTOTAL(" in (c.book.cells[k].formula or "").upper())
    return Rng(area, keys, skip)
```
(Keep the existing `where` and `agg` methods of `Rng` unchanged.)

`src/unexcel/fn/criteria.py`: line 141 `if any(c.may_err(k) for k in keys):` becomes `if keys.may_err:`. Change the return annotations of `_range` and `_column` from `tuple[Area, list]` to `tuple[Area, object]`.

`src/unexcel/fn/lookup.py` `_type_of`:
```python
def _type_of(c, area: Area, keys) -> str:
    kinds = set(keys.kinds) - {"e"}  # an error found is returned as that error, whatever the type
    full = len(keys) == (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1)
    ...
```
(keep the rest of the function). Change `_area`'s annotation `tuple[Area, list]` to `tuple[Area, object]`.

`src/unexcel/fn/text.py` line 39: `if any(c.may_err(k) for k in keys):` becomes `if keys.may_err:`.

- [ ] **Step 5: Run the whole suite**

Run: `python -m pytest tests -q`
Expected: everything passes except `test_running_sum_over_a_formula_column_does_not_make_a_batch_per_row` (needs Task 7). If any other test fails, the failure shows which caller still iterates keys; fix that caller (do not weaken the test).

- [ ] **Step 6: Measure compile time**

Run: `python tests/bench_shapes.py running_formula 4000` and `python tests/bench_shapes.py running_input 4000`
Expected: `compile=` under 2 s for both (it was 20.5 s and 5.9-10 s); `pass=` full. Record the lines.

- [ ] **Step 7: Commit**

```bash
git add -u && git add tests/test_shapes.py
git commit -m "perf: Compiler.ready() answers from a range index instead of walking every cell"
```

---

### Task 6: Verifier appends rows in place

**Files:**
- Modify: `src/unexcel/compiler.py` (`render_batch`)
- Modify: `src/unexcel/verifier.py:14-26` (`execute`)
- Test: `tests/test_batches.py` (append)

**Interfaces:**
- Consumes: nothing new.
- Produces: `render_batch(steps, prev, own_only=False)`. With `own_only=True` it returns only the new rows of the batch (`WITH ctes SELECT row UNION ALL SELECT row ...`), reading `prev` for ranges; with the default it returns what it returns today. `execute()` keeps one table `cells` and runs `INSERT INTO cells <own_only SQL>` per batch.

- [ ] **Step 1: Write the failing tests** (append to `tests/test_batches.py`)

```python
from unexcel.compiler import render_batch


def test_own_only_batch_has_only_the_new_rows():
    book = book_of({"S": {"A1": 2, "A2": ("=A1*2", 4), "A3": ("=A1*3", 6)}})
    steps = compile_book(book).batches[0]
    full, own = render_batch(steps, "prev"), render_batch(steps, "prev", own_only=True)
    assert "SELECT * FROM prev" in full and "SELECT * FROM prev" not in own
    assert own.count("UNION ALL") == len(steps) - 1


def test_workbook_without_formulas_verifies_and_executes():  # Review Focus 1
    book = book_of({"S": {"A1": 1.0, "A2": "x"}})
    compiled = compile_book(book)
    assert compiled.batches == [] and execute(book, compiled) == {}
    assert verify(book, compiled) == []
```

- [ ] **Step 2: Run, expect failure**

Run: `python -m pytest tests/test_batches.py -q`
Expected: `test_own_only_batch...` FAILS with `TypeError: render_batch() got an unexpected keyword argument 'own_only'`.

- [ ] **Step 3: Edit `render_batch`** in `src/unexcel/compiler.py`

Change the signature and last lines:
```python
def render_batch(steps: list[Step], prev: str, own_only: bool = False) -> str:
    """One batch as a SELECT: the new row of every cell this batch computes, after every row of `prev` unless own_only."""
    ...  # ctes and rows exactly as today
    head = "" if own_only else f"SELECT * FROM {CELLS}\nUNION ALL\n"
    sql = f"WITH\n{ctes}\n{head}" + "\nUNION ALL\n".join(rows)
    return sql.replace(CELLS, prev)
```
(Today's `sql = f"WITH\n{ctes}\nSELECT * FROM {CELLS}\nUNION ALL\n" + ...` becomes the two lines above; output for `own_only=False` is byte-identical.)

- [ ] **Step 4: Edit `execute`** in `src/unexcel/verifier.py`

```python
def execute(book: Book, compiled: Compiled) -> dict[Key, tuple]:
    """The same row SQL as the dbt project: the seed, then each batch appended in place. Returns formula cells only."""
    con = duckdb.connect()
    with tempfile.TemporaryDirectory() as tmp:
        seed = Path(tmp) / "input_cells.csv"
        write_seed(book, seed)
        con.execute(f"CREATE TABLE cells AS SELECT * FROM {READ_SEED}", [str(seed)])
    for steps in compiled.batches:
        con.execute(f"INSERT INTO cells {render_batch(steps, 'cells', own_only=True)}")
    wanted = {s.key for s in compiled.steps}
    rows = con.execute("SELECT * FROM cells").fetchall()
    return {(s, r, c): (k, n, t) for s, r, c, k, n, t in rows if (s, r, c) in wanted}
```

- [ ] **Step 5: Run the whole suite**

Run: `python -m pytest tests -q`
Expected: all pass (apart from the Task 7 shape test).

- [ ] **Step 6: Commit**

```bash
git add -u
git commit -m "perf: verifier appends each batch in place instead of copying the cells table"
```

---

### Task 7: Batches close by SQL weight

**Files:**
- Modify: `src/unexcel/compiler.py` (constants, `compile_book`)
- Modify: `tests/test_batches.py` (`test_thousands_of_cells_no_longer_crash`)
- Test: `tests/test_batches.py` (append)

**Interfaces:**
- Produces: `compiler.BATCH_WEIGHT` (int, characters), `compiler.ROW_COST` (int), `step_weight(e: Expr) -> int`. A batch closes when it holds `BATCH` cells or its summed weight reaches `BATCH_WEIGHT`. A cell is always accepted into an empty batch, however heavy.

- [ ] **Step 1: Write the failing tests** (append to `tests/test_batches.py`)

```python
def test_batches_close_by_weight(monkeypatch):
    monkeypatch.setattr(compiler, "BATCH_WEIGHT", 1000)
    book = book_of({"S": {f"A{r}": (f"={r}", float(r)) for r in range(1, 21)}})
    sizes = [len(b) for b in compile_book(book).batches]
    assert sum(sizes) == 20 and max(sizes) < 20


def test_a_cell_heavier_than_the_budget_runs_alone(monkeypatch):  # Review Focus 2
    monkeypatch.setattr(compiler, "BATCH_WEIGHT", 1)
    book = book_of({"S": {f"A{r}": (f"={r}+{r}", 2.0 * r) for r in range(1, 6)}})
    compiled = compile_book(book)
    assert [len(b) for b in compiled.batches] == [1, 1, 1, 1, 1]
    assert {v.status for v in verify(book, compiled)} == {"PASS"}
```

In the same file replace the body line of `test_thousands_of_cells_no_longer_crash`:
`assert len(compiled.batches) == n // compiler.BATCH` becomes `assert len(compiled.batches) >= n // compiler.BATCH`.

- [ ] **Step 2: Run, expect failure**

Run: `python -m pytest tests/test_batches.py tests/test_shapes.py -q`
Expected: the two new tests FAIL (`AttributeError: ... has no attribute 'BATCH_WEIGHT'` from monkeypatch), and the running-sum shape test still FAILS.

- [ ] **Step 3: Edit `src/unexcel/compiler.py`**

Replace the comment and constant block around `BATCH = 250` with:

```python
# A batch closes at BATCH cells or when its SQL weighs BATCH_WEIGHT characters, whichever comes first. Time inside one
# statement grows faster than linearly with its cell count, and lookups are far heavier than plain arithmetic, so a
# count alone makes lookup-heavy batches slow (measured on DuckDB 1.5.6: 25 VLOOKUP cells 0.2 s, 250 cells 17.8 s).
# BATCH also keeps the UNION ALL chain below DuckDB's expression depth limit of 1000.
BATCH = 250
BATCH_WEIGHT = 20_000  # characters; Task 7 step 6 of the Plan 2c plan records how this was chosen
ROW_COST = 250  # characters a cell adds besides its expression: its CTE wrapper and the row it emits
```

Add below `step`-related definitions (after `class Step`):

```python
def step_weight(e: Expr) -> int:
    return ROW_COST + len(e.sql) + len(e.err) + len(e.txt) + len(e.kind)
```

Rewrite the loop head of `compile_book` (keep everything else):

```python
    batches: list[list[Step]] = []
    current: list[Step] = []
    weight = 0
    for key in order:
        cell = book.cells[key]
        if cell.kind == "blank":
            unresolved[key] = NO_VALUE
            continue
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
            ...  # unchanged
        step = Step(key, e, f"c{len(current) + 1}")
        comp.open[key] = (step.cte, e.type, e.err != NO_ERR)
        current.append(step)
        weight += step_weight(e)
    _close(comp, batches, current)
```

Update the module docstring sentence "A batch is up to BATCH consecutive formula cells in dependency order" to "A batch is up to BATCH consecutive formula cells (or BATCH_WEIGHT characters of SQL) in dependency order".

- [ ] **Step 4: Run the whole suite**

Run: `python -m pytest tests -q`
Expected: all pass, including `test_running_sum_over_a_formula_column_does_not_make_a_batch_per_row`.

- [ ] **Step 5: Sweep the budget**

Run, for each W in 5000 10000 20000 40000 80000:
```bash
python tests/bench_shapes.py vlookup 1000 --batch-weight W
python tests/bench_shapes.py running_input 4000 --batch-weight W
python tests/bench_shapes.py flat 20000 --batch-weight W
```
Expected: every line `pass=N/N`. Add the three `verify=` times per W. Choose the W with the smallest sum; if two are within 5% take the larger one (fewer batches, smaller dbt project). If the best W is the sweep's smallest or largest value, extend the sweep one step further before choosing.

- [ ] **Step 6: Set `BATCH_WEIGHT` to the chosen value**, replace the placeholder comment with the sweep table (three lines per W), run `python -m pytest tests -q`, expect all pass.

- [ ] **Step 7: Commit**

```bash
git add -u
git commit -m "perf: close batches by SQL weight, not cell count (sweep table in compiler.py)"
```

---

### Task 8: dbt models hold only their own rows

**Files:**
- Modify: `src/unexcel/emitter.py` (`dbt_sql`, `write_project`, new `sources_for`, `checkpoint_name`)
- Modify: `tests/test_emitter.py` (adapt three tests, add four)

**Interfaces:**
- Consumes: `render_batch(steps, prev, own_only=True)` (Task 6).
- Produces:
  - `CHECKPOINT = 64`; `checkpoint_name(k) -> "cp_0064"`; `model_name(k)` unchanged (`cells_0001`).
  - `sources_for(k) -> list[str]`: the models whose union is every cell known before batch `k` (1-based): the latest checkpoint at or below `k - 1` (or `input_cells`) plus the batch models after it.
  - `dbt_sql(steps, sources)`: the model SQL; defines `all_cells` as the union of `sources`, then the own-rows batch SQL reading `all_cells`.
  - `write_project`: `cells_NNNN.sql` per batch (own rows), `cp_NNNN.sql` every 64 batches and after the last batch, `workbook_cells.sql` reading the last checkpoint (or `input_cells` when there are no batches).

- [ ] **Step 1: Write the failing tests**

Add to `tests/test_emitter.py`:

```python
import re
from pathlib import Path

import unexcel.compiler as compiler
from unexcel.compiler import READ_SEED
from unexcel.emitter import CHECKPOINT, checkpoint_name, model_name, sources_for
from unexcel.verifier import execute


def test_sources_for_reads_the_latest_checkpoint_plus_later_batches():  # Review Focus 3
    assert sources_for(1) == ["input_cells"]
    assert sources_for(64) == ["input_cells"] + [model_name(i) for i in range(1, 64)]
    assert sources_for(65) == [checkpoint_name(64)]
    assert sources_for(66) == [checkpoint_name(64), model_name(65)]
    assert sources_for(129) == [checkpoint_name(128)]
    assert CHECKPOINT == 64


def _run_models(out: Path):
    """Run the generated models in DuckDB without dbt: ref('x') becomes the table x, built in batch order."""
    con = duckdb.connect()
    con.execute(f"CREATE TABLE input_cells AS SELECT * FROM {READ_SEED}", [str(out / "seeds" / "input_cells.csv")])
    models = out / "models"
    n = len(list((models / "unexcel_cells").glob("cells_*.sql")))
    names = []
    for k in range(1, n + 1):
        names.append(model_name(k))
        if (models / "unexcel_cells" / f"{checkpoint_name(k)}.sql").exists():
            names.append(checkpoint_name(k))
    for name in names:
        sql = (models / "unexcel_cells" / f"{name}.sql").read_text(encoding="utf-8")
        con.execute(f"CREATE TABLE {name} AS " + re.sub(r"\{\{ ref\('(\w+)'\) \}\}", r"\1", sql))
    final = (models / "workbook_cells.sql").read_text(encoding="utf-8")
    return con.execute(re.sub(r"\{\{ ref\('(\w+)'\) \}\}", r"\1", final)).fetchall()


def test_generated_models_equal_the_verifier_across_checkpoints(tmp_path, monkeypatch):  # Review Focus 3
    monkeypatch.setattr(compiler, "BATCH", 1)  # one cell per batch: 130 batches, so cp_0064, cp_0128 and a final cp_0130
    spec = {"A1": 1.0}
    spec.update({f"B{r}": (f"=A1+{r}", 1.0 + r) for r in range(1, 131)})
    spec["C1"] = ("=SUM(B1:B130)", sum(1.0 + r for r in range(1, 131)))
    book = book_of({"S": spec})
    compiled = compile_book(book)
    assert len(compiled.batches) > 129
    write_project(book, compiled, tmp_path, "many")
    names = {p.name for p in (tmp_path / "models" / "unexcel_cells").iterdir()}
    assert {"cp_0064.sql", "cp_0128.sql", f"cp_{len(compiled.batches):04d}.sql"} <= names
    got = {(s, r, c): (k, n, t) for s, r, c, k, n, t in _run_models(tmp_path)}
    assert got == execute(book, compiled)


def test_project_without_formulas_runs(tmp_path):  # Review Focus 1
    book = book_of({"S": {"A1": 1.0}})
    write_project(book, compile_book(book), tmp_path, "plain")
    assert _run_models(tmp_path) == []


def test_each_batch_model_selects_only_its_own_rows(tmp_path):
    book = book_of({"S": {"A1": 2, "A2": ("=A1*2", 4)}})
    write_project(book, compile_book(book), tmp_path, "own")
    sql = (tmp_path / "models" / "unexcel_cells" / "cells_0001.sql").read_text(encoding="utf-8")
    assert "all_cells AS" in sql and "{{ ref('input_cells') }}" in sql and "SELECT * FROM all_cells" not in sql
```

Adapt existing tests in `tests/test_emitter.py`:
- `test_writes_dbt_project`: change the last assertion to
  `assert "{{ ref('cp_0001') }}" in (tmp_path / "models" / "workbook_cells.sql").read_text(encoding="utf-8")`.
- `test_rerun_removes_stale_batches`: the assertion becomes
  `assert sorted(p.name for p in (tmp_path / "models" / "unexcel_cells").iterdir()) == ["cells_0001.sql", "cp_0001.sql"]`.
- `test_dbt_sql_survives_jinja`: replace its last two lines with
```python
    rendered = jinja2.Template(dbt_sql(steps, ["input_cells"])).render(ref=lambda name: name)
    prefix = "WITH\nall_cells AS (\nSELECT * FROM input_cells\n),\n"
    assert rendered == render_batch(steps, "all_cells", own_only=True).replace("WITH\n", prefix, 1)
```

- [ ] **Step 2: Run, expect failure**

Run: `python -m pytest tests/test_emitter.py -q`
Expected: collection error `ImportError: cannot import name 'CHECKPOINT' from 'unexcel.emitter'`.

- [ ] **Step 3: Edit `src/unexcel/emitter.py`**

Update the module docstring: models hold their own rows; checkpoints every 64 batches. Replace `dbt_sql`, `final_sql` use and `write_project`:

```python
CHECKPOINT = 64  # a model reads the latest checkpoint plus at most this many batch models


def checkpoint_name(k: int) -> str:
    return f"cp_{k:04d}"


def sources_for(k: int) -> list[str]:
    """Models whose union is every cell known before batch k (1-based): latest checkpoint, then later batches."""
    base = (k - 1) // CHECKPOINT * CHECKPOINT
    first = "input_cells" if base == 0 else checkpoint_name(base)
    return [first] + [model_name(i) for i in range(base + 1, k)]


def _union(models: list[str]) -> str:
    return "\nUNION ALL\n".join("SELECT * FROM {{ ref('" + m + "') }}" for m in models)


def dbt_sql(steps: list[Step], sources: list[str]) -> str:
    """One batch as a dbt model: the new rows only, reading all earlier cells through the CTE all_cells.
    Braces from workbook text are escaped so Jinja renders them literally."""
    body = re.sub(r"[{}]", lambda m: "{{ '" + m.group() + "' }}", render_batch(steps, "all_cells", own_only=True))
    return body.replace("WITH\n", f"WITH\nall_cells AS (\n{_union(sources)}\n),\n", 1)
```

`write_project` body from `prev = "input_cells"` on:

```python
    for old in batches_dir.glob("*.sql"):  # a rerun must not leave models from a bigger workbook
        old.unlink()
    ...  # dbt_project.yml, profiles.yml, seed as before
    n = len(compiled.batches)
    last = "input_cells"
    for k, steps in enumerate(compiled.batches, 1):
        (batches_dir / f"{model_name(k)}.sql").write_text(dbt_sql(steps, sources_for(k)) + "\n", encoding="utf-8")
        if k % CHECKPOINT == 0 or k == n:
            last = checkpoint_name(k)
            (batches_dir / f"{last}.sql").write_text(_union(sources_for(k) + [model_name(k)]) + "\n", encoding="utf-8")
    (out / "models" / "workbook_cells.sql").write_text(final_sql(last) + "\n", encoding="utf-8")
    (out / "models" / "schema.yml").write_text(SCHEMA, encoding="utf-8")
```

The earlier loop `for old in batches_dir.glob("cells_*.sql")` is replaced by the `"*.sql"` glob above (checkpoints live in the same folder). Remove the now-unused `CELLS` import from `emitter.py` if nothing else uses it.

- [ ] **Step 4: Run emitter tests, then the whole suite**

Run: `python -m pytest tests/test_emitter.py -q` then `python -m pytest tests -q`
Expected: all pass, including `test_dbt_build_passes_on_generated_project` when `dbt` is installed (activate `.venv` so `dbt` is on PATH; otherwise it is skipped, say so in the report).

- [ ] **Step 5: Measure**

Run: `python tests/bench_shapes.py running_input 2000 --dbt` and compare `db_mb` and `dbt=` with the Task 1 line.
Expected: `dbt_ok=True`; `db_mb` at most one tenth of the Task 1 value (criterion 4 in the spec). If not, stop and report the number.

- [ ] **Step 6: Commit**

```bash
git add -u
git commit -m "perf: dbt batch models store only their own rows, with a checkpoint every 64 batches"
```

---

### Task 9: Close-out

**Files:**
- Modify: `PROGRESS.md`, `AGENTS.md` (status paragraph only), `docs/superpowers/specs/2026-10-14-plan-2c-speed-design.md` (status line + results table)

- [ ] **Step 1: Run every benchmark and compare with the Task 1 baseline**

```bash
python tests/bench_shapes.py running_formula 4000
python tests/bench_shapes.py running_input 4000
python tests/bench_shapes.py vlookup 1000
python tests/bench_shapes.py flat 20000
python tests/bench_shapes.py running_input 2000 --dbt
UNEXCEL_PERF=1 python -m pytest tests/test_perf.py -s
```
Check against the spec's success criteria 2, 3, 4 and 5: VLOOKUP 1k verify at most 15 s; running SUM over inputs 4k verify at most 12 s; running SUM over a formula column 4k at most 10 batches and compile at most 2 s; 200k flat no slower than the Task 1 baseline; dbt storage at most one tenth. Any miss is reported with its number, not hidden; do not change the targets.

- [ ] **Step 2: Run the full suite and the saved-Excel tests**

Run: `python -m pytest tests -q` and `python -m pytest tests/test_excel_saved.py -v`
Expected: all pass (169+ passed; the same 2 skipped plus any dbt skip). The saved-Excel tests are the check that no verdict changed against real Excel.

- [ ] **Step 3: Update docs**

In `PROGRESS.md`: replace the "Where we stopped" block with the real state (Plan 2c built on branch `plan-2c`, results table from step 1, next = Plan 2d table mode, backlog unchanged) and remove the two items of "Known gaps" that are fixed. In `AGENTS.md` replace the status block at the top with a pointer to `PROGRESS.md` and the date. In the spec set `Status: implemented` and append the "before/after" table from step 1.

- [ ] **Step 4: Commit**

```bash
git add -u
git commit -m "docs: Plan 2c results, PROGRESS and AGENTS point at the current state"
```

- [ ] **Step 5: Stop and report.** Do not push or open a PR; ask the user.

---

## Self-review (done while writing)

- **Spec coverage:** weight batching (Task 7), level scheduling (Task 3), compile index and graph (Tasks 3-5), append in place (Task 6), dbt own-rows + checkpoints (Task 8), criteria 1-5 (Task 9 step 1-2), baseline re-measurement on this machine (Task 1). The spec's "quadratic graph" amendment is in Task 3.
- **Placeholders:** none; the one number chosen by measurement (`BATCH_WEIGHT`) has a default (20,000) and an explicit rule to replace it.
- **Type consistency:** `Keys.kinds/.may_err/len/iter/text_cells`, `RangeIndex.complete/summary/text_cells`, `field`, `render_batch(..., own_only)`, `sources_for`, `checkpoint_name`, `dbt_sql(steps, sources)` are used with the same names and signatures in every task that mentions them.
- **Known risk carried from the spec:** the weight constants were tuned on one machine; Task 7 step 5 re-derives them here.
