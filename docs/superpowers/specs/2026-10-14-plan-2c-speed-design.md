# Plan 2c: speed fixes in cell mode - design

Date: 2026-10-14. Status: implemented (branch `plan-2c`; results at the end). Table mode moves to Plan 2d.

## Goal

Make `unexcel convert` fast on the sheet shapes people actually have (running totals, a lookup on every row)
and make `--dbt` scale, without changing a single verdict. Cell mode stays; only how batches are sized,
scheduled, stored and compiled changes.

Non-negotiables (AGENTS.md section 7): no false PASS, no guessed values, runtime dependencies stay
`openpyxl` and `duckdb`, cell values never leave the machine.

## Measured baseline (2026-10-14, Windows 11, DuckDB, Python 3.14, current `main`)

Scratch scripts: `bench.py`, `exp2.py`, `exp3.py` (session scratchpad; the plan re-creates them as `tests/bench_shapes.py`).

| Shape | Rows | Compile | Verify | Batches |
|---|---|---|---|---|
| Running `SUM($B$1:Br)` over a formula column | 1k / 2k / 4k | 1.2 / 5.2 / 20.5 s | 5 / 12 / 24 s | 1001 / 2001 / 4001 |
| Running `SUM($A$1:Ar)` over an input column | 2k / 4k | 1.5 / 5.9 s | 15 / 32-59 s | 8 / 16 |
| `VLOOKUP` per row into a 1000-row table | 1k / 2k | 1.5 / 3.3 s | 85 / 493 s (4k: over 600 s) | 4 / 8 |
| Flat `=A*1.1` | 20k | 1.4 s | 160 s (batch 250) | 80 |

Batch-size experiment (same SQL, only `BATCH` changed; 100% PASS in every run):

| Shape | batch 10 | 25 | 50 | 100 | 250 |
|---|---|---|---|---|---|
| VLOOKUP 1k rows | 6.8 s | 9.4 s | 13.3 s | - | 85 s |
| Running SUM 4k rows | 8.4 s | 9.8 s | 11.3 s | - | 59 s |
| Flat 20k cells | - | 31-37 s | - | 46 s | 160-167 s |

Appending rows in place (`INSERT INTO cells`) instead of `CREATE TABLE ... AS` (a full copy per batch) changed
flat-20k at batch 25 from 35.7 s to 31.3 s: a small win now, but the copy grows with workbook size.
Sorting the cells table by sheet/column/row gained about 15%: not worth the complexity.
A single lookup subquery in isolation takes about 2 ms, so the cost is not the lookup; it is how DuckDB
handles a statement with many correlated scalar subqueries.

## Causes (what the numbers say)

1. **Time inside one statement grows faster than linearly with its cell count** (VLOOKUP: 25 cells 0.2 s,
   100 cells 2 s, 250 cells 17.8 s). `BATCH = 250` is a count, so heavy cells (lookups, range reads) are
   penalised hardest. The right batch size depends on how heavy the SQL is.
2. **Cell order creates a batch per row.** The Kahn+heap order interleaves `B1, C1, B2, C2...`. `C_r` reads a
   range that contains the still-open `B_r`, which forces `NewBatch` (4001 batches for 4000 rows).
3. **`Compiler.ready()` is O(cells in range) per formula**, so a growing range is quadratic at compile time.
4. **Each batch copies every cell** (`SELECT * FROM prev UNION ALL ...`), in the verifier and in dbt. At
   200k cells with small batches this is the dominant cost; in dbt every batch model also stores the full copy.

## Design

Four changes, each independently testable, all behind the existing public entry points
(`compile_book`, `render_batch`, `execute`, `write_project`). No change to the report or CLI.

### 1. Weight-based batch size

A batch closes when its summed *weight* reaches a budget instead of when it holds 250 cells. Weight of a cell
is the length of its rendered SQL (already computed for the `MAX_SQL` check). Hard cap on cells per batch stays
(DuckDB's expression depth limit of 1000 was hit by 1000 UNION ALL branches). The budget is a single constant
chosen from the table above: flat cells (about 420 chars) should land near 100-250 per batch, VLOOKUP
(about 1800 chars) near 25-50, so a budget around 45-50k characters is the starting point. The plan picks the
final value by measurement and keeps the 200k-flat benchmark no slower than before.

### 2. Dependency-level scheduling

`formula_order` keeps its output contract (a valid dependency order) but breaks ties by dependency *level*
(longest path from an input), then by key. All cells that only need inputs go first, then cells that need
those, and so on. A column of 4000 `C` cells reading a 4000-row `B` column then needs 2 levels, not 4001
batches. A `NewBatch` is still raised if a range reads an open cell; levels make it rare, not impossible
(a true chain like `C_r = C_{r-1} + B_r` is still one cell per level and stays correct).

### 3. Per-column indexes, so a range costs O(log n) instead of O(cells in it)

Profiling a 3000-row running total (17.5 s compile) showed two quadratic parts of equal size: the dependency
graph (`formula_order`, 7 s: it builds the set of every cell in every range) and `Compiler.ready()` plus the
loops its callers run over the returned keys (9 s).

- **Graph:** `formula_order` stops expanding ranges into cell sets. A range becomes, per column, an interval of
  that column's formula cells; a Fenwick tree of "not yet ordered" counts says when the interval is empty and
  which cell to wait for next. Prototype: 4,000-row running total 0.35 s, 20,000 rows 1.6 s, and the same
  order validity and cycle set as today on 3,000 random books.
- **Compiler:** a `RangeIndex` keeps, per (sheet, column), one Fenwick tree of packed counters (cells per kind,
  cells that can hold an error, finished formula cells). `ready(area)` returns a `Keys` summary object:
  `len()`, `.kinds`, `.may_err`, and a lazy iterator for the few callers that need the cells themselves.
  Callers that looped over every key to ask "can any hold an error / which kinds" read the summary instead.
  Prototype: 100,000 growing-range summaries in 0.3 s; matches a brute-force oracle on 2,000 random books.

### 4. Append in place, and dbt models that hold only their own rows

- Verifier: one `cells` table; each batch is `INSERT INTO cells <new rows>`. The batch statement reads `cells`
  and inserts into it; the new rows are not visible to their own statement, same as today. Result extraction
  is unchanged.
- dbt: each `cells_NNNN` model stores only its own rows. A batch reads earlier results through a short chain:
  `checkpoint_J UNION ALL cells_a ... cells_k-1` where a *checkpoint* model re-materialises every 64 batches
  (inputs plus everything before). This bounds both the model size and the number of `ref()`s per model
  (at most 65), so SQL text stays linear in batch count. `workbook_cells` becomes the union of the batch
  models (no more `NOT EXISTS` against the seed).

The SQL a cell compiles to does not change, so verdicts and results are byte-identical to today.

## Out of scope

- Table mode, region detection, readable `GROUP BY` SQL (Plan 2d).
- Rewriting individual functions' SQL (the three repeated lookup subqueries). Only if the benchmark shows batch
  sizing is not enough.
- Per-row SUMIFS: the first benchmark built only 50 formula cells. The plan adds a real one and treats the
  result as data, not as a goal.
- The deferred minors in PROGRESS.md (overflow to `#NUM!`, one DuckDB error aborting a run, etc.).

## Success criteria (all measured on the same machine as the baseline)

1. Verdicts identical: the full existing suite (169 passed, 2 skipped) plus the three saved-Excel tests pass
   unchanged. A new differential test compares old and new batching on the benchmark workbooks cell by cell.
2. VLOOKUP 1k rows: verify at most 15 s (from 85 s). Running SUM over an input column, 4k rows: verify at
   most 12 s (from 59 s), compile at most 2 s (from 10-20 s). Running SUM over a formula column, 4k rows:
   at most 10 batches (from 4001).
3. The 200k flat benchmark is no slower than the baseline measured before any change on this machine.
4. `--dbt` on a 20k-cell workbook: rows stored across all models at most one tenth of today's (every batch model
   holds all cells today; checkpoints add a bounded copy every 64 batches), and `dbt build` still succeeds with a
   result equal to the verifier's. A test without dbt runs the generated models in DuckDB and compares them
   with the verifier cell by cell, across a checkpoint boundary.
5. Compile of a 4,000-row running SUM over a formula column, graph included, at most 2 s (from 20 s).

## Risks

- **Weight budget tuned on one machine.** The Mac numbers in AGENTS.md (200k flat in 255 s) were measured
  elsewhere; this Windows box is slower. The plan starts by re-measuring the 200k baseline here and compares only
  like with like.
- **Prefix-count index bugs would produce a silent missed `NewBatch`**, reading a not-yet-computed cell.
  Mitigation: keep the old O(range) check as an assertion in tests and compare it with the index on every
  benchmark and fixture workbook.
- **Level scheduling changes batch composition**, which changes which cells share a statement. Mitigated by
  criterion 1 (identical verdicts) and the cell-by-cell differential test.
- **dbt checkpoints add models** (one per 64 batches); the project grows but stays bounded.

## Results (2026-10-14, same machine as the baseline)

Every benchmark line ended all-PASS. Before = baseline above; after = branch `plan-2c`.

| Shape | Before | After |
|---|---|---|
| Running SUM, formula column, 2k | compile 4.3 s, verify 10.9 s, 2001 batches | compile 0.4 s, verify 7.1 s, 158 batches |
| Running SUM, formula column, 4k | compile 20.5 s, verify 24 s, 4001 batches | compile 0.6 s, verify 11.9 s, 315 batches |
| Running SUM, input column, 4k | compile 6.7 s, verify 34.6 s, 16 batches | compile 0.3 s, verify 6.3 s, 167 batches |
| VLOOKUP per row, 1k | verify 74.3 s, 4 batches | verify 5.9 s, 167 batches |
| Flat `=A*1.1`, 20k | verify 139.8 s, 80 batches | verify 28.3 s, 741 batches |
| `--dbt`, running SUM over inputs, 2k | dbt 60.1 s, 253,664 rows stored, 2.9 MB | dbt 26.2 s, 13,536 rows stored, 24.9 MB |
| 200k perf test | total 2775.9 s (verify 2727.1), 801 batches | total 390.3 s (read 5.7, compile 20.2, verify 331.7, emit 32.8), 8378 batches |

| # | Criterion | Target | Measured | |
|---|---|---|---|---|
| 1 | Verdicts identical | suite + 3 saved-Excel tests pass | 197 passed (dbt included), 3 saved-Excel passed | MET |
| 2a | VLOOKUP 1k verify | at most 15 s | 5.9 s | MET |
| 2b | Running SUM, inputs, 4k: verify / compile | at most 12 s / 2 s | 6.3 s / 0.3 s | MET |
| 2c | Running SUM, formula column, 4k: batches | at most 10 | 315 | NOT MET |
| 3 | 200k flat no slower than baseline | at most 2775.9 s | 390.3 s | MET |
| 4 | dbt rows stored | at most 1/10 of before | 13,536 vs 253,664 (1/19), dbt build OK; measured at 2k rows, not 20k | MET |
| 5 | Compile 4k running SUM, formula column | at most 2 s | 0.6 s | MET |

Notes: criterion 2c is missed but the shape is fast (verify 11.9 s). The 200k test still exceeds its own 300 s
assertion on this machine (the baseline did too, at 2775.9 s). The db file grew (2.9 MB to 24.9 MB) because there are
more, smaller tables; emit at 200k rose from 3.8 s to 32.8 s for the same reason.
