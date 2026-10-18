# Plan 2d: table mode, slice 1 (row-wise columns) - design

Date: 2026-10-14. Status: implemented on branch `plan-2d` (2026-10-16); results at the end. Builds on Plan 2c (merged, PR #6).

## Goal

A workbook column that repeats one formula down its rows becomes one readable SQL statement instead of one CTE per
cell. Two payoffs, both measured as the point of this plan:

1. **Readable dbt output.** A region compiles to a model an analyst can read and maintain
   (`amount * 1.1 AS tax`), which is the product promise in `DESIGN.md` D8.
2. **Speed and a small dbt project.** One statement per region, not one per 10-30 cells.

Cell mode stays as the permanent fallback (D8). Slice 1 handles only **row-wise** columns; everything else keeps
compiling exactly as today, in the same project.

Non-negotiables (AGENTS.md section 7): no false PASS; Excel's saved values stay the answer key; every formula cell
still gets PASS, FAIL or UNRESOLVED; runtime dependencies stay `openpyxl` and `duckdb`; cell values never leave the
machine. A region that cannot be compiled is never an error: it dissolves to cell mode with its reason recorded.

## Measured basis

Prototype (scratch, this machine): three formula columns (`A*1.1`, `IF(tax>100,tax,0)`, `net-A`) over 200,000 rows
= 600,000 cells, written as one `INSERT INTO cells WITH <pivot CTE>, <one CTE per column> SELECT ... UNION ALL ...`
statement: **0.20 s**. For comparison, today's cell mode takes 390 s for 200,000 cells (Plan 2c results) and 28 s for
20,000. The set-based shape is the win; the pivot CTE costs almost nothing.

## Decisions

- **Approach C: regions as set-based batches over the existing long `cells` table.** One data representation
  (rejected: wide per-sheet grid seeds, which would need two representations kept in step).
- **Same compiler.** The region compiler is the current `Compiler` with a different cell reader, so every verified
  Excel rule (error codes, text-to-number, DIV/0, powers, 93 functions) carries over unchanged.
- **Two-layer output.** A wide, readable model per region (the deliverable), plus a mechanical unpivot of it back into
  long `cells` rows (what later cell-mode batches read and what the verifier compares). The unpivot is derived from the
  wide model, so what is verified is exactly what ships.

## Design

### 1. Region detection (`src/unexcel/regions.py`, new)

- **R1C1 key of a formula:** the formula with every reference rewritten relative to its own cell
  (`$` marks absolute parts). Two cells have the same key iff they are the same formula "pasted down".
- **Column run:** consecutive rows in one (sheet, column) where every cell is a non-array formula with the same R1C1
  key, a saved value (kind not `blank`) and no volatile function. Minimum `MIN_ROWS = 3`.
- **Region:** column runs on one sheet with exactly the same row span, grouped together. Runs with different spans
  are separate regions.
- **Admission of a column** (every reference in its template formula must be one of):
  - a **same-row** cell (relative row offset 0) that is an *input* or another *region column*;
  - an **absolute single cell**, or a **range entirely outside the region's rows**, read as a scalar through the
    existing `Compiler`/`RangeIndex` machinery (these do not depend on the row).
- **Input columns read by a region** must be *dense over the region's rows* (every row present), hold a single value
  kind (all numbers, all text or all booleans) and contain no error cells. Otherwise the referencing column is
  rejected (blank-vs-zero and mixed-kind semantics stay with cell mode in slice 1).
- **Rejected columns** stay in cell mode: formulas that read a previous or next row of the region (running totals),
  a growing range such as `SUM($B$1:Br)`, a range overlapping the region's rows, lookups and conditional aggregates
  over region rows, or anything the region compiler raises `Unsupported` for.
- **Names:** the cell above the region, if it is unique text in every column of the region, names the SQL column
  (made a valid lower-case identifier); otherwise `col_<letter>`.

### 2. Ordering (`src/unexcel/graph.py`)

A region is one node. `formula_order` gains an optional `groups` argument (region cells grouped per region):

- a group's waiting count is the sum of its members' waits minus the dependencies *inside* the group (same-row
  references between region columns); when the group is released all members are ordered at once and the "unordered"
  Fenwick counts and users are updated for every member;
- the return value lists *units*: a single `Key` or a `Region`.
- **Cycle safety.** Grouping can create a cycle the cell graph does not have (region A reads a cell that reads
  region B which reads A). If grouped ordering reports a `CYCLE` for any cell that the ungrouped ordering does not,
  the regions involved are dissolved and ordering is recomputed. The loop only ever shrinks the set of regions, so it
  terminates. The common path (no conflict) costs one ordering pass.

### 3. Compilation (`src/unexcel/compiler.py` plus a new `src/unexcel/region.py`)

- `compile_book` walks units. A cell unit behaves exactly as today. A region unit first closes the open batch
  (everything it reads from the table must be finished), then compiles the region into **its own batch**; the batch
  after it starts empty. A region that raises `Unsupported` is dissolved: its cells are compiled one by one in the
  region's own order (columns in dependency layers, then rows), so it can never turn a cell from PASS into a worse
  state than cell mode gave it.
- **Region SQL.** `WITH <pivot of the columns it reads, per row>, <one CTE per formula column in dependency order>
  SELECT ...`. Same-row references compile to column references (`amount`), with an error column only when the
  column can hold an error (so a plain `A * 1.1` is plain SQL). Absolute references stay scalar subqueries.
- **Output rows.** For every formula column, `(sheet, row, col, kind, num, txt)` rows with the same kind/num/txt rules
  `render_batch` uses; the unpivot appends them to `cells` (`INSERT INTO cells`, as the verifier already does).
- `Compiled` carries region batches next to cell batches; `compiled.steps`-style iteration used by `verify` yields
  every formula cell (region cells included) so the per-cell comparison and the report are unchanged.

### 4. Verifier and dbt

- **Verifier:** runs the region batch like any other (`INSERT INTO cells <unpivot of the region statement>`); no new
  comparison logic.
- **dbt:** a region becomes two models. `<sheet>_<names>` is the **wide readable model** (the pivot CTE is a short,
  commented block at the top; the formula columns are the body). `<same>_cells` is the mechanical **unpivot** of it
  into long rows. Later batch models and checkpoints read the long model through the existing `sources_for` union, so
  Plan 2c's checkpoint arithmetic is unchanged. `schema.yml` documents the region (sheet, rows, columns, R1C1 form).
- A flat 20,000-cell workbook then yields a handful of models instead of 741.

### 5. Report

`REPORT.md` gets a short "Regions" section: each region (sheet, rows, columns, model name) and, for each rejected or
dissolved candidate, the reason (for example "reads the previous row"). The reasons are the product's map of what
slice 2 should cover.

## Out of scope for slice 1

- Previous-row references and running totals as window functions; growing ranges (`SUM($B$1:Br)` as a window).
- `SUMIFS` as `GROUP BY` and `VLOOKUP`/`XLOOKUP` as a join; any aggregate over region rows.
- Regions across sheets, regions with gaps, merged cells, hidden-row semantics, header rows of more than one line.
- Input columns with blanks, errors or mixed value kinds.
- Anything touching how Excel functions evaluate (this plan changes where SQL runs, not what it means).

## Testing

- **Differential test (the main safety net).** Generated row-wise workbooks (random mixes of arithmetic, `IF`,
  text, error-producing formulas, absolute and cross-sheet scalars, tiny and large row counts, regions that must
  dissolve) are compiled twice, with regions and with `regions=False`; every cell's status and value must be
  identical. Run on the 3 real-Excel saved workbooks too (their verdicts must not change).
- **Unit tests:** R1C1 key, column runs, admission and rejection reasons, header naming (duplicates, empty,
  non-text), dissolve on cycle, dissolve on `Unsupported`, group ordering vs cell ordering.
- **dbt:** the without-dbt model-run harness from Plan 2c extended to region models; a real `dbt build` on a workbook
  with a region (the existing optional test).
- **Golden readability test:** the wide model of a small sample region is checked into the tests as expected text,
  so a reviewer sees exactly what an analyst will read.
- **Benchmarks:** `tests/bench_shapes.py` gains the row-wise shape(s) it already has (`flat`) with regions on/off.

## Success criteria (measured on the same machine as the Plan 2c numbers)

1. Verdicts identical: full suite, the 3 saved-Excel tests and the differential test pass; zero cells change status.
2. `flat` 20,000 rows: verify at most 3 s (from 28 s); 200,000-formula benchmark total at most 60 s (from 390 s).
3. `--dbt` on flat 20,000 rows: at most 10 models (from 741 batch models), `dbt build` passes and its result equals the
   verifier's; run time well below 226 s.
4. Shapes that do not become regions (running totals, lookups) are not slower than after Plan 2c.
5. The golden wide model for the sample region reads as plain SQL a reviewer accepts (reviewed by the user).

## Risks

- **A wrong relative/absolute classification** (mapping `$B$2` to the row-wise column, or the reverse) would silently
  compute wrong values. Mitigation: classification comes from the parsed reference text, not from position; the
  differential test includes absolute refs that happen to point into the region's own rows.
- **Grouped ordering bugs** (a region placed before something it reads, or after a reader). Mitigation: the Plan 2c
  range index still checks every read at compile time (an early read becomes UNRESOLVED, never a wrong value), plus
  the dissolve-on-conflict loop and an oracle test against ungrouped ordering on random workbooks.
- **Blank and mixed-kind semantics** are deliberately not handled in slice 1 (rejected, cell mode).
- **SQL text size / depth** for regions with many columns (one CTE per column): bounded by the existing `MAX_SQL`
  check; an over-large region dissolves.
- **Header naming collisions** produce `col_<letter>` names rather than guessing.
- **Pivot cost on wide sheets:** the pivot reads only the columns the region uses (`col_num IN (...)`), measured in
  the prototype at negligible cost for 200k rows.

## Results (2026-10-16, same machine as the Plan 2c numbers)

| Shape | After Plan 2c | After Plan 2d |
|---|---|---|
| Flat 20k, verify | 28.3 s, 741 batches | 0.2 s, 1 batch |
| Running SUM over inputs, 4k | verify 6.3 s | verify 6.1 s (first run 7.1 s) |
| Running SUM over formula column, 4k | verify 11.9 s, 315 batches | verify 6.2 s, 168 batches |
| VLOOKUP per row, 1k | verify 5.9 s | verify 6.1 s, 167 batches (re-measured after the final review; lookup columns stay in cell mode) |
| `--dbt` flat 20k | 225.9 s, 213 MB, 741 batches | 9.6 s, 1.6 MB, 1 batch |
| 200k test | 390.3 s, 8378 batches | 120.6 s (read 5.5, compile 17.9, verify 89.2, emit 8.0), 1853 batches |

Criteria: 1 MET as identical statuses and values within the verifier tolerance (evidence: the 400-workbook differential test, about 25 adversarial workbooks from the final review, and lookup/two-sheet workbooks through real dbt; the 3 saved-Excel workbooks contain no regions and do not test regions); 2 flat-20k part MET (0.2 s), 200k part NOT MET (120.6 s against at most 60 s); 3 MET (9.6 s, 4 model
files; `dbt build` equals the verifier for flat lookup-free regions plus the test workbooks); 4 MET; 5 golden text in `tests/test_region.py`, the user has not yet read it. Not covered by slice 1: a cell-mode column
between region columns dissolves the whole region; regions need one row span; functions that read a reference
themselves stay in cell mode, and so does any column DuckDB cannot plan row-wise in both run forms (lookups: INDEX, MATCH,
VLOOKUP, HLOOKUP, XLOOKUP, COUNTIFS with a row-wise criterion; slice 2: lookups as joins). The 200k figure predates the
final-review fixes and was not re-run.
