# Running totals as window functions (table mode slice 2, piece 2)

Status: accepted 2026-10-18 (the user delegated every design decision; approach A chosen).

## Goal

Columns that carry a running total become row-wise region columns computed with one SQL window function, instead of
one cell-mode step per row. This is the shape that keeps the 200k benchmark (`tests/test_perf.py`) above its 60 s
target: `C{r} = C{r-1} + B{r}`.

## Shapes accepted (checked on the region's first row; the R1C1 key makes every row the same)

1. **Recurrence**: the whole formula is `X{r-1} + E`, `X{r-1} - E` or `E + X{r-1}`, where `X{r-1}` is the formula's
   own column one row up (no sheet prefix, relative row) and `E` is any expression a region already accepts.
   - seed = the cell `X{r1-1}` above the run, converted with Excel's `+` rules (`num`): empty -> 0, header text ->
     `#VALUE!` on every row, a formula seed is read once from the cells table (it is ordered before the region).
   - value: `seed ± running SUM of num(E)` from `r1` to `r`.
   - error: `X{r-1} ± E` -> the earliest error (seed first, then E row by row); `E + X{r-1}` -> the latest error of
     E at or above the row, else the seed's error. This is exactly what Excel's left-to-right error rule gives when
     the formula is evaluated row by row.
2. **Growing sum**: a call `SUM(<range>)` with that one argument, the range `$B$k:B{r}` (or `B$k:B{r}`): one column
   other than the formula's own, anchored start row `k <= r1`, end on the formula's row, no sheet prefix. It may sit
   anywhere inside a larger formula.
   - value: SUM of the numbers in rows `k..r` of B; text and TRUE/FALSE are skipped like cell-mode SUM; rows `k..r1-1` (above the region) are a constant read once with the cell-mode SUM.
   - error: the earliest error by row (rows above the region first).

Everything else that reads another row stays in cell mode with the existing reason ("moves with the row").

## Where it lives

- `regions.find_regions`: recognises both shapes on the parsed template; a recurrence stores `recur[col] = (side,
  op)` and keeps `E` as the template; a growing sum is replaced by a placeholder ref `__RUN_<src>_<k>__` and recorded
  in `runs[col]`; the source column joins `reads[col]` so column order computes it first. `Region` carries both maps.
- `regions._reads_outside_formula`: a recurrence seed that is a formula, or formula cells in rows `k..r1-1` of a
  growing sum, are reads of formula cells outside the region (so such a column is not "pure").
- `region.compile_region`: each column becomes one or two CTE layers. A growing sum adds a window layer
  (`SUM(...) OVER (ORDER BY row_num ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)` plus `arg_min` of the error by
  row) before the column's own layer. A recurrence adds a per-row layer for `num(E)` and then the window layer that is
  the column. Errors use `arg_min`/`arg_max(err, row)`; checked on DuckDB 1.5.6.
- Invariant kept: a region stays internally acyclic, because within one column a row reads only earlier rows of
  itself (recurrence) or of earlier columns (growing sum), and `column_order` orders the columns.

## Limits (stay in cell mode, with a reason)

Other recurrences (`X{r-1}*1.05`, `MAX(X{r-1}, ...)`, `E - X{r-1}`), other functions over growing ranges, SUM with more
than one argument around a growing range, multi-column or other-sheet growing ranges, a run starting on row 1.
Floating point: the window adds in another order than Excel. Two orders of n terms differ by at most about
2 n eps sum|x|, so a column becomes a window only when Excel's saved values prove that bound is inside the verifier's
tolerance on every row (`region._order_safe`). Terms of one sign always pass; a balance that cancels back to near zero
(a ledger) stays in cell mode, which adds in Excel's order. (Found by the review: 20 FAILs on a 1,000-row ledger
without this check, none with it.)

## Testing

Each shape against cell mode (`regions=True` vs `False`, identical unresolved set and values): errors partway down
(both error rules), text and booleans in the source, empty / text / formula seed, range starting above and at the
region, a growing sum inside a larger formula, a rejected recurrence (`*1.05`); one real dbt build; the 200k
benchmark measured.
