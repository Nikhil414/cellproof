# Lookups row-wise

## Problem

A lookup pasted down a column (`=VLOOKUP(A2,$E$1:$F$1000,2,FALSE)`) was one query per cell. In a region its SQL
references the row's value, and DuckDB 1.5 cannot plan the correlated forms the lookups used (`ORDER BY ... LIMIT 1`,
`LEFT JOIN` on the row's values): "Cannot perform non-inner join on subquery", "Failed to cast expression to type".
`_validate` then sent the column to cell mode.

## Ruling

Every lookup search is an aggregate over a table that does not depend on the row; only the WHERE does:

    (SELECT arg_min(COALESCE(k.r_num, 0), k.row_num)
     FROM (SELECT k.*, r.kind AS r_kind, r.num AS r_num, r.txt AS r_txt FROM cells k LEFT JOIN cells r ON ...) k
     WHERE <area> AND k.kind = 'n' AND k.num = _.a)

DuckDB unnests this into one hash join (exact match) or one range join (approximate) for the whole region.

- First match: `arg_min(value, position)`; last (XLOOKUP search -1): `arg_max`.
- Approximate: `arg_max(value, (key, position))` for "largest <= x, last of equal", `arg_min(value, (key, position))`
  for "smallest >= x"; a negated position flips the tie-break. Same order as the old `ORDER BY`.
- `arg_min` skips NULL values, so every picked value is COALESCEd (an empty cell found reads 0 / '' / 'n' as before).
- Found / not found: `EXISTS`. INDEX: `max()` over the one addressed cell, COALESCEd.
- One form for cell mode and region mode, so the region-equals-cell-mode tests compare the same Excel rules.

## Measured

VLOOKUP into a 1000-row table: 1k rows verify 5.7 s -> 0.1 s; 50k rows 8.9 s total, all PASS. Cell mode (a lookup not
in a region) costs about 20% more per cell; accepted, since pasted-down lookups no longer use it.

## Criteria functions (same day)

Most of the *IF(S) family already planned row-wise. Two forms did not, and both are fixed:

- COUNTIFS built its driving rows in a `WITH o AS (...)` inside the subquery: DuckDB fails to bind it ("inequal types
  (VARCHAR != DOUBLE)"). Now a derived table.
- SUMIF/AVERAGEIF(S)/... over a sum range that can hold an error found the first error with `ORDER BY ... LIMIT 1`.
  Now `arg_min(txt, row_num)`.

Measured: COUNTIFS with two criteria, 1k rows, verify 29.0 s -> 0.1 s. No known formula now fails to plan row-wise;
the fallback test forces a failure with a patched `_validate`.
