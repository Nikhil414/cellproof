# cellproof: progress

Last updated: 2026-10-20.

## State

Everything below is on `main`, built test-first, each piece reviewed by a fresh reviewer.

| Piece | What it does |
|---|---|
| Plan 1 | reader, parser, references, dependency graph, cell-mode compiler, verifier, dbt emitter, report, CLI |
| Plan 2a | batched execution (fixes the crash at ~1k formulas), Excel error codes, recursion and size guards, optional `formulas` answer key |
| Plan 2b | 93 Excel functions, run-time types, Excel's conversions, one criteria engine for the *IF(S) family |
| Plan 2c | speed in cell mode: Fenwick-tree dependency order, range index, weight-based batches, small dbt models |
| Plan 2d | table mode slice 1: regions (columns repeating one formula) compile to one readable statement |
| Slice 2 | a region with a cell between its columns splits instead of dissolving; running totals (`=C1+B2`, `=SUM($B$2:B2)`) become window functions |
| Lookups row-wise | VLOOKUP, HLOOKUP, MATCH, INDEX, XLOOKUP pasted down compile into the region's statement (one join each instead of one query per cell) |
| Tokenize once | a formula pasted down a column is tokenized once; the next rows' text is predicted and only re-tokenized when it differs |
| Implicit intersection | a one-column range where one value is expected takes the formula's row; `INDIRECT` of fixed text; MODE (found in a user's workbook) |
| Criteria row-wise | the *IF(S) family with a criterion from the row (`=COUNTIFS($E:$E,">"&A2,...)`) compiles into the region's statement too |
| Impact-ranked unresolveds | the report groups UNRESOLVED cells by the reason to fix first ("Where to start"): the compiler records which unresolved cell stopped a compile (`Unsupported.dep`), each chained cell is attributed to the first blocker its chain reaches, reasons are ranked by reach; report.json gains `unresolved_groups` and per-cell `blocked_by` |
| Approximate lookups unsorted | MATCH type 1 / VLOOKUP / HLOOKUP / LOOKUP over a plain numeric input table out of order compiles Excel's binary-search path (a decision tree of probes, depths and x-intervals, one aggregate in SQL; verified differentially against the bisection and against real Excel answers). A changed table is #STALE -> UNRESOLVED, never a stale answer. Type -1 and text stay UNRESOLVED (Excel's rule there is not derivable; LibreOffice differs). examples/functions.xlsx: 4 UNRESOLVED -> 2 |
| Arrays slice 1 | element-wise arrays in the aggregates: `SUM(A1:A10*B1:B10)`, `SUMPRODUCT((A1:A9>5)*C1:C9)`, `MAX(A1:A9-1)`. An array argument is one relation with the cells table's columns (a row per element position, blanks and all), which SUM/PRODUCT/MAX/MIN/AVERAGE/COUNT/COUNTA/SUMPRODUCT read like a range. Text element -> #VALUE!, empty -> 0, first error wins. SUM moved from functions.py onto the aggregate machinery. A single-cell CSE formula compiles (`Cell.ref` kept from the file); multi-cell stays UNRESOLVED. Next: IF over ranges (`MAX(IF(...))`), & and ^ on arrays, text comparisons |
| Open source | MIT license, README with limits, SECURITY.md, CI on Windows/macOS/Linux x Python 3.10-3.13 |

Tests: about 437 (dbt build included), plus 3 workbooks saved by real Excel that all PASS (4 cells UNRESOLVED on
purpose). Public on GitHub (one-commit history start). 0.1.0 was released as `unexcel`; 0.2.0 (2026-10-19) and 0.2.1 (2026-10-20) are on PyPI as `cellproof` (see CHANGELOG.md).
A user's practice workbook: 3 PASS / 47 UNRESOLVED -> 48 PASS / 0 FAIL / 2 UNRESOLVED.

## Public test workbooks (2026-10-19)

`python tests/corpus.py DIR` runs the 85 .xlsx/.xlsm test files of pycel, formulas and xlcalculator (Excel-saved,
except 8 LibreOffice ones): 6,874 PASS, 10 FAIL (all in LibreOffice files, flagged in their reports), 16,764
UNRESOLVED, no crashes (was 4,693 PASS before LOOKUP, YEARFRAC and array constants). The UNRESOLVED are mostly functions those files test on purpose: CONVERT, YEARFRAC,
DEC2HEX/BIN/OCT, array formulas and `{1,2}` constants, CODE/UNICODE, LOOKUP. For business workbooks the ones that
matter are array formulas, array constants, LOOKUP and YEARFRAC. Findings are tests in `tests/test_corpus_findings.py`.

## Measured speed (Windows 11 laptop)

| Workbook | Before table mode | Now |
|---|---|---|
| 200k formulas (4 columns x 50k rows, one running total) | 390 s | 26.0 s (read 8.3, compile 11.3; was 31.0 s with compile 19.9), all PASS |
| Flat `=A*1.1`, 20k cells | verify 28.3 s | verify 0.2 s |
| `--dbt`, flat 20k | 225.9 s, 213 MB | 9.6 s, 1.6 MB |
| VLOOKUP per row, 1k rows | 5.9 s | verify 0.1 s (was 5.7 s), 1 batch |
| VLOOKUP per row, 50k rows | - | 4.6 s (read 2.2, compile 1.8, verify 0.6), all PASS |
| COUNTIFS per row (`">"&A2`, 2 criteria), 1k rows | verify 29.0 s, 500 batches | verify 0.1 s, 1 batch |
| SUMIFS / COUNTIFS per row, 50k rows | - | verify 0.6-0.8 s; compile 8.7-9.5 s |

## Real Excel validation (2026-10-14)

| Workbook | Result |
|---|---|
| `examples/commission.xlsx` (4 formulas) | all PASS |
| `examples/errors.xlsx` (15) | all PASS |
| `examples/functions.xlsx` (76) | all PASS; 4 UNRESOLVED on purpose (approximate MATCH on unsorted data) |

What Excel taught us (each has a test in `tests/test_excel_findings.py`):

- COUNTIF `"<>5"` skips only the number 5 (the text "5" is counted); `"<>"` counts a formula that returns "".
- DAYS360 (US): only a 31st is moved; the last day of February ends the period unchanged unless the start is also the
  last day of February.
- `TEXT(-0.001,"0.00")` is `"0.00"`. `MOD(1,0.1)` is `0.09999999999999995` (DuckDB's `%`, not `fmod()`). Inside 1900
  a date is linear: `DATE(1900,2,29)`, `DATE(1900,1,60)` and `DATE(1900,3,0)` are all 60.
- Number -> text is plain digits from 1E-11 to 1E+19 and scientific from 1E+20.
- Errors typed into a formula come before errors inside referenced cells: `SUM(H3,1/0)` is #DIV/0!.
- `(-8)^(1/3)` is the real cube root (-1.9999999999999998); other fractional powers of a negative number are #NUM!.
- Approximate MATCH/VLOOKUP/HLOOKUP on unsorted data is UNRESOLVED; a model of Excel's binary search predicted all 4
  Excel answers, so it could be emulated later.
- `XLOOKUP` must be stored as `_xlfn.XLOOKUP` in the file.

## Known gaps

- A lookup outside a region (cell mode) is about 20% slower than before the row-wise form (1k cells: 6.9 s, was 5.8 s).
- Running totals other than sums (`=C1*1.05+B2`, `=MAX(C1,B2)`), other functions over growing ranges, and a split
  piece that needs a second split stay in cell mode.
- Compile on the 200k run is 11.3 s: the dependency-order loop and reference resolving, one pass per cell. Next step
  there would be ordering a region's cells per formula shape instead of per cell. Read (8.3 s, openpyxl) is the other half.
- Minors: a DuckDB error in one batch stops the run (cleanly: exit 3 naming the cell; none known); overflow gives `inf`, not `#NUM!`; workbook text equal to the
  placeholder `__cellproof_cells__` is rewritten; computed-answer results are not counted separately; old tables stay in
  `cellproof.duckdb` after a rerun.
- Backlog for one more Excel round: number -> text below 1E-11, `(-8)^(2/3)`, `(-16)^(1/4)`, `DATE(1900,13,1)`.

## Next

1. Array evaluation, next slices: IF over ranges (`MAX(IF(...))`, the classic conditional aggregate), & and ^ on
   arrays, text comparisons per Excel's collation, then legacy multi-cell array formulas and dynamic arrays.
   (Done: element-wise arithmetic and comparisons inside the aggregates - see Arrays slice 1 above.)
2. Test on more real analyst workbooks (one user workbook found 4 gaps the 85 public files did not).
3. Read time (openpyxl loads the workbook twice) and the per-cell dependency loop.
4. Plan 3: AI retry loop for failed cells (it sees formula structure only, never values).
