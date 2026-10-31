# Changelog

## 0.4.0 (2026-10-31)

- **Array formulas over a block** (legacy Ctrl+Shift+Enter blocks and dynamic-array spills) compile cell by cell:
  a range where one value is expected takes this cell's element, a range one row (column) high repeats down
  (across). Rules learned from formulas' Excel-saved test.xlsx: a block cell past the array is #N/A whatever the
  formula does with errors (`{=IFERROR(G9:Q9,R9)}`), but an IS function's own range is padded with #N/A elements
  (`{=ISERROR(C8:Q8)}` past Q8 is TRUE); EDATE, EOMONTH, MROUND and YEARFRAC do not take a range there (#VALUE!).
  A single-cell array formula is its first element (`{=MONTH(I6:J6)}`).
- **FILTER, SORT, UNIQUE** spill their results and nest in each other (`SORT(UNIQUE(FILTER(...)))`): FILTER by rows
  or columns with if_empty, SORT by one key (ascending or descending, equal keys keep their order), UNIQUE with
  exactly_once, text compared ignoring case. A dynamic spill (Excel's cm mark on the cell, now read) whose result
  is not exactly as big as the block Excel saved shows #SPILL! in its first cell, a FAIL. Not guessed (UNRESOLVED):
  an empty FILTER without if_empty (#CALC! in Microsoft's docs, #VALUE! in an Excel-saved file), SORT keys that
  are empty, errors or text with punctuation, UNIQUE over empty cells or errors.
- **Excel 365 error codes** (#CALC!, #SPILL!, #FIELD!, ...) are read as errors; functions stored `_xlfn._xlws.` are
  recognised.
- **Security:** an array block in a crafted file can no longer overwrite another cell's formula or create cells the
  file does not hold.
- **`A1#`** (ANCHORARRAY) is the block Excel spilled A1 into; **`@`** (SINGLE) is implicit intersection.
- **INDEX with a row or column of 0** is the whole column or row: element by element in an array formula,
  UNRESOLVED elsewhere (was a wrong #REF!).
- **Fixes:** SUBSTITUTE with an instance number and cell arguments stopped the run on DuckDB 2.0 (exit 3), and
  refuses TRUE as the instance like Excel; ISBLANK of a range read its first cell only.

## 0.3.1 (2026-10-31)

- **WORKDAY with a huge holidays range** (`WORKDAY(A1,5,I1:XFD1048576)`): the days it searched were sized by the
  range's area, so one workbook could make DuckDB fail or exhaust memory. The search now covers 2 days per holiday
  actually in the range, never past Excel's last date.

## 0.3.0 (2026-10-31)

- **Finance:** PMT, PV, FV, NPER, NPV. Microsoft's closed forms; NPV skips text, TRUE/FALSE and empty cells in a
  reference without spending a period on them (checked against pycel's Excel-saved npv.xlsx).
- **Statistics:** LARGE, SMALL, RANK(.EQ), STDEV(.S/.P/P), VAR(.S/.P/P), PERCENTILE(.INC), QUARTILE(.INC). Real Excel
  (formulas' test.xlsx) taught two rules: k's error comes before an error in the range, and QUARTILE checks quart
  before truncating it (QUARTILE(x, -0.1) is #NUM!).
- **Working days:** NETWORKDAYS and WORKDAY, with a holidays range.
- **Excel tables:** structured references (`Sales[Amount]`, `Sales[@Qty]`, `Sales[[#This Row],[Qty]]`, `#All`,
  `#Data`, `#Headers`, `#Totals`, `Sales[[Rep]:[Amount]]`) are rewritten to A1 references by the reader; a
  calculated table column compiles as a region.
- **LET**, nested LETs and a name bound to a range (`LET(r, A1:A9, SUM(r)/COUNT(r))`) included.
- **Not guessed:** edge cases of the new functions that were never checked in real Excel (zero periods, rate <= -100%,
  payment type other than 0/1, a k that is not whole, ...) are UNRESOLVED with that reason.
- **Fixes found by the public corpus** (FAIL 92 -> 16, all 16 explained: 10 in LibreOffice files, 6 last-digit
  cancellations around NPV): INDEX with a negative row or column is #VALUE!, not #REF!; SWITCH reads an empty cell as
  0 (it matches 0, not FALSE or ""); a range where one value is expected inside an array formula is UNRESOLVED
  instead of implicitly intersected (Excel evaluates it per element there).
- **Array arithmetic in aggregates**, the first slices of array evaluation: `SUM(A1:A10*B1:B10)`,
  `SUMPRODUCT((A1:A9>5)*C1:C9)`, `MAX(IF(A1:A9="x",B1:B9))`, `SUM(IFERROR(A1:A9/B1:B9,0))`, `COUNTA(A1:A9&B1:B9)`,
  `SUM(A1:A9^2)` and friends evaluate element by element like modern Excel - an empty element is 0, text converts
  when it is a number and is #VALUE! otherwise, comparisons rank numbers < text < TRUE/FALSE case-insensitively,
  concatenation writes numbers the Excel way, and the first error wins. SUM, SUMPRODUCT, MAX, MIN, AVERAGE, COUNT,
  COUNTA and PRODUCT take arrays; IF, IFERROR and IFNA take them as their arguments. A single-cell array formula
  (Ctrl+Shift+Enter) now compiles; one over several cells is UNRESOLVED. Plain-entered range arithmetic or IF as an
  aggregate argument follows modern Excel (the element-wise answer), not the legacy first-row intersection - a
  workbook an old Excel last saved can show FAILs there, with both values shown.
- **Approximate lookups on unsorted data** (MATCH type 1, VLOOKUP/HLOOKUP approximate, LOOKUP) over a plain numeric
  input table: cellproof now compiles Excel's binary-search path itself - the answer is the last probe the search
  moves right from, not the sorted bracket, and it matches Excel's own answers on files we checked (the 9.99E+307
  last-number trick included). The SQL refuses table values that changed after translation (#STALE, UNRESOLVED)
  instead of answering from a stale search. MATCH type -1, text, and tables with empty cells or formulas stay
  UNRESOLVED: Excel's answer there is not derivable from the file.
- **Where to start:** the report groups the unresolved cells by the reason to fix first and ranks the reasons by
  reach: a cell that reads an unresolved cell is attributed to the reason above it, so the top row of
  "Where to start" names the fix that lets the most cells try again. report.json gains `unresolved_groups` and a
  `blocked_by` root per unresolved cell.

## 0.2.2 (2026-10-27)

- **LOOKUP** (vector and array form). Like the other approximate lookups, UNRESOLVED on data out of order.
- **YEARFRAC**, all five bases, the basis written in the formula or computed. The 1,992 YEARFRAC cells of pycel's
  test workbook all PASS; a date before 1 March 1900 (Excel's 29 February 1900) is UNRESOLVED.
- **Array constants** (`{0,50,100}`, `{1,"a";2,"b"}`) wherever a range is taken: `LOOKUP(x,{0,50,100},{"F","C","A"})`,
  `MATCH(x,{...},0)`, `SUM({...})`. Used as one value (`A1*{1,2}`) they stay UNRESOLVED.

## 0.2.1 (2026-10-20)

Found by running all 85 public test workbooks of pycel, formulas and xlcalculator (`python tests/corpus.py DIR`):
FAIL 229 -> 10 (all 10 in files LibreOffice saved, which the report now says), no crashes.

- **Reading:** a defined name that is a formula of names (`=h+i`) no longer crashes the load; a date-formatted number
  keeps its serial (Excel's 29 February 1900 and date-formatted negatives were lost); text escaped by the file format
  (`_x000D_`, written for a carriage return) is decoded as Excel does.
- **AND/OR/XOR** skip text and empty cells in references, and typed text other than TRUE/FALSE, like Excel.
- **An empty or omitted result is 0:** `IF(c,,"x")`, `IFERROR(A1,"x")` with A1 empty.
- **EDATE/EOMONTH/MROUND/VALUE** refuse TRUE/FALSE (`#VALUE!`); `EOMONTH(0,0)` is 31.
- **SUMPRODUCT** of ranges of different sizes returns an error inside them before `#VALUE!`.
- **PROPER** accepts another cell's value (DuckDB refused the SQL) and capitalizes accented letters.
- **Locale-dependent text** (`"1,000"`, `"1/2/2020"`) is UNRESOLVED instead of a wrong `#VALUE!`.
- **Report:** warns when the file was not last saved by Excel; CHAR above 127 in a Mac-calculated file is UNRESOLVED.

From a user's practice workbook (3 PASS / 47 UNRESOLVED -> 48 PASS / 2 UNRESOLVED):

- **Implicit intersection:** a one-column or one-row range where one value is expected takes the formula's own row or
  column (`=IF(E$3:E$41="F",B$3:B$41)`); outside it is `#VALUE!`. A range IF passes into another function is
  UNRESOLVED (Excel gives that function the whole range).
- **`INDIRECT("Sales")`** of fixed text is the reference it names (dependencies included).
- **MODE / MODE.SNGL.**
- **A bare word is a name, never a column:** `=i` without a name `i` was read as column I.

## 0.2.0 (2026-10-19)

- **Renamed** from `unexcel` to `cellproof` (the old name clashed with an R package): `pip install cellproof`, command
  `cellproof convert`, import `cellproof`, `CELLPROOF_PERF` for the benchmark, dbt project and database `cellproof`.
- **Lookups row-wise:** VLOOKUP, HLOOKUP, MATCH, INDEX and XLOOKUP pasted down a column compile into one statement.
  1,000 VLOOKUPs: verify 5.7 s -> 0.1 s.
- **Criteria row-wise:** COUNTIFS (and SUMIF/AVERAGEIF(S) over a range holding an error) with a criterion from the row
  no longer fall back to cell mode. 1,000 COUNTIFS: verify 29 s -> 0.1 s.
- **Faster compile:** a pasted-down formula is tokenized once per run. 200k-formula benchmark 31 s -> 26 s; 50k
  VLOOKUPs 8.9 s -> 4.6 s.
- A lookup outside a region is about 20% slower than in 0.1.0 (one SQL form for both modes).

## 0.1.0 (2026-10-18)

First release, as `unexcel`.
