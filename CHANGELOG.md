# Changelog

## Unreleased

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
