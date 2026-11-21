# cellproof

[![PyPI](https://img.shields.io/pypi/v/cellproof)](https://pypi.org/project/cellproof/)
[![tests](https://github.com/Nikhil414/cellproof/actions/workflows/ci.yml/badge.svg)](https://github.com/Nikhil414/cellproof/actions/workflows/ci.yml)
[![Python](https://img.shields.io/pypi/pyversions/cellproof)](https://pypi.org/project/cellproof/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)

**Turn an Excel workbook into SQL, and prove every number matches.**

<img src="docs/demo.gif" alt="cellproof converting a workbook: every formula checked against Excel's saved value" width="320" align="right">

**[Try it in your browser](https://nikhil414.github.io/cellproof/)**: drop a workbook, get every verdict and the dbt
project back. It runs entirely in the page (Python and DuckDB in WebAssembly); your file is never uploaded.

**Or on your machine, in 30 seconds:**

```bash
pip install cellproof
curl -LO https://github.com/Nikhil414/cellproof/raw/main/examples/functions.xlsx
cellproof convert functions.xlsx
# 74 pass, 0 fail, 2 unresolved -> functions_dbt/REPORT.md
# Where to start (unresolved, ranked by how many cells each blocks):
#   2  MATCH approximate match needs descending sorted data: ...  (Functions!G13, Functions!G73)
```

Then open `functions_dbt/REPORT.md`: every formula, the value Excel saved, the value the SQL gave, and why the two
unresolved cells were not guessed.

- **11,080 cells PASS** on 85 public test workbooks (pycel, formulas, xlcalculator). The 16 FAILs are 10 cells in files LibreOffice saved and 6 last-digit NPV
  cancellations; none is a wrong translation.
- **152 Excel functions**, array formulas, dynamic spills (FILTER, SORT, UNIQUE), Excel tables, LET.
- **Fast:** 200,000 formulas converted and verified in 26 s.
- **Local:** workbook values never leave your machine.

Have a workbook cellproof gets wrong? [Open an issue](https://github.com/Nikhil414/cellproof/issues/new) with the
formula (not your data): that is the most useful contribution there is.

(Released as `unexcel` up to 0.1.0; renamed to cellproof in 0.2.0.)

<br clear="right">

Business logic often lives in a spreadsheet that one person understands. Moving it into a database is risky because
nobody can show the new numbers equal the old ones. cellproof translates every formula of an `.xlsx` into SQL, runs that
SQL on [DuckDB](https://duckdb.org), and compares each result with the value Excel itself saved in the file. Excel's
saved values are a free answer key.

```
$ cellproof convert examples/functions.xlsx
74 pass, 0 fail, 2 unresolved -> examples/functions_dbt/REPORT.md
Where to start (unresolved, ranked by how many cells each blocks):
  2  MATCH approximate match needs descending sorted data: ...  (Functions!G13, Functions!G73)
```

Every formula cell gets one verdict:

| Verdict | Meaning |
|---|---|
| **PASS** | the SQL gives the value Excel saved |
| **FAIL** | the SQL gives a different value; the report shows both |
| **UNRESOLVED** | cellproof could not translate the formula exactly; the report says why. It never guesses. |

Exit code: `0` no FAIL, `1` at least one FAIL (or `dbt build` failed), `2` the workbook could not be read, `3`
DuckDB rejected SQL that cellproof generated (a bug in cellproof; the message names the cell, please report it).

## Install

Python 3.10 or newer.

```bash
pip install cellproof
pip install "cellproof[dbt]"      # also run dbt build (--dbt)
pip install "cellproof[compute]"  # answer key for files never saved in Excel
pip install "cellproof[postgres]"                 # verify on Postgres too (--target postgres)
pip install "cellproof[postgres,dbt-postgres]"    # ... and run the dbt-postgres project (--dbt)
```

Runtime dependencies are only `openpyxl` and `duckdb`.

## Use

```bash
cellproof convert book.xlsx                 # writes book_dbt/ next to the file
cellproof convert book.xlsx --out out_dir   # choose the folder
cellproof convert book.xlsx --dbt           # also run `dbt build` on the generated project
cellproof convert book.xlsx --stats         # print the unresolved cells' reasons, most common first
cellproof sql book.xlsx --sheet Sales --range D2:D1000   # the SQL for those cells only
cellproof watch book.xlsx                  # verify again on every save; print only what changed
cellproof prove book.xlsx                  # verify, then write a bundle someone else can check
cellproof convert book.xlsx --target postgres            # verify on Postgres too (see below)
cellproof convert book.xlsx --fast         # read the workbook in one pass (same verdicts, less time)
```

After the counts line, the console shows what to fix first: the top unresolved reasons ranked by how many cells each
blocks, and any wrong answers other formulas read (direct readers only, so a difference that travels is visible).
`REPORT.md` and `report.json` hold the full lists.

The output folder holds:

- **`REPORT.md` / `report.json`**: every verdict, failed cells with both values, unresolved cells with the reason,
  findings (for example a number stored as text, which Excel skips in SUM), and which blocks of rows became readable
  table models. The unresolved reasons are ranked by how many cells depend on them ("Where to start"), so a long list
  of unresolved cells says what to fix first. Failures are split: "Failures to look at" lists the ones whose own
  inputs were all right (the difference starts there), and a line says how many others read one of them and may be
  right themselves. A failed cell that other formulas read is also listed with its readers ("Wrong answers other cells
  read"), because a wrong value that feeds other cells can spread. When the file itself says its saved values may be
  out of date (manual calculation, a full recalculation pending, no `calcId`, or last saved by a program that is not
  Excel), `REPORT.md` says so above the tables: recalculate in Excel and run again before believing a FAIL.
- **A dbt project** (`dbt_project.yml`, `models/`, `seeds/input_cells.csv`) that recomputes the workbook on DuckDB.
  It runs exactly the SQL the verifier checked.

### One cell's answer, as SQL

`cellproof sql` prints a runnable script instead of writing a project: paste it into `duckdb`, or pipe it.

```bash
cellproof sql book.xlsx --sheet Sales --range D2:D1000 > answer.sql
duckdb < answer.sql
cellproof sql book.xlsx --out answer.sql --stats    # write the file, say how much of the workbook it needed
```

The script creates the `cells` table, inserts the input cells those formulas read, adds one statement per formula
batch in dependency order, and ends with a `SELECT` for the cells you asked for. Everything the selection does not
need is dropped: with `--range D2` on `examples/functions.xlsx`, all 74 formulas and 16 statements of the workbook
become one row and no formulas at all; asking for `G20:G30` keeps 26 formulas in 4 statements.

It answers exactly what `cellproof convert` would compute for those cells. A cell that could not be translated is not
in the script (nothing could read it anyway); the reason is printed to stderr, ranked, and the command exits 1, so a
piped script never quietly loses a cell. Without `--sheet` the whole workbook is the selection (hidden array-constant
sheets are left out).

### A result you can hand over

`cellproof prove book.xlsx` verifies the workbook exactly as `convert` does, then writes the result as a folder
someone else can check without trusting you or cellproof:

```bash
cellproof prove book.xlsx                                # writes book_proof/ next to the file
cellproof prove book.xlsx --out bundle                   # choose the folder
cellproof prove --check bundle                           # re-check a bundle: hashes, then the SQL itself
cellproof prove --check bundle --workbook book.xlsx      # ... and re-read the workbook it names
```

The folder holds `cells.sql` (one runnable DuckDB script: the input cells, every formula batch in dependency order,
and a final `SELECT` for every cell), `report.json` and `REPORT.md` (what `convert` would have written), and
`proof.json` (the workbook's size and hash, the cellproof/DuckDB/Python versions and the platform, the workbook's own
`<calcPr>` and last-saving program, the counts, and a hash of each file plus one over the set).

Anyone with DuckDB can check the claim by hand: run `cells.sql`, compare each row with `report.json`, hash the
workbook. `--check` does the same thing and says how it went:

```
3 files intact
workbook hash matches
4 saved values match the report
cells.sql reproduces 4 of 4 answers
the bundle checks out
```

Exit code: `0` every check held, `1` an answer moved (the SQL disagrees with the report, or the workbook no longer
holds the values the report records), `2` a file, a hash or the format is wrong, `3` DuckDB rejected SQL cellproof
generated.

The seal is **not a signature**. The hashes catch a file changed by accident or by a casual edit; they cannot stop
someone with write access to the folder, who can reseal it. What that person still cannot fake is `cells.sql`
disagreeing with `report.json`, which is why `--check` runs the script instead of stopping at the hashes. `REPORT.md`
says this in one line under its title. A cell that could not be translated is a hole in `cells.sql`, not a wrong
answer: `report.json` says why. `cells.sql` is about 10 MB for 200,000 formula cells - no compression, because a proof
should be readable.

### Editing while watching

`cellproof watch book.xlsx` verifies the workbook once, then verifies again on every save (it watches the file's
modification time and size, `--interval` seconds between checks, Ctrl+C to stop) and prints only what changed:

- **broke / fixed**: a cell's verdict changed between PASS and FAIL or UNRESOLVED, with both values or the reason.
- **still FAIL**: the cell was already failing and the SQL answer moved.
- **changed**: a FAIL became UNRESOLVED or the other way round (the translation changed what can be said, not the
  answer).
- **new cells / cells gone**: rows added to or removed from a formula column.
- **new unresolved reasons**, counted, so a new unsupported function is visible at once.

The first run also prints the counts and where to start, like `convert` does on the console. A workbook that cannot
be read is not an error while you watch: Excel may still be writing it, the last result is kept and the next save is
another chance.

### Reading faster

`--fast` (on `convert`, `sql` and `watch`) reads each worksheet once instead of twice. Nothing new to install, no new
dependency, and no change to what you get out:

```bash
cellproof convert book.xlsx --fast
```

On a 200,000-formula workbook the read goes from 2.40 s to 1.35 s and a whole `convert` from 5.32 s to 4.45 s; the
85 public test workbooks together read in 0.51 s instead of 0.85 s. The verdicts are identical - loading is now only
about a quarter of a conversion, so the rest of the time is compiling and verifying, not reading.

The fast loader is the standard library's XML parser plus openpyxl's own parsers for everything except the cell loop,
and `tests/test_fastread.py` checks it cell-for-cell against the openpyxl load: values, kinds, formulas, arrays,
names, tables, merges, hyperlinks, the Mac 1904 epoch, and the errors a corrupt file raises. If it meets something it
does not handle it does not guess: it prints one note to stderr and the file is read with openpyxl instead, so
`--fast` can only ever be slower.

### On Postgres

```bash
cellproof convert book.xlsx --target postgres                      # on a Postgres of cellproof's own (pgserver)
cellproof convert book.xlsx --target postgres --pg-url postgresql://user:pw@host/db
cellproof convert book.xlsx --target postgres --dbt                # and run dbt build on that same Postgres
```

Every cell is verified twice: on DuckDB, then again on Postgres with the SQL translated (sqlglot plus rewrites and
helper functions that give Postgres DuckDB's semantics). The report's verdicts are Postgres's. A batch that does not
translate or run on Postgres makes its cells UNRESOLVED ("not yet on Postgres: ..."), and every cell reading them too,
so the dbt project only holds SQL that ran on Postgres. Nothing stays in the database you verify on: the cells table
is temporary and the helpers live in `pg_temp`.

The dbt project targets `dbt-postgres`. Its helper functions are created in the target schema by an `on-run-start`
macro, and `profiles.yml` reads the connection from `CELLPROOF_PG_HOST`, `CELLPROOF_PG_PORT`, `CELLPROOF_PG_USER`,
`CELLPROOF_PG_PASSWORD`, `CELLPROOF_PG_DBNAME` and `CELLPROOF_PG_SCHEMA` (default: the workbook's name).

The workbook must have been saved by Excel (so the file holds Excel's values). For a file generated by code and never
opened in Excel, install the `compute` extra: the [`formulas`](https://pypi.org/project/formulas/) library then
computes the expected values, and the report labels them `computed` (a weaker answer key than Excel).

## How it works

```
reader -> regions -> dependency graph -> compiler -> verifier -> dbt emitter -> report
```

1. **Read** the workbook with openpyxl: twice for the formulas and the values Excel saved, or once with `--fast`
   (see "Reading faster"). Dates become Excel serial numbers.
2. **Find regions.** A block of rows where columns repeat one formula down (`=A2*1.1`, `=IF(B2>100,B2,0)`) is a
   *region*. Running totals in a region (`=C1+B2`, `=SUM($B$2:B2)`) become SQL window functions. A region compiles to
   one readable statement (`amount * 1.1 AS tax`) instead of one statement per cell.
3. **Order** every formula after the cells it reads. Circular references are UNRESOLVED.
4. **Compile** each formula to SQL with Excel's rules: Excel's type conversions (`"3"+1 = 4`), its error codes
   (`#DIV/0!`, `#N/A`, `#VALUE!` ...) which must match exactly, the left-to-right order in which Excel reports errors,
   and DuckDB guards (division by zero, `ln(0)`, overflow of `POWER`) so DuckDB never fails where Excel returns an
   error. Cells outside regions run in batches ("cell mode").
5. **Verify** on DuckDB in memory: numbers must match within a relative 1e-9, text and booleans exactly, errors by
   code.
6. **Emit** the dbt project and the report.

152 Excel functions are supported: ABS, AND, AVERAGE, AVERAGEIF, AVERAGEIFS, BIN2DEC, BIN2HEX, BIN2OCT, CEILING, CHAR,
CHOOSE, CLEAN, CODE, COLUMN, COLUMNS, CONCAT, CONCATENATE, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, DATE, DATEDIF,
DAY, DAYS, DAYS360, DEC2BIN, DEC2HEX, DEC2OCT, EDATE, EOMONTH, EXACT, EXP, FALSE, FILTER, FIND, FLOOR, FV, HEX2BIN,
HEX2DEC, HEX2OCT, HLOOKUP, IF, IFERROR, IFNA, IFS, INDEX, INT, IRR, ISBLANK, ISERR, ISERROR, ISEVEN, ISLOGICAL, ISNA,
ISNUMBER, ISODD, ISTEXT, LARGE, LEFT, LEN, LET, LN, LOG, LOG10, LOOKUP, LOWER, MATCH, MAX, MAXIFS, MEDIAN, MID, MIN,
MINIFS, MOD, MODE, MODE.SNGL, MONTH, MROUND, N, NA, NETWORKDAYS, NOT, NPER, NPV, OCT2BIN, OCT2DEC, OCT2HEX, OR,
PERCENTILE, PERCENTILE.INC, PI, PMT, POWER, PRODUCT, PROPER, PV, QUARTILE, QUARTILE.INC, RANK, RANK.EQ, RATE, REPLACE,
REPT, RIGHT, ROUND, ROUNDDOWN, ROUNDUP, ROW, ROWS, SEARCH, SEQUENCE, SIGN, SINGLE (the @ operator), SMALL, SORT,
SORTBY, SQRT, STDEV, STDEV.P, STDEV.S, STDEVP, SUBSTITUTE, SUBTOTAL, SUM, SUMIF, SUMIFS, SUMPRODUCT, SWITCH, TEXT,
TEXTJOIN, TRIM, TRUE, TRUNC, UNICHAR, UNICODE, UNIQUE, UPPER, VALUE, VAR, VAR.P, VAR.S, VARP, VLOOKUP, WEEKDAY,
WORKDAY, XLOOKUP, XMATCH, XOR, YEAR, YEARFRAC.

**Excel tables:** structured references (`Sales[Amount]`, `Sales[@Qty]`, `Sales[[#This Row],[Qty]]`, `Sales[#Totals]`,
`Sales[[Rep]:[Amount]]`) are rewritten to plain references when the file is read, so a calculated table column
compiles like any column pasted down.

**Speed** (Windows 11 laptop, measured): a 200,000-formula workbook (four formula columns over 50,000 rows, one of them
a running total) converts and verifies in 26 s, all PASS. 50,000 rows of VLOOKUP into a 1,000-row table: 4.6 s.
COUNTIFS/SUMIFS with a criterion from the row run at the same speed. The read is a quarter of that time and `--fast`
cuts it by about 44% (measured on a 2024 laptop: 5.32 s -> 4.45 s for the whole conversion).

## Limitations

cellproof prefers an honest UNRESOLVED or FAIL to a guess. Known limits, each reported with a reason:

- **File formats:** `.xlsx` and `.xlsm` only (no `.xls`, `.xlsb`, `.ods`). Macros (VBA) are ignored.
- **Array constants** (`{0,50,100}`) work wherever a range is taken (`SUM`, `MATCH`, `LOOKUP`, `VLOOKUP`, `INDEX`...),
  inside array arithmetic and in array formulas; where one value is expected in a plain formula (`A1*{1,2}`) they are
  UNRESOLVED.
- **Array formulas over a block and dynamic arrays** (legacy Ctrl+Shift+Enter blocks, Excel 365 spills): every cell of
  the block is verified. A range where one value is expected takes the cell's element (one row or column repeats);
  past the array a cell is #N/A, while an IS function's own range is padded per element, as real Excel does. `A1#`
  and `@` are read. **FILTER, SORT and UNIQUE** spill their results and nest in each other; a dynamic spill whose
  result is not exactly as big as the block Excel saved shows #SPILL! in its first cell (a FAIL). The generated SQL
  computes the block Excel saved: if the data later changes the size of a FILTER's result, re-run cellproof.
  SORTBY, SEQUENCE and XMATCH too, and all of them inside aggregates (`SUM(FILTER(...))`, `ROWS(UNIQUE(...))`).
  Not yet: TAKE/DROP/CHOOSECOLS and the other array functions, these functions inside other non-aggregate
  calculations, SORT by several keys or by columns (SORTBY does both), SORT/SORTBY keys that are empty, errors or text
  with punctuation (Excel's own text order), UNIQUE over empty cells or errors - all UNRESOLVED.
- **Array arithmetic in aggregates** (`SUM(A1:A10*B1:B10)`, `SUMPRODUCT((A1:A9>5)*C1:C9)`, `MAX(IF(A1:A9="x",B1:B9))`,
  `SUM(IFERROR(A1:A9/B1:B9,0))`, `COUNTA(A1:A9&B1:B9)`, `SUM(A1:A9^2)`) follows modern Excel: element by element, an
  empty element is 0, text converts when it is a number ("3" -> 3) and is #VALUE! otherwise, comparisons rank
  numbers < text < TRUE/FALSE case-insensitively, concatenation writes numbers the Excel way, and the first error
  wins. A plain-entered range arithmetic or IF as an aggregate argument follows modern Excel,
  not the legacy first-row intersection - a workbook an old Excel last saved can show FAILs there, with both values
  shown.
- **Not supported:** volatile functions (`NOW`, `TODAY`, `RAND`, `OFFSET`, and
  `INDIRECT` of text built while the sheet runs; `INDIRECT("Sales")` is fine), functions outside the list above,
  external workbooks, a structured reference naming several separate columns (`T[[A],[C]]`), INDEX used as a reference, the 1904 date system in
  TEXT date formats.
- **Ranges where one value is expected** (`=IF(E$3:E$41="F",B$3:B$41)`) take the formula's own row or column, as
  Excel does (implicit intersection). As an aggregate's argument the element-wise answer is compiled instead (see
  array arithmetic above); in an array formula, element by element (see array formulas above).
- **Approximate lookups on unsorted data** (LOOKUP, MATCH/VLOOKUP/HLOOKUP with match type 1): compiled for plain
  numeric tables written as inputs (up to about 1,000 values): the SQL carries Excel's binary-search path itself, and
  refuses changed table values instead of answering from a stale search. MATCH with type -1, text tables, tables with
  empty cells or formulas, and bigger tables stay UNRESOLVED, because Excel's answer there is not derivable from the
  file (LibreOffice, the only open source that emulates the search, provably differs from Excel in places).
- **Text to number:** only plain numbers convert (`"3"`, `" -1.5e3 "`). Text Excel reads with the computer's language
  settings (`"1,000"`, `"$5"`, `"50%"`, `"1/2/2020"`, `"2:00 PM"`, `"Jan 5"`) makes the cell UNRESOLVED, because the
  file does not record those settings.
- **Who saved the file:** the report warns when it was last saved by another program (LibreOffice, a script): the
  saved values then follow that program's rules. In a file Excel for Mac calculated, `CHAR` above 127 is UNRESOLVED
  (the Mac character set differs).
- **Edge cases not checked in real Excel yet** are UNRESOLVED with that reason, never guessed: PMT/PV/FV/NPER with
  zero periods, a rate of -100% or less or a payment type other than 0 or 1; NPV at a rate of -100%; LARGE/SMALL with
  a k that is not whole; RANK over a range holding an error; NETWORKDAYS/WORKDAY with text or errors among the
  holidays. NETWORKDAYS.INTL, WORKDAY.INTL and the other finance and statistics functions outside the list above are not in
  yet.
- **Some function options:** SUBTOTAL 101-111 (hidden rows), DATEDIF units MD/YM/YD, WEEKDAY types 11-17, TEXT formats
  outside the supported table, COUNT or MEDIAN of values typed into the formula, criteria taken from an empty cell.
- **Floating point:** a running total becomes a fast window function only when Excel's saved values prove the different
  addition order cannot matter; balances that cancel back to zero (ledgers) are added in Excel's order instead, which
  is slower. Elsewhere SQL can still differ from Excel in the last digits, which would show as an honest FAIL. A
  number that overflows becomes `inf` instead of `#NUM!`.
- **Speed:** regions cover columns that repeat one formula over the same rows, lookups and *IF(S) criteria included.
  A formula that is not part of such a column runs cell by cell (a lookup alone: about 7 s per 1,000). Running totals
  other than plain sums (`=C1*1.05+B2`, `=MAX(C1,B2)`) also run cell by cell.
- **The answer key:** a workbook must have been saved by Excel; `formulas` (the `compute` extra) is a weaker fallback
  and is known to disagree with Excel in places.
- **dbt:** the generated project targets DuckDB (`dbt-duckdb`), or Postgres (`dbt-postgres`) with `--target postgres`.
  Other warehouses are not generated yet.

## Security and privacy

- **No network.** cellproof never sends workbook data anywhere; everything runs locally in DuckDB.
- **Untrusted files:** workbook text only ever reaches SQL as escaped string literals; headers become SQL names only
  after being reduced to `[a-z0-9_]`; braces are escaped so dbt's Jinja cannot run workbook text. A file that unzips to
  more than 2 GB is refused. For extra XML hardening on untrusted files, `pip install defusedxml` (openpyxl then uses it
  automatically).
- **Output folder:** a rerun deletes the old `*.sql` files in `<out>/models/cellproof_cells/` so stale models cannot
  linger; point `--out` at a folder of its own.
- **`--dbt`** runs your installed `dbt` on the generated project (no shell).

See [SECURITY.md](SECURITY.md) to report a vulnerability.

## Development

```bash
python -m venv .venv && . .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -e ".[dev,dbt,compute]"
python -m pytest tests -q                          # ~8 min with dbt on PATH
CELLPROOF_PERF=1 python -m pytest tests/test_perf.py -s   # the 200k-formula benchmark
```

Rules the code follows: never guess a value (anything not translated exactly is UNRESOLVED with a reason a user
understands); a false PASS is the worst bug; no DuckDB exception may escape from a supported function; tests first; real
Excel is the ground truth. Design notes: [DESIGN.md](DESIGN.md) and `docs/superpowers/specs/`.

## License

[MIT](LICENSE)
