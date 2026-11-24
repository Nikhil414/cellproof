# Changelog

## Unreleased

- **Six dynamic-array functions: `TAKE`, `DROP`, `CHOOSEROWS`, `CHOOSECOLS`, `TOCOL`, `TOROW`.** Each either
  translates exactly or raises `Unsupported`; none guesses. Rows and columns count from the start of the array, a
  negative count from its end, and a fully emptied array is handled the same way `FILTER`'s empty result is.
  **The boundary error codes are not yet proven.** For `TAKE`/`DROP` with a count of zero, and for a `CHOOSEROWS`/
  `CHOOSECOLS` index that is zero or past the end, Excel documents `#CALC!` and `#VALUE!`. The tests agree with those
  codes, but the expected values come from Microsoft's documentation, not from an Excel-saved workbook, so the
  fixtures confirm the code does what the code says rather than that Excel does it. Microsoft's documentation was
  already wrong once about this exact family (`FILTER` returning nothing), so these need an Excel-saved file before
  they count as ground truth. A real workbook that disagrees would turn an expectation, not a translation, into a fix.
- **Fixed: `DROP` past the end of a dynamically-sized array answered a wrong number instead of raising.**
  `DROP` checked for "this drops everything" only against arrays whose height it knew while compiling. For a
  `FILTER` result, whose height is known only when the query runs, the check was skipped: dropping more rows than
  existed produced an empty relation, and `SUM` over it answered `0.0` and `COUNT` answered `0.0`, where Excel
  raises an error. A count is the worst result to invent - it looks like a value and the report marks the cell
  PASS. Such a result now becomes `UNCHECKED` and reads as UNRESOLVED, the rule `FILTER` already follows, because
  which code Excel raises here is the open question above. Dropping rows that leave rows behind is still answered
  exactly, and a spill of the emptied array no longer shows a bare `#N/A`.
- **Fixed: `cellproof diff` could miss a change inside a long arithmetic formula.** A formula too big to keep as one
  expression has its sub-results bound into numbered layers, so what the cell actually computes is split across them.
  The comparison read only the outer expression, so two different long formulas could compare as the same SQL: a
  change that exists but is never reported. The layers are now part of the compared text. The same fix covers the
  other direction: a cell that moves to a different position in its batch used to be renumbered and could report a
  fake "SQL changed", because the layer names carry that position. A compared SQL that now runs long is prefixed in
  `DIFF.md` (the whole of it stays in `diff.json`), so one change cannot make an unreadable table row.
- **`README` says what happens between Excel's two ceilings.** A result above 9.99e307 is already `#NUM!` in Excel
  while DuckDB can still hold it (up to 1.79e308). Only a non-finite result is turned into `#NUM!`, so such a cell
  FAILs honestly instead of passing. Nothing is guessed.

## 1.2.0 (2026-11-22)

- **`cellproof diff old.xlsx new.xlsx`: what changed between two versions, and what it broke.** Both workbooks are
  verified exactly as `convert` verifies them, so every verdict is one `convert` would give. It reports, with counts:
  cells added and removed (by A1 address), saved input values that changed with both values (including a change of
  kind, `2` becoming `"2"`), formula texts that changed with both texts, verdicts that changed per cell with both
  values and the reason (a new FAIL is the headline and comes first, then UNRESOLVED, then PASS; vanished cells and
  cells that newly appear and pass are left out), and regions that verified as a block in one version and not the
  other, with the compiler's own note. The case only a diff exists for: **a formula whose text did not change but
  whose SQL does** - a named range that grew, a table that gained rows, a range's corner that moved. That comparison
  comes from a second compile in cell mode, canonicalised to one text per cell (cell tags in place of batch names),
  so a block of rows regrouping into a region - or out of one - does not read as hundreds of statements changing.
  Cells are matched by sheet name and A1 address only: a renamed sheet reads as everything removed and everything
  added, the truth rather than a guess at which sheet became which. `diff.json` is always written; `DIFF.md` only
  when a verdict changed or ten changes in total, and its headline names the new FAILs and says a formula that reads
  a wrong answer is usually wrong too - it never claims to know which to fix first. Exit 1 only when a cell that
  passed before fails now: a difference by itself is not a failure. `--fast` works here too.
- **FAILs are split: the ones to look at, and the ones that only read them.** A primary FAIL reads no FAIL, directly
  or through other formula cells, so the difference starts at that cell; an inherited FAIL reads one and may be
  right itself, and is attributed to the nearest primary it reaches. `REPORT.md` replaces the flat failed-cell table
  with "Failures to look at" (primary, with how many failed cells read each) plus a line counting the inherited ones;
  `report.json` gains a `failures` key, and `--stats` prints the two counts. Cells whose expected value came from the
  `formulas` library have no saved value to disagree with, so they are excluded and counted separately. Nothing is
  ever labelled stale: both kinds stay FAILs.
- **The report quotes the file's own evidence that its saved values may be out of date.** `calcMode="manual"`, a
  pending `fullCalcOnLoad="1"`, a missing `calcId`, or a last-saving program that is not Microsoft Excel each put one
  sentence above the failure tables: recalculate (F9) and save in Excel, then run again. A clean Excel file gets no
  sentence, and a file with no FAILs gets none either.
- **`cellproof prove`: a result you can hand to someone.** It verifies the workbook as `convert` does, then writes a
  folder holding `cells.sql` (one runnable DuckDB script: the input cells, every formula batch in dependency order, a
  final `SELECT` for every cell), `report.json` and `REPORT.md` (what `convert` would have written), and `proof.json`
  (the workbook's size and hash, the cellproof/DuckDB/Python versions and the platform, the workbook's own `<calcPr>`
  attributes and last-saving program, the counts, a hash per file and one over the set).
- **`cellproof prove --check DIR [--workbook book.xlsx]`** re-checks a bundle without trusting the machine that made
  it: it re-hashes the files, optionally re-reads the workbook and compares the values it holds with the ones the
  report records as Excel's, and runs `cells.sql` on a fresh DuckDB to see whether it still answers what `report.json`
  says. Exit code 0 (all held), 1 (an answer moved), 2 (a file, a hash or the format is wrong), 3 (cellproof generated
  SQL DuckDB rejected). Console output is unchanged for every other command.
- **The seal is not a signature, and the report says so.** The hashes catch an accident or a casual edit; they cannot
  stop someone with write access to the folder, who can reseal. What that person cannot fake is `cells.sql`
  disagreeing with `report.json`, which is why `--check` runs the script rather than stopping at the hashes. One line
  under `REPORT.md`'s title states it.
- **One notion of "equal".** `verifier.matches` is now a thin wrapper over a new plain `verifier.same(kind, num, txt,
  other_kind, other)`; `prove --check` compares answers with it, so a bundle can never pass `--check` on a value the
  verifier would call a FAIL.
- **`Book.calc`** records the workbook's `<calcPr>` (`calcMode`, `fullCalcOnLoad`, `calcId`, all `None` when the file
  is silent) next to `Book.app`, through both loaders. It is only read, never acted on: 1.2 does not decide whether a
  saved value is stale.
- Runtime dependencies are still only `openpyxl` and `duckdb`; the bundle is DuckDB-only, and `proof.json` records
  which engine produced the verdicts so a Postgres proof can come later without changing the format.

## 1.1.1 (2026-11-18)

- **`--fast`: read the workbook in one pass.** `cellproof convert`, `cellproof sql` and `cellproof watch` accept
  `--fast`, which reads each worksheet's XML once instead of openpyxl's two full loads (formulas, then the values
  Excel saved) and keeps only the cells cellproof needs. It is the same `Book`: on the 85 public test workbooks
  (61,510 cells) and on a 200,000-formula workbook the fast load is **identical to the openpyxl load**, and
  `tests/test_fastread.py` is a differential test against `reader.load` (values, kinds, formulas, arrays, names,
  tables, merges, hyperlinks, max rows, epoch, app, and the errors a corrupt file raises).
- **Measured:** 200,000 formula cells: read **2.40 s -> 1.35 s** (1.78x), and a full `convert` **5.32 s -> 4.45 s**.
  The 85 public workbooks together: 0.85 s -> 0.51 s. Loading is now about a quarter of a `convert`, so the ceiling
  for this work is real: the rest is compiling and verifying.
- **Nothing new to install, and nothing to lose:** the loader uses the standard library's `xml.etree`, reuses
  openpyxl's own parsers for everything except the cell loop, and adds no dependency. When it meets something it does
  not handle it raises rather than guess, and the CLI prints one note to stderr and re-reads the file with openpyxl -
  so `--fast` can be slower, never different. Without the flag nothing changes.
- **Fixed: TEXT's number formats on the Postgres target.** sqlglot 30.33 parses DuckDB's `printf(...)` and
  `format(...)` into an expression of its own, so the old rewrite missed them and Postgres ran its *own* `FORMAT` -
  `TEXT(1234.5678,"0.00")` failed, and `TEXT(1234.5678,"#,##0.00")` would have returned the pattern itself. Both are
  now translated to a `pg_temp` helper with DuckDB's result (grouping, ties to even, `-0.00` keeps its sign); a
  pattern cellproof never writes is UNRESOLVED instead of a wrong value. Postgres needs **sqlglot 30.10 or newer**.

## 1.1.0 (2026-11-17)

- **`cellproof watch book.xlsx [--interval SECONDS] [--runs N]`:** verify the workbook, then verify again on every
  save and print only what changed - cells that broke or were fixed (with both values or the reason), FAILs whose
  SQL answer moved, FAIL<->UNRESOLVED flips, rows added to or removed from a formula column, and new unresolved
  reasons counted so a new unsupported function is visible at once. It watches the file's modification time and size
  (default check every 1.0 s; Ctrl+C to stop); the first run prints the counts and where to start like `convert`
  does. A workbook that cannot be read while you watch is not an error (Excel may still be writing it): the last
  result is kept and the next save is another chance. A workbook that cannot be read at all exits 2.
- **`cellproof sql book.xlsx [--sheet X] [--range D2:D1000] [--out FILE]`:** prints a runnable DuckDB script for a
  selection of cells instead of writing a whole dbt project. The workbook is compiled as usual, then everything the
  selection does not need is dropped: only the input cells the formulas read, the formula batches they depend on (in
  dependency order), and a final `SELECT` for the cells asked for. A batch is kept or dropped whole, because it is one
  statement; the kept batches are closed under "reads", so the script always answers exactly what `convert` computes
  for those cells. `--range D2` on `examples/functions.xlsx` turns 16 statements into none; `G20:G30` keeps 4 of 16.
  SQL goes to stdout and the reasons an asked-for cell is missing go to stderr (ranked, "Where to start"), so
  `cellproof sql book.xlsx > run.sql` always yields a file you can run; exit code 1 says a cell was left out.

## 1.0.0 (2026-11-17)

- **Postgres target:** `cellproof convert book.xlsx --target postgres [--pg-url URL] [--dbt]`. Every cell is verified
  on DuckDB, then again on Postgres with the SQL translated (sqlglot, tree rewrites, and `pg_temp` helper functions with
  DuckDB's semantics). On the 85 public test workbooks Postgres gives the same verdicts as DuckDB: 11,080 PASS, the
  same 16 FAIL, none new. A batch that does not translate or run makes its cells UNRESOLVED, and the book is compiled
  again so every cell reading them is too: the dbt project only holds SQL that ran on Postgres.
- **dbt-postgres project:** with `--target postgres` the generated project targets dbt-postgres. Its helpers are
  created in the target schema by an `on-run-start` macro, and `profiles.yml` reads `CELLPROOF_PG_*` environment
  variables. `--dbt` runs `dbt build` on the Postgres that verified the workbook. New extras: `postgres`,
  `dbt-postgres`.
- **Browser playground** (https://nikhil414.github.io/cellproof/): drop a workbook, get every verdict and the dbt
  project, in the page (Pyodide, DuckDB 1.2); nothing is uploaded. CONCAT and TEXTJOIN no longer need DuckDB 1.3.
- **The console says what to fix first.** After the counts line, every conversion prints the top five unresolved
  reasons ranked by how many cells each blocks ("Where to start"), the same ranking `REPORT.md` has always carried -
  so a long list of unfinished cells no longer needs the report opened to be useful. `--stats` still adds the full
  tally.
- **Wrong answers other cells read are called out** (`REPORT.md`, `report.json`, console): a FAIL whose value feeds
  other formulas can carry its difference further than a FAIL nobody reads. Only direct readers are listed ("read
  by", not "blocks"): a reader may add a difference of its own, so the list is a neighbourhood, not a guarantee.
  The reverse pass reuses the compile's tokenizing and costs one bisect per column, so whole-column references stay
  cheap.
- Add DEC2BIN, DEC2HEX, DEC2OCT and their reverse conversions; unverified edge cases remain UNRESOLVED.
- CI: a Postgres 16 service runs the Postgres tests and dbt build on every push.

## 0.6.1 (2026-10-31)

- **ROUND rounds the value Excel shows** (15 significant digits, half away from zero): `ROUND(1.005, 2)` is 1.01,
  as in Excel. The double 1.005 is 1.00499999..., and the binary rounding before gave 1.0 - an honest FAIL, never a
  false PASS. Same method as ROUNDUP/ROUNDDOWN; Microsoft's ROUND examples are tests. (Released from release-0.6.)

## 0.6.0 (2026-10-31)

- **A batch DuckDB rejects no longer stops the run:** its cells are UNRESOLVED with DuckDB's message, the book is
  compiled again so every cell reading them is UNRESOLVED too (none can read a missing value), and the report is
  written. The exit code is still 3: it is a bug in cellproof, please report it.
- **ROW, COLUMN** (of the formula's cell, of a cell, and in an array formula of a range: one column of row numbers,
  one row of column numbers - real Excel), **CODE** (Windows-1252, CHAR's inverse; a character outside it, or above
  127 in a file Excel for Mac calculated, is UNRESOLVED), **UNICODE, UNICHAR, N** (of a range in an array formula:
  its first cell), **ISODD, ISEVEN** (no ranges in an array formula: #VALUE!, like EDATE).
- **IRR and RATE** by Newton's method in SQL (a recursive CTE). Not guessed: cash flows with several sign changes
  (several roots), an error among IRR's cells, a RATE whose answer is 0, nper not a positive whole number.

## 0.5.0 (2026-10-31)

- **FILTER, SORT, UNIQUE inside other calculations:** `SUM(FILTER(...))`, `COUNTA(UNIQUE(...))`, AVERAGE, MIN, MAX,
  PRODUCT, STDEV, VAR of them; **ROWS** and **COLUMNS** (of a range or of these). An empty element reaching an
  aggregate is UNRESOLVED (whether Excel skips it or takes 0 was not checked), even through COUNT.
- **XMATCH** (exact, next smaller/larger, wildcards; first-to-last or last-to-first), **SORTBY** (several keys; by
  rows, or by columns when the keys are one row: real Excel), **SEQUENCE** (more than a million elements is
  UNRESOLVED and builds nothing). FILTER's if_empty can be a range, which spills whole.
- **Spills are one statement:** a spilled block is one batch computing its relation once (verify 1,334 cells
  31.0 s -> 0.03 s; 33,334 cells in 0.7 s), and one dbt model.
- **Not guessed:** XLOOKUP/XMATCH next smaller/larger across types or over text with punctuation (Excel ranks numbers
  < text < TRUE/FALSE and orders text by its own collation), and exact text matches with a German sharp s (Excel
  matched "ss" to it). Some of these used to PASS by luck; they are UNRESOLVED now.

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
