# cellproof: the whole journey, from idea to today

This file tells the story of the project from the first design (2026-10-11) to now (2026-10-20): what was built, in
what order, why, what broke, and what was learned. The technical details of each step live in `DESIGN.md`,
`docs/superpowers/specs/`, `docs/superpowers/plans/`, `PROGRESS.md` and `CHANGELOG.md`; this is the readable summary.

---

## 1. The idea (2026-10-11)

**Problem.** A lot of business logic lives in Excel workbooks that one person understands. Moving that logic into a
database is slow and risky, because nobody can prove that the new numbers equal the old ones. Existing tools copy the
*data* out of Excel; none translated the *formulas*.

**Idea.** A command-line tool that:

1. reads every formula of an `.xlsx`,
2. translates each one into SQL,
3. runs that SQL on DuckDB (a fast local database),
4. compares every result with the value Excel itself saved in the file.

Excel's saved values are a free answer key. Every formula cell gets one verdict:

| Verdict | Meaning |
|---|---|
| PASS | the SQL gives the value Excel saved |
| FAIL | the SQL gives a different value; the report shows both |
| UNRESOLVED | the tool could not translate it exactly, and says why. It never guesses. |

The output is a **dbt project** (the standard way data teams ship SQL), so the result is something a team can use.

**Rules decided on day one** (they still hold):

- Never guess a value. Anything not translated exactly is UNRESOLVED with a reason a user understands.
- A false PASS is the worst possible bug. When unsure, give an honest FAIL or UNRESOLVED.
- Real Excel is the ground truth, not our idea of what Excel does.
- Cell values never leave the machine. No network calls with workbook data.
- Only two runtime dependencies: `openpyxl` (reads the file) and `duckdb` (runs the SQL).
- Tests first, and measure before claiming speed.

The working name was **unexcel**, already known to clash with an R package.

---

## 2. Timeline at a glance

| Date | Step | Result |
|---|---|---|
| 2026-10-11 | Design + Plan 1: the core | reader, parser, references, dependency order, SQL compiler, verifier, dbt output, report, CLI |
| 2026-10-11 | Plan 2a: scale and Excel errors | batches (no crash at ~1k formulas), Excel error codes, safety limits |
| 2026-10-12 | Plan 2b: function library | 93 Excel functions, run-time types, Excel's conversions |
| 2026-10-14 | First check against real Excel | 3 workbooks saved by real Excel: every cell PASS (4 UNRESOLVED on purpose) |
| 2026-10-14 | Plan 2c: speed in cell mode | faster dependency order, range index, smarter batches |
| 2026-10-16 | Plan 2d: table mode | a column of one pasted-down formula becomes one readable SQL statement |
| 2026-10-18 | Slice 2 | regions split instead of dissolving; running totals as window functions |
| 2026-10-18 | Open-source release 0.1.0 | public GitHub repo (fresh history), MIT license, CI, PyPI as `unexcel` |
| 2026-10-19 | Lookups and criteria row-wise | VLOOKUP/INDEX/MATCH/XLOOKUP and COUNTIFS/SUMIFS in table mode |
| 2026-10-19 | Faster compile | a pasted-down formula is tokenized once |
| 2026-10-19 | Renamed to **cellproof**, 0.2.0 | new name, new PyPI package, repo renamed |
| 2026-10-19 | 85 public test workbooks | FAIL 229 -> 10 (all 10 in LibreOffice-saved files), 0 crashes |
| 2026-10-19 | First real user workbook | 3 PASS / 47 UNRESOLVED -> 48 PASS / 0 FAIL / 2 UNRESOLVED |

---

## 3. Step by step

### 3.1 Plan 1: the core (2026-10-11)

Built test-first, one module at a time:

- **Reader**: loads the workbook twice with openpyxl, once for formulas and once for the values Excel saved.
- **Parser**: formula text to a tree, with Excel's operator precedence (`-2^2` is 4 in Excel).
- **References**: `A1`, `Sheet!B2:C9`, whole columns, defined names.
- **Dependency order**: every cell after the cells it reads; circular references detected.
- **Compiler (cell mode)**: each formula becomes a SQL expression.
- **Verifier**: runs everything on DuckDB and compares with Excel's saved values.
- **Emitter**: writes the dbt project; `--dbt` also runs `dbt build`.
- **Report**: `REPORT.md` and `report.json`.
- **CLI**: `unexcel convert workbook.xlsx`.

The first end-to-end test was a small commission workbook.

### 3.2 Plan 2a: scale and Excel errors (2026-10-11)

- The first version crashed around 1,000 formulas (one giant SQL statement). Fix: run formulas in **batches** over one
  shared cells table.
- Excel **error codes** (`#DIV/0!`, `#N/A`, `#VALUE!`...) now flow through operators and functions, and must match
  exactly.
- Safety: a deeply nested formula or one whose SQL would grow exponentially becomes UNRESOLVED instead of crashing or
  eating memory; a styled cell on the last row no longer costs seconds or gigabytes.
- Optional fallback answer key (the `formulas` library) for files with no saved values, always labelled as such.
- A 200,000-formula benchmark was added: 390 s at that point.

### 3.3 Plan 2b: the function library (2026-10-12)

- **93 Excel functions** (SUM, IF, VLOOKUP, XLOOKUP, INDEX/MATCH, SUMIFS, TEXT, dates, ...).
- Run-time types: a cell's type can be known only when the SQL runs ("any").
- One criteria engine for the whole COUNTIF/SUMIF/AVERAGEIF(S) family (`">"&B1`, wildcards, `"<>"`).

### 3.4 First check against real Excel (2026-10-14)

Three workbooks were saved by real Microsoft Excel and checked cell by cell. Excel taught us things no document says,
each now a test (`tests/test_excel_findings.py`):

- COUNTIF `"<>5"` skips only the number 5 (the text "5" is counted).
- DAYS360, `TEXT(-0.001,"0.00")`, `MOD(1,0.1)`, dates inside 1900, `(-8)^(1/3)`, error order in `SUM(H3,1/0)`.
- Approximate MATCH/VLOOKUP on unsorted data depends on Excel's binary search, so it is UNRESOLVED.

Result: every cell of the three workbooks PASS, 4 UNRESOLVED on purpose.

### 3.5 Plan 2c: speed in cell mode (2026-10-14)

Measured first, then four changes: a Fenwick tree for dependency order without expanding ranges, a range index,
batches closed by SQL size instead of cell count, smaller dbt models.

### 3.6 Plan 2d: table mode (2026-10-16)

Most real sheets repeat one formula down a column (`=B2*C2`, `=B3*C3`, ...). Such a column is now a **region**:
compiled once into one readable SQL statement (a "wide model" with real column names), instead of one expression per
cell. Hard parts solved along the way: safe SQL names from headers, regions that would create false cycles, proving
that table mode gives exactly the same answers as cell mode (a test rule since then).

### 3.7 Slice 2 (2026-10-18)

- A region with another cell between its columns is split instead of thrown away.
- Running totals (`=C1+B2`, `=SUM($B$2:B2)`) become SQL window functions, but only when Excel's saved values prove
  the different addition order cannot change the answer. A ledger whose balance cancels back to zero stays in
  Excel's exact order.
- 200k-formula benchmark: **390 s -> 31 s**, all PASS.

### 3.8 Open-source release 0.1.0 (2026-10-18)

- MIT license, README with an honest limits section, SECURITY.md, CI on Windows/macOS/Linux x Python 3.10-3.13.
- Security review: zip bombs refused, crafted numbers beyond Excel's range UNRESOLVED, workbook text only reaches SQL
  as escaped literals.
- **Privacy**: the public repository was created with a fresh one-commit history and a no-reply author address.
- Published to PyPI as `unexcel` 0.1.0 with trusted publishing (no token stored anywhere).

### 3.9 Lookups and criteria in table mode (2026-10-19)

- **Problem**: a VLOOKUP pasted down a column was still one query per cell. DuckDB could not plan the SQL form the
  lookups used inside a region.
- **Fix**: every lookup became an aggregate (`arg_min` by position) over a table that does not depend on the row,
  which DuckDB turns into one join for the whole column.
  - 1,000 VLOOKUPs: verify **5.7 s -> 0.1 s**. 50,000 VLOOKUPs: all PASS.
- Same idea for COUNTIFS/SUMIFS with a criterion from the row:
  - 1,000 COUNTIFS: verify **29 s -> 0.1 s**.

### 3.10 Faster compile (2026-10-19)

Profiling showed the formula tokenizer ran twice per cell and was most of the compile time. Now the first cell of a
pasted-down column is tokenized and the next rows are predicted by shifting row numbers; a cell is tokenized again only
when the prediction does not match its real text.

- 50k VLOOKUPs: compile **6.1 s -> 1.8 s**.
- 200k benchmark: **31 s -> 26 s**.

### 3.11 Renamed to cellproof, 0.2.0 (2026-10-19)

- `unexcel` clashed with an R package. Candidates were checked on PyPI, CRAN, npm and GitHub.
- **cellproof** was free everywhere and says what is unique: every cell is proven against Excel.
- Package, command, import, dbt project and GitHub repository were renamed (old links redirect).
- 0.2.0 was published to PyPI as `cellproof`. The old `unexcel` 0.1.0 stays on PyPI under the old name.

### 3.12 Testing on 85 public workbooks (2026-10-19)

The tool was run on every test workbook of three other Excel formula engines (pycel, formulas, xlcalculator), most
of them saved by real Excel. First run: 1 crash, 229 FAIL. Each FAIL was traced to its cause:

- **Bugs in cellproof, fixed**:
  - a name defined as a formula crashed the reader;
  - Excel's 29 February 1900 and date-formatted negative numbers were read wrong;
  - escaped control characters in text were not decoded;
  - AND/OR must skip text and empty cells;
  - an empty result of IF/IFERROR is 0, not "";
  - EDATE/EOMONTH/MROUND/VALUE refuse TRUE/FALSE;
  - error order in SUMPRODUCT;
  - PROPER of another cell built SQL DuckDB rejected.
- **Not bugs, now handled honestly**:
  - 8 files were saved by LibreOffice, not Excel, so their saved values follow LibreOffice; the report now warns.
  - One file was calculated by Excel for Mac, where CHAR above 127 gives other characters; now UNRESOLVED.
  - Text like `"1/2/2020"` or `"1,000"` is read by Excel using the computer's language settings, which the file does
    not record; now UNRESOLVED instead of a wrong answer.
- One of my own earlier assumptions (`AND(TRUE,"0")` is `#VALUE!`) was wrong: real Excel says TRUE. The test was
  corrected to match Excel, never the other way round.

Result: **FAIL 229 -> 10** (all 10 in the LibreOffice files), **0 crashes**. `python tests/corpus.py DIR` repeats it.

### 3.13 The first real user workbook (2026-10-19)

A practice workbook full of array-style formulas: **3 PASS, 0 FAIL, 47 UNRESOLVED**. Added:

- **Implicit intersection**: `=IF(E$3:E$41="F",B$3:B$41)` uses the formula's own row, as Excel does (40 cells).
- **`INDIRECT("English")`** of fixed text is just the named range.
- **MODE**.

The first version of implicit intersection gave 89 for `=MAX(IF(E3:E41="F",math,0))` where Excel saved 99: the
verifier caught it as a FAIL. Investigation showed Excel passes the whole `math` range through IF into MAX, so that
case is now honestly UNRESOLVED. The same investigation found an older bug: a bare word like `=i` with no matching
name was read as the whole column I.

Result: **48 PASS, 0 FAIL, 2 UNRESOLVED** (one real array formula, one range passed through IF).

---

## 4. How it works now (short version)

```
workbook.xlsx
   |  reader: formulas + Excel's saved values (and which program saved the file)
   v
parse every formula -> find columns of one pasted-down formula (regions)
   |
   v
dependency order (regions as one block) -> SQL: one statement per region, small batches for the rest
   |
   v
run on DuckDB -> compare every cell with Excel's saved value -> PASS / FAIL / UNRESOLVED
   |
   v
REPORT.md + report.json + a dbt project (wide readable models per region)
```

---

## 5. Where things stand (2026-10-20)

| | |
|---|---|
| Public repository | github.com/Nikhil414/cellproof (MIT, CI green) |
| PyPI | `pip install cellproof` (0.2.1) |
| Tests | 425 passing (dbt build included) |
| Functions | 94 Excel functions |
| Real Excel checks | 3 own workbooks all PASS; 85 public workbooks: only FAILs in LibreOffice files; 1 user workbook 48/0/2 |

**Speed** (Windows 11 laptop):

| Workbook | Time |
|---|---|
| 200k formulas (4 columns x 50k rows, a running total) | 26 s, all PASS (was 390 s) |
| 50k VLOOKUPs into a 1,000-row table | 4.6 s, all PASS |
| 1,000 COUNTIFS with a criterion from the row | verify 0.1 s (was 29 s) |

---

## 6. What we learned

1. **Real Excel is the only reliable reference.** Several "obvious" Excel behaviours turned out wrong once checked
   against files Excel saved.
2. **The answer key catches our own mistakes.** Every wrong translation so far showed up as a FAIL, including two of
   my own wrong fixes; none slipped through as a false PASS.
3. **Real files find what tests do not.** One user workbook needed three new features; 85 public workbooks found
   eight bugs.
4. **Not every FAIL is our bug.** A file can be saved by LibreOffice, calculated on a Mac, or depend on language
   settings. The tool now says so instead of blaming itself or guessing.
5. **Measure, then optimize.** Every speed-up came from a profile or a benchmark, not from a hunch.

---

## 7. What is next

1. Release 0.2.1 with the fixes since 0.2.0.
2. Launch prep: a short demo and a sample report on the README front page; then posts on LinkedIn and Twitter/X.
3. Array formulas and `{1,2}` array constants (the biggest remaining gap for real workbooks), LOOKUP, YEARFRAC.
4. More real workbooks from users.
5. Plan 3: an AI retry for cells that fail, which sees only the formula structure, never the values.
