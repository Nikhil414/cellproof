# cellproof: progress

Last updated: 2026-11-20. This file is the single place that says where the project stands. Read it cold: a person
or a coding agent should be able to pick up the work from here. For the rules of the code read `AGENTS.md`; for what
the tool does and its limits, `README.md`.

## What cellproof is (one paragraph)

`cellproof convert book.xlsx` translates every formula of an Excel workbook into SQL, runs it (DuckDB; Postgres too
with `--target postgres`), and compares each result with the value Excel saved in the file. Every formula cell gets
PASS, FAIL (both values shown) or UNRESOLVED (with the reason). It never guesses: a false PASS is the worst possible
bug. Output: `REPORT.md` / `report.json` and a dbt project that runs exactly the SQL that was checked. Earlier name:
`unexcel` (until 0.2.0; old plans still use it).

## Status at a glance

| What | State |
|---|---|
| PyPI | **1.1.0** is the latest release (`pip install cellproof`), published 2026-11-17 from tag `v1.1.0` (trusted publishing); **1.1.1** (`--fast`) is on `main`, awaiting its release tag |
| `main` | **1.1.1**: `--fast` (read each worksheet once, identical `Book`), on top of 1.1.0's `cellproof sql` (PR #19) and `cellproof watch` (PR #20); PR #22 also fixes TEXT's number formats on the Postgres target (sqlglot 30.33 turned `printf`/`format` into an expression the rewrite missed). Tag `v1.1.0` = the 1.1.0 release commit |
| Release | [v1.1.0](https://github.com/Nikhil414/cellproof/releases/tag/v1.1.0), 2026-11-17: 13/13 CI checks green, `publish.yml` publishes to PyPI. 1.0.0 ([v1.0.0](https://github.com/Nikhil414/cellproof/releases/tag/v1.0.0), 2026-11-17) added the Postgres target, dbt-postgres, the playground and the console triage |
| Playground | Live at https://nikhil414.github.io/cellproof/ (GitHub Pages, deployed by `.github/workflows/pages.yml` on every push to `main` that touches `src/`, `playground/` or the example) |
| Tests | `main` (1.1.1: `--fast`): **741 pass, 3 skip** (744 collected) with dbt-postgres installed and the 85-workbook corpus present (2026-11-20; the 3 skips are the optional `compute` extra and the opt-in 200k benchmark), including 85 parametrized differential tests that read every public workbook with both loaders and require the same `Book`, and the Postgres format tests. 1.1.0: 622 pass, 3 skip |
| CI | `tests` workflow: Windows/macOS/Linux x Python 3.10-3.13, plus the `postgres` job (Postgres 16 service, Postgres tests and `dbt build`): **13 checks** on every pull request. Green on the release commit |
| Public test workbooks (85) | **11,080 PASS / 16 FAIL / 37,814 UNRESOLVED**, identical on DuckDB 2.0, on DuckDB 1.2 (the playground) and on Postgres (with `dbt build` passing for all 85). The 16 FAIL are explained below. A fresh re-download of the corpus on 2026-11-15 (the source repos changed; 66 files now) measured **13,043 PASS / 34 FAIL / 19,183 UNRESOLVED** — one workbook (formulas' `test.xlsx`, a function catalog) holds 17,454 of the UNRESOLVED, and `CONVERT` alone 3,593 (19%). Of the 34 FAIL, 19 were wrong error codes from PR #7's base conversions ([issue #10](https://github.com/Nikhil414/cellproof/issues/10)), 9 are the known LibreOffice drift in `INFORMATION.xlsx`/`SQRT.xlsx`; 6 are the known npv float cells; the 19 error codes were fixed the same day (PR #12), bringing the fresh-corpus count to **13,062 PASS / 15 FAIL**. 2026-11-15: the six base cross-pairs (HEX2BIN/HEX2OCT/BIN2HEX/BIN2OCT/OCT2HEX/OCT2BIN, semantics verified against Excel-saved values in `test.xlsx`) turned **669** more cells PASS: fresh corpus now **13,732 PASS / 15 FAIL / 18,513 UNRESOLVED**, 0 FAILs on `BASES.xlsx` and `test.xlsx` |
| Contributors | Open: PR #5 (overflow, changes requested). Merged: PR #7 (DEC2BIN/HEX/OCT and back, MilindLate, 2026-11-13, closes #2; unverified cases guarded, never guessed), PR #9 (`--stats` flag: count-sorted tally of UNRESOLVED reasons on the console, 2026-11-15), PR #12 (base conversions: match real Excel's error codes, closes #10, 2026-11-15), PR #15 (the six base cross-pairs, 2026-11-16, issue #14's first roadmap item), PR #17 (console triage: ranked unresolved reasons and the wrong answers other cells read, 2026-11-16), PR #8 (1.0: the Postgres target and dbt-postgres project, 2026-11-17, released as v1.0.0). Closed: PR #6 (superseded by #5). Issues #1, #3, #4, #14 open, #4 pinned; #3 has a contributor spec comment (2026-11-15: TAKE, DROP, CHOOSECOLS, CHOOSEROWS, TOCOL, TOROW against the `fn/dynamic.py` relation model); #10 was fixed in PR #12 and closed 2026-11-15; #14 is the 2.0 roadmap (ranked UNRESOLVED table) |
| Launch | Public launch (Show HN, r/dataengineering, LinkedIn, Shorts/Reels) planned for Tuesday 2026-11-10. 1.0 is the second announcement |

## Where things live

| Path | What |
|---|---|
| `src/cellproof/` | the package; module map in `AGENTS.md` |
| `src/cellproof/postgres.py` | Postgres target: `translate()` (DuckDB SQL -> Postgres via sqlglot + tree rewrites), `PRELUDE` (helper functions with DuckDB semantics, in `pg_temp`), `verify()`, `verify_book()` (1.0 branch), `dbt_env()` |
| `src/cellproof/emitter.py` | dbt project; on the 1.0 branch `write_project(target="postgres")`, `to_postgres()`, `helpers_macro()` |
| `playground/` | browser playground: `index.html` (page), `worker.js` (Pyodide in a Web Worker), `cellproof_web.py` (bridge: same conversion as the CLI), `build.sh` (builds the site) |
| `tests/` | pytest suite; `tests/corpus.py DIR [convert options]` runs the 85 public workbooks; `tests/xlsxkit.py` builds test workbooks |
| `src/cellproof/watch.py` | `cellproof watch` (1.1): `Run` (one verification's counts, verdict snapshot, seconds, "where to start"), `changes()` (what changed between two runs, worst first), `run_watch()` (the poll loop; fully injectable `out`/`sleep`/`stat` so tests are deterministic) |
| `src/cellproof/query.py` | `cellproof sql` (1.1): a runnable DuckDB script for a selection instead of a whole project |
| `examples/` | Excel-saved workbooks (`functions.xlsx` is the playground's example) and the scripts that made them |
| `docs/superpowers/specs/`, `docs/superpowers/plans/` | design and plan of each piece; `PLAN*.md` at the root are the early plans, kept for history |
| `docs/demo.gif` | README demo, cut from the promo video |
| `CHANGELOG.md` | what each release changed; 1.0.0 is written as "unreleased" on the 1.0 branch |

Local-only files on the maintainer's machine (kept out of git with `.git/info/exclude`): `LAUNCH.md` (launch posts and
launch-day answers), `PR_REVIEWS.md` (review drafts), `last_histoty.md` (an older session log).

## How to work on it

```bash
python -m venv .venv && . .venv/bin/activate
pip install -e ".[dev,dbt,compute]"           # + ",postgres" for the Postgres tests, + ",dbt-postgres" on the 1.0 branch
                                              # the postgres extra needs sqlglot 30.10+: older ones translate some SQL wrongly
python -m pytest tests -q                       # dbt must be on PATH (activate the venv) or the dbt tests skip
python tests/corpus.py /some/dir                # the 85 public workbooks (downloads them once); add --target postgres
CELLPROOF_PERF=1 python -m pytest tests/test_perf.py -s   # the 200k-formula benchmark
CELLPROOF_TEST_PG_URL=postgresql://... python -m pytest tests/test_postgres.py   # against a given Postgres (CI does)
playground/build.sh _site && python -m http.server -d _site   # the playground locally (Pyodide cannot load file://)
```

Release: bump `version` in `pyproject.toml`, move the CHANGELOG entry from "unreleased" to the date, tag `vX.Y.Z`
and publish a GitHub release (`gh release create vX.Y.Z`). `.github/workflows/publish.yml` uploads it to PyPI with
trusted publishing (no token stored).

Contributor PRs: GitHub holds the first CI run of a new contributor until the maintainer approves it ("Approve and run
workflows", or `gh api -X POST repos/Nikhil414/cellproof/actions/runs/<id>/approve`). Review from the diff and let CI
run their code; don't run an outside contributor's code on a personal machine.

## What is built (history)

Everything was built test-first; each piece was reviewed by a fresh reviewer.

| Piece | What it does |
|---|---|
| Plan 1 | reader, parser, references, dependency graph, cell-mode compiler, verifier, dbt emitter, report, CLI |
| Plan 2a | batched execution (fixes the crash at ~1k formulas), Excel error codes, recursion and size guards, optional `formulas` answer key |
| Plan 2b | 93 Excel functions, run-time types, Excel's conversions, one criteria engine for the *IF(S) family |
| Plan 2c | speed in cell mode: Fenwick-tree dependency order, range index, weight-based batches, small dbt models |
| Plan 2d | table mode: regions (columns repeating one formula) compile to one readable statement |
| Slice 2 | a region with a cell between its columns splits instead of dissolving; running totals (`=C1+B2`, `=SUM($B$2:B2)`) become window functions |
| Lookups row-wise | VLOOKUP, HLOOKUP, MATCH, INDEX, XLOOKUP pasted down compile into the region's statement (one join each instead of one query per cell) |
| Tokenize once | a formula pasted down a column is tokenized once; the next rows' text is predicted and only re-tokenized when it differs |
| Implicit intersection | a one-column range where one value is expected takes the formula's row; `INDIRECT` of fixed text; MODE |
| Criteria row-wise | the *IF(S) family with a criterion from the row (`=COUNTIFS($E:$E,">"&A2,...)`) compiles into the region's statement |
| Impact-ranked unresolveds | the report groups UNRESOLVED cells by the reason to fix first ("Where to start"), ranked by how many cells each blocks; report.json has `unresolved_groups` and per-cell `blocked_by` |
| Approximate lookups unsorted | MATCH type 1 / VLOOKUP / HLOOKUP / LOOKUP over a plain numeric input table out of order compile Excel's binary-search path; a changed table is #STALE -> UNRESOLVED, never a stale answer. Type -1 and text stay UNRESOLVED |
| Arrays in aggregates | `SUM(A1:A10*B1:B10)`, `SUMPRODUCT((A1:A9>5)*C1:C9)`, `MAX(IF(A1:A9="x",B1:B9))`, element by element, modern Excel's rules |
| 0.3.0 | PMT, PV, FV, NPER, NPV; LARGE, SMALL, RANK, STDEV/VAR, PERCENTILE, QUARTILE; NETWORKDAYS, WORKDAY; Excel table structured references; LET |
| 0.3.1 | security: WORKDAY's day span sized by the holidays present, not the range's area (DoS) |
| 0.4.0 | array formulas over a block and dynamic spills, cell by cell; `A1#` and `@`; FILTER, SORT, UNIQUE; Excel 365 error codes; security: array blocks cannot overwrite other formulas |
| 0.5.0 | FILTER/SORT/UNIQUE inside aggregates; ROWS, COLUMNS, XMATCH, SORTBY, SEQUENCE; each spill computed in one statement (1,334 cells: 31 s -> 0.03 s) |
| 0.6.0 | a batch DuckDB rejects no longer stops the run (recompiled, its cells UNRESOLVED); ROW, COLUMN, CODE, UNICODE, UNICHAR, N, ISODD, ISEVEN, IRR, RATE |
| 0.6.1 | ROUND rounds the 15-digit value Excel shows (`ROUND(1.005,2)` = 1.01) |
| Postgres verification (`main`) | `--target postgres [--pg-url]`: every verified DuckDB batch is translated and every cell verified again on Postgres (bundled `pgserver` or your own). Untranslatable batches and their readers are UNRESOLVED ("not yet on Postgres: ..."). Nothing stays in the database: temp table, helpers in `pg_temp` |
| Browser playground (`main`) | drop a workbook, get every verdict, "Where to start", the cell table, REPORT.md and the dbt project as a zip; Pyodide 0.27 + DuckDB 1.2 in a Web Worker; nothing uploaded |
| Console triage (`main`) | every conversion prints the top five unresolved reasons (ranked by cells blocked) and the wrong answers other formulas read (direct readers only, "read by" never "blocks"), so the counts alone no longer need the report; the reverse pass reuses the compile's tokens and bisects per column, so whole-column references are O(log n); `postgres._readers` now delegates to it |
| 1.0 (released 2026-11-17) | `postgres.verify_book` recompiles without what Postgres cannot run, so the dbt-postgres project only holds SQL that ran; helpers created in the target schema by an `on-run-start` macro; `profiles.yml` reads `CELLPROOF_PG_*`; `--dbt` builds on the Postgres that verified; CI Postgres 16 job |
| `cellproof sql` (1.1) | a runnable DuckDB script for a selection instead of a dbt project: the input cells the formulas read, the batches they depend on in dependency order (kept or dropped whole - one batch is one statement - and closed under "reads"), then a `SELECT` for the asked-for cells. SQL to stdout, ranked reasons to stderr, exit 1 when a cell was left out. Bypasses the verifier, so it costs one compile and no DuckDB run: 0.7 s for a 20k-cell workbook, 0.4 s to run the printed script |
| PR #7 (merged 2026-11-13) | DEC2BIN, DEC2HEX, DEC2OCT and BIN2DEC, HEX2DEC, OCT2DEC (issue #2); hex text cast to BIGINT so DuckDB 1.2 works; cases not checked in real Excel (`HEX2DEC("")`, a negative invalid `places`) stay UNRESOLVED, never guessed |

## The public test workbooks

`python tests/corpus.py DIR` runs the 85 .xlsx/.xlsm test files of pycel, formulas and xlcalculator (Excel-saved,
except 8 LibreOffice ones). Now: **11,080 PASS / 16 FAIL / 37,814 UNRESOLVED**, the same on DuckDB 2.0, DuckDB 1.2
and Postgres. History: 4,693 PASS (before LOOKUP, YEARFRAC, array constants) -> 6,874 (2026-10-19) -> 7,528 (0.3.0)
-> 11,080 (0.6.1).

The 16 FAIL are not wrong translations: 10 are cells in LibreOffice-saved files (LibreOffice's answer differs from
Excel's), 6 are in pycel's `npv.xlsx` where `=C7-A7` subtracts two numbers equal to 1e-13 relative, so the last digits
differ (every NPV cell itself PASSes). The UNRESOLVED are mostly functions those files test on purpose (CONVERT 3,593
cells in one file alone, DEC2HEX/BIN/OCT, CHAR above 127 in a Mac-calculated file, ...). Findings are tests in
`tests/test_corpus_findings.py`.

A user's practice workbook went 3 PASS / 47 UNRESOLVED -> 48 PASS / 0 FAIL / 2 UNRESOLVED; it found 4 gaps the 85
public files did not.

## Measured speed (Windows 11 laptop)

| Workbook | Before table mode | Now |
|---|---|---|
| 200k formulas (4 columns x 50k rows, one running total) | 390 s | 26.0 s (read 8.3, compile 11.3), all PASS |
| Flat `=A*1.1`, 20k cells | verify 28.3 s | verify 0.2 s |
| `--dbt`, flat 20k | 225.9 s, 213 MB | 9.6 s, 1.6 MB |
| VLOOKUP per row, 50k rows | - | 4.6 s (read 2.2, compile 1.8, verify 0.6), all PASS |
| COUNTIFS per row (`">"&A2`, 2 criteria), 1k rows | verify 29.0 s, 500 batches | verify 0.1 s, 1 batch |
| SUMIFS / COUNTIFS per row, 50k rows | - | verify 0.6-0.8 s; compile 8.7-9.5 s |

The read path (measured 2026-11-18 on a 2024 Apple laptop, Python 3.12, openpyxl 3.1.5, medians of 3):

| Workload | openpyxl (two-pass) | `--fast` (one pass) | |
|---|---|---|---|
| 200,000 formula cells, 14.6 MB | 2.40 s | 1.35 s | 1.78x |
| the same workbook, `convert` end to end | 5.32 s | 4.45 s | 1.20x |
| the 85 public workbooks, 61,510 cells | 0.85 s | 0.51 s | 1.68x |

Loading is now about a quarter of a conversion, so the read path is no longer where the time goes; compiling and
verifying are. The loader is the standard library's `xml.etree` plus openpyxl's own parsers for everything except the
cell loop (no new dependency), and is required to produce a `Book` identical to `reader.load`'s, which 85 corpus
differential tests check.

Playground (Chromium): the example converts in 3.3 s; a 23,313-cell workbook in about 60 s with the CLI's exact
verdicts. Startup is 2.5-3.5 s on a good connection (Pyodide from jsDelivr, DuckDB from jsDelivr checked by SHA-256,
the site's own copy as fallback); on a network that crawls on github.io it took minutes.

## What real Excel taught us (each is a test)

From Excel-saved files (`tests/test_excel_findings.py`, `tests/test_corpus_findings.py`):

- COUNTIF `"<>5"` skips only the number 5 (the text "5" is counted); `"<>"` counts a formula that returns "".
- DAYS360 (US): only a 31st is moved; the last day of February ends the period unchanged unless the start is also the
  last day of February.
- `TEXT(-0.001,"0.00")` is `"0.00"`. `MOD(1,0.1)` is `0.09999999999999995`. Inside 1900 a date is linear:
  `DATE(1900,2,29)`, `DATE(1900,1,60)` and `DATE(1900,3,0)` are all 60.
- Number -> text is plain digits from 1E-11 to 1E+19 and scientific from 1E+20; ROUND rounds the 15-digit value shown.
- Errors typed into a formula come before errors inside referenced cells: `SUM(H3,1/0)` is #DIV/0!.
- `(-8)^(1/3)` is the real cube root (-1.9999999999999998); other fractional powers of a negative number are #NUM!.
- Past an array's size a block cell is #N/A (even under IFERROR), but IS functions see per-element #N/A padding.
- EDATE, EOMONTH, MROUND, YEARFRAC, ISODD, ISEVEN do not lift over ranges in array formulas (#VALUE!).
- ROW of a range is one column of row numbers; COLUMN one row; N of a range is its first cell.
- PERCENTILE/LARGE: k's error comes before the range's; QUARTILE checks quart before truncating.
- SORTBY with one-row keys sorts columns; FILTER's if_empty range spills whole.
- XMATCH next-smaller/larger compares across types with Excel's own collation; ß matches "ss" (both UNRESOLVED).
- Empty FILTER without if_empty: #CALC! in the docs but #VALUE! in a saved file (UNRESOLVED).
- `XLOOKUP` is stored as `_xlfn.XLOOKUP` in the file.

## Engine differences worth knowing (learned building the Postgres target and the playground)

- **DuckDB versions:** 2.0 accepts only `lambda x: ...`; 1.2 only `x -> ...` (no common syntax), so the playground's
  bridge rewrites lambdas for DuckDB < 1.3. `list_concat` with more than two lists needs DuckDB 1.3 (cellproof nests
  pairs). Hex text cast to HUGEINT fails before DuckDB 2.0 (BIGINT works everywhere). Python 3.10 cannot install
  DuckDB 2.0, so CI's 3.10 jobs run DuckDB 1.5: a good early warning for version-specific SQL.
- **Pyodide:** the newest DuckDB built for it is 1.2 (Pyodide 0.27, Python 3.12); micropip-installing it crashed in
  Chromium ("null function" on the first `read_csv`), unpacking the wheel works.
- **Postgres:** unnamed columns are all called `?column?`, so CREATE TABLE AS fails (the dbt models name them);
  `max()` has no boolean form (`bool_or`); a bare `NULL` in a CTE is text; sqlglot writes function names in upper case
  (fine: Postgres folds them).

## Open work

- **PR #5 (overflow as #NUM!, imshah-id):** the right approach (checked per operation, so IFERROR sees it; regions,
  cells and arrays agree; CI green). Changes requested: drop the COUNTIF/SUMIF error-criterion change (Excel likely
  counts cells holding the error; unconfirmed, so it must stay UNRESOLVED), and measure SQL size (the check repeats
  the expression per level) and the 200k benchmark. Merging it closes the "arithmetic overflow gives `inf`" gap.
- **Issues:** #1 overflow (PR #5), #3 TAKE/DROP/CHOOSECOLS/TOCOL (help wanted, with a contributor spec), #4 "send a
  workbook cellproof gets wrong" (pinned), #14 the 2.0 roadmap (ranked UNRESOLVED table).
- **Merged, kept for context:** PR #7 (DEC2BIN/HEX/OCT and back, MilindLate, 2026-11-13, closes #2; the Python 3.10
  hex-cast fix and the guards for cases not checked in real Excel `HEX2DEC("")` and a negative invalid `places` came
  in with it).

## Known gaps

- Arithmetic overflow gives `inf`, an honest FAIL instead of `#NUM!` (PR #5).
- A lookup outside a region (cell mode) runs about 7 s per 1,000 cells.
- Running totals other than sums (`=C1*1.05+B2`, `=MAX(C1,B2)`), other functions over growing ranges, and a split
  piece that needs a second split stay in cell mode.
- Compile on the 200k run is 11.3 s (the per-cell dependency loop and reference resolving); read is 8.3 s (openpyxl
  loads the workbook twice).
- Minors: workbook text equal to the placeholder `__cellproof_cells__` is rewritten; computed-answer results are not
  counted separately; old tables stay in `cellproof.duckdb` after a rerun.
- Not checked in real Excel yet (UNRESOLVED until they are): number -> text below 1E-11, `(-8)^(2/3)`, `(-16)^(1/4)`,
  `DATE(1900,13,1)`, COUNTIF with an error as the criterion, `HEX2DEC("")`.

## Next

`main` is 1.1.1 and about to be released: 1.1.0 (`cellproof sql` in PR #19, `cellproof watch` in PR #20, merged
2026-11-17) plus `--fast`, the one-pass read path, which closes the last item of "Rest of 1.1". The order below is
what makes it a tool people reach for daily (remaining 1.2) before the function-coverage work (2.0, issue #14).

1. **1.2 - trust.**
   - `cellproof prove book.xlsx`: a self-contained proof bundle (the workbook's hash, cellproof's version, the SQL
     for every cell, Excel's saved value, the result, and the report), so a converted model can be audited later
     without trusting the machine that ran it.
   - Stale-cached-value detection: a formula cell whose saved value does not match its formula (Excel never
     recalculated after an edit) - today those read as FAILs, which teaches people to ignore FAILs.
   - `cellproof diff old.xlsx new.xlsx`: what changed between two versions of a workbook, and which SQL cells moved.
2. **Merge PR #5** (overflow as `#NUM!`) once its measurements are in - the last known-wrong-answer gap.
3. **2.0 roadmap (issue #14), in order** (the cross-pairs landed first, PR #15):
   - **TTEST / T.TEST** — 384 UNRESOLVED cells, and **every** TTEST formula in the corpus sits inside an **array
     (CSE) block** (`test.xlsx` holds 17 `<f t="array">` blocks such as
     `TTEST(C37:D38,E35:E38,{1;2;1;2;1;2},{1;2;3;1;2;3})`, each spilling 6 rows). So this needs the
     ranges-in-a-CSE-block behaviour settled against Excel's saved values first, plus a new numerical core.
   - Approximate-match evidence hunt (~617 cells): find real-Excel saved answers for the lookup modes currently left
     UNRESOLVED, then implement only what is proven.
   - `CONVERT` (3,593 cells) stays the contributor-bait item (the largest single reason in the corpus).
   - `CHAR` > 127 from a Mac-calculated file stays honest UNRESOLVED.
4. TAKE/DROP/CHOOSECOLS/CHOOSEROWS/TOCOL/TOROW (issue #3).
5. More real analyst workbooks: each one found gaps the public files did not.
6. Other warehouses through the same translation path as Postgres (ask launch users which one).
7. Plan 3: an AI retry loop for failed cells (it sees formula structure only, never values), every answer still
   verified against Excel.
