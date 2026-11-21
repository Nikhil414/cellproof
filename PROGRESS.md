# cellproof: progress

Last updated: 2026-11-21. This file is the single place that says where the project stands. Read it cold: a person
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
| PyPI | **1.1.1** is the latest release (`pip install cellproof`), published 2026-11-20 from tag `v1.1.1` (trusted publishing) |
| `main` | **1.1.1**: `--fast` (read each worksheet once, identical `Book`), on top of 1.1.0's `cellproof sql` (PR #19) and `cellproof watch` (PR #20); PR #22 also fixes TEXT's number formats on the Postgres target (sqlglot 30.33 turned `printf`/`format` into an expression the rewrite missed). Tag `v1.1.1` = the 1.1.1 release commit, `43844f7`. Plus the two DuckDB 2.0.1 fixes: PR #26 (`TEXT` agrees with `ROUND`, Postgres overflow guarded before scaling) and PR #28 (`ROUND`/`ROUNDUP`/`ROUNDDOWN`/`TRUNC` decide their boundary before scaling) |
| Release | [v1.1.1](https://github.com/Nikhil414/cellproof/releases/tag/v1.1.1), 2026-11-20: 13/13 CI checks green, `publish.yml` published to PyPI. [v1.1.0](https://github.com/Nikhil414/cellproof/releases/tag/v1.1.0), 2026-11-17: `cellproof sql` and `cellproof watch`. 1.0.0 ([v1.0.0](https://github.com/Nikhil414/cellproof/releases/tag/v1.0.0), 2026-11-17) added the Postgres target, dbt-postgres, the playground and the console triage |
| Playground | Live at https://nikhil414.github.io/cellproof/ (GitHub Pages, deployed by `.github/workflows/pages.yml` on every push to `main` that touches `src/`, `playground/` or the example) |
| Tests | `main` (1.1.1: `--fast`): **746 pass, 3 skip** (749 collected) with dbt-postgres installed and the 85-workbook corpus present (2026-11-20; the 3 skips are the optional `compute` extra and the opt-in 200k benchmark), including 85 parametrized differential tests that read every public workbook with both loaders and require the same `Book`, and the Postgres format tests. The same suite is green on **DuckDB 2.0.0 and 2.0.1** (736 pass / 13 skip in the lean CI-shaped venv: sqlglot, jinja2, dbt and the corpus are absent there). 1.1.0: 622 pass, 3 skip |
| CI | `tests` workflow: Windows/macOS/Linux x Python 3.10-3.13, plus the `postgres` job (Postgres 16 service, Postgres tests and `dbt build`): **13 checks** on every pull request. Green on the release commit |
| Public test workbooks (85) | **11,080 PASS / 16 FAIL / 37,814 UNRESOLVED**, identical on DuckDB 2.0, on DuckDB 1.2 (the playground) and on Postgres (with `dbt build` passing for all 85). The 16 FAIL are explained below. A fresh re-download of the corpus on 2026-11-15 (the source repos changed; 66 files now) measured **13,043 PASS / 34 FAIL / 19,183 UNRESOLVED** — one workbook (formulas' `test.xlsx`, a function catalog) holds 17,454 of the UNRESOLVED, and `CONVERT` alone 3,593 (19%). Of the 34 FAIL, 19 were wrong error codes from PR #7's base conversions ([issue #10](https://github.com/Nikhil414/cellproof/issues/10)), 9 are the known LibreOffice drift in `INFORMATION.xlsx`/`SQRT.xlsx`; 6 are the known npv float cells; the 19 error codes were fixed the same day (PR #12), bringing the fresh-corpus count to **13,062 PASS / 15 FAIL**. 2026-11-15: the six base cross-pairs (HEX2BIN/HEX2OCT/BIN2HEX/BIN2OCT/OCT2HEX/OCT2BIN, semantics verified against Excel-saved values in `test.xlsx`) turned **669** more cells PASS: fresh corpus now **13,732 PASS / 15 FAIL / 18,513 UNRESOLVED**, 0 FAILs on `BASES.xlsx` and `test.xlsx` |
| Contributors | Open: PR #5 (overflow as `#NUM!`; both 2026-11-07 requests are met - the criteria change is now *guarded* by `_criterion_error`, so a computed error criterion stays `#UNCHECKED`/UNRESOLVED instead of guessing Excel's answer, and `BIND_SQL` materializes oversized arithmetic - but changes requested 2026-11-20 because the binder covers cell mode only: the region path has no binder and grows quadratically, and `postgres` is red only from a stale base; both points restated 2026-11-21 as a two-item checklist), PR #23 (TAKE/DROP/CHOOSE*/TOCOL/TOROW, reviewed 2026-11-20 and given a final checklist 2026-11-21: negative indices return `#SPILL!` where Excel returns a value, `TAKE`/`DROP` refuse the 0/negative counts Excel defines, `TOCOL`/`TOROW` `ignore` 2/3 missing, and it needs a rebase (6 commits behind, `CONFLICTING`) whose `postgres.py` resolution must **drop** the PR's hunk - 1.1.1's `fmt_call` already does that job properly, and `cellproof_printf` rejects `{:,.<p>f}`, so keeping it would raise on every grouped format; all four were proven against Excel-saved cells in `test_test_files_test.xlsx` sheet `LOOKUP`). Merged: PR #7 (DEC2BIN/HEX/OCT and back, MilindLate, 2026-11-13, closes #2; unverified cases guarded, never guessed), PR #9 (`--stats` flag: count-sorted tally of UNRESOLVED reasons on the console, 2026-11-15), PR #12 (base conversions: match real Excel's error codes, closes #10, 2026-11-15), PR #15 (the six base cross-pairs, 2026-11-16, issue #14's first roadmap item), PR #17 (console triage: ranked unresolved reasons and the wrong answers other cells read, 2026-11-16), PR #8 (1.0: the Postgres target and dbt-postgres project, 2026-11-17, released as v1.0.0), PR #26 (survive DuckDB 2.0.1's `string_split('', '')` and ties-to-even `round`: `TEXT` now agrees with `ROUND` by construction and Postgres overflow is guarded before scaling, 2026-11-21), PR #28 (the same trap for `ROUND`/`ROUNDUP`/`ROUNDDOWN`/`TRUNC`, decided before scaling: no more dead batch, and `ROUND(5,-400)` is 0 on Postgres rather than 5, 2026-11-21). Closed: PR #6 (superseded by #5). Issues #1, #3, #4, #14 open, #4 pinned; #3 has a contributor spec comment (2026-11-15: TAKE, DROP, CHOOSECOLS, CHOOSEROWS, TOCOL, TOROW against the `fn/dynamic.py` relation model); #10 was fixed in PR #12 and closed 2026-11-15; #14 is the 2.0 roadmap (ranked UNRESOLVED table) |
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
| `src/cellproof/proof.py` | `cellproof prove` (1.2): `cells_sql` (the whole chain as one script), `manifest`/`seal` (the hashes), `bundle` (write the four files), `check` (`--check`: hashes, the workbook's saved values, then `cells.sql` on a fresh DuckDB) |
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
| `cellproof prove` (1.2, piece 1) | verify as `convert` does, then write a bundle a stranger can check: `cells.sql` (the whole chain, one runnable DuckDB script), `report.json`, `REPORT.md`, and `proof.json` (workbook hash and size, tool versions, the workbook's own `<calcPr>` and last-saving program, counts, a hash per file and one over the set). `prove --check DIR [--workbook f]` re-hashes, optionally re-reads the workbook and compares its saved values with the report, then runs `cells.sql` on a fresh DuckDB and compares every row - exit 0/1/2/3. The seal is not a signature and `REPORT.md` says so in one line; `verifier.matches` became a wrapper over `verifier.same` so `--check` and `convert` cannot disagree about "equal" |
| FAIL split (1.2, piece 2) | `impact.split_failures` walks the read graph once (`readers_of` with every formula's key) to mark each FAIL *primary* (reads no FAIL, so the difference starts there) or *inherited* (reads one, attributed to the nearest primary it reaches, so a three-deep chain names the first). Cells whose expected value came from the `formulas` library are excluded and counted. `REPORT.md` replaces the flat failed-cell table with "Failures to look at" plus an inherited count, `report.json` gains `failures`, `--stats` prints both; the file's own `<calcPr>`/last-saving-program evidence puts one "saved values may be out of date, recalculate in Excel" sentence above the tables, and only when there are FAILs. No verdict changes - the corpus counts do not move. Cost on the 200k benchmark: +1.0 s and ~14-33 MB when every formula fails |
| `cellproof diff` (1.2, piece 3) | what changed between two saves and what it broke. Both books are verified exactly as `convert` does, so the verdicts are real; the diff reports added/removed cells, changed input values (kind included: `2` vs `"2"`), changed formula texts, changed verdicts (a new FAIL sorts first, then UNRESOLVED, then PASS; vanished cells and newly appearing passing cells are left out), and regions that verified as a block in one version and not the other, with the compiler's note. The SQL compare comes from a second cell-mode compile, canonicalised to one text per cell (cell tags in place of batch names, `COALESCE` and `kind='e'` reads folded), so region regrouping, renumbered CTEs and moving batch boundaries are not noise - what remains is the case worth a diff: identical formula text, different answer (a named range that grew, a table that gained rows). Cells are matched by sheet+address only, so a renamed sheet reads as everything removed and everything added - the truth. `diff.json` always; `DIFF.md` only when a verdict changed or ≥10 changes, and the headline never claims to know which to fix first. Exit 1 only for a new FAIL |
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
- **DuckDB 2.0.1 changed two answers under cellproof** (uploaded 2026-11-20, after `main`'s last green CI at 11:42, so
  it surfaced only on new pull requests): `string_split('', '')` went from `['']` (length 1) to `[]` (length 0), and
  `round(x::DOUBLE, n)` went from ties-away-from-zero to ties-to-even, so `round(2.675::DOUBLE, 2)` is `2.67`, not
  `2.68`. Both were real: `PROPER("")` returned NULL instead of `''` (poisoning every cell that read it) and
  `TEXT(2.675,"0.00")` returned `"2.67"`. Fixed in code, not by pinning, because both corrections are correct on every
  version: `PROPER` now returns `''` for empty text itself, and `_fixed` scales, snaps with `xl15` and rounds a whole
  number, the way `ROUND` already did. Only the DOUBLE path drifted - `round(2.675::DECIMAL(18,3), 2)` is still `2.68`,
  and `printf('%.2f', 2.675)` was `2.67` on every version, which is why `TEXT` never leaned on it.
  Fixing this surfaced a **pre-existing** bug on all versions: `TEXT(1.005,"0.00")` said `"1.00"` while
  `ROUND(1.005,2)` said `1.01`, the answer `maths.py` documents as real Excel's. `TEXT` now goes through the same
  `xl15` snap, so the two agree by construction (a test compares them cell by cell). The value is bound once in a
  derived table: repeating it per branch made the SQL 1991 chars per cell instead of 380 (441 with the overflow
  guard below). Scaling caught a **Postgres-only** regression that the plan had not foreseen: Postgres *raises*
  `value out of range: overflow` on float8 overflow where DuckDB returns `inf`, and one such cell makes
  `postgres.execute` mark the **whole batch** UNRESOLVED (`ROUND(1e308,2)` already does this on `main` - a separate,
  pre-existing bug, left for a follow-up). `_fixed` therefore tests the magnitude **before** scaling
  (`CASE WHEN abs(sv) < 1e300 THEN abs(sv) * <shown*p> END`): a value too large to scale keeps the digits 1.1.1
  showed it with (byte-identical to `main` for `1e300`/`1e307`/`1e308`/`-1e308`/`1e-320` over six formats), and a
  percent of such a value has no answer without scaling, so it is NULL and the cell FAILs visibly instead of dying.
- **`xl15` and constant folding:** `xl15(0)` raises `Conversion Error: ... value -inf can't be cast ... INT32` when the
  argument is a bare literal, but not for a column or a `CAST(... AS DOUBLE)`. Neither `ROUND` nor `TEXT` can hit it:
  the compiler emits literals as `CAST(0.0 AS DOUBLE)`, and `_fixed` hands `xl15` the column `sc` from a derived
  table, so the argument is never a literal. `=TEXT(0,"0.00")` is a test, so that stays true.
- **Pyodide:** the newest DuckDB built for it is 1.2 (Pyodide 0.27, Python 3.12); micropip-installing it crashed in
  Chromium ("null function" on the first `read_csv`), unpacking the wheel works.
- **Postgres:** unnamed columns are all called `?column?`, so CREATE TABLE AS fails (the dbt models name them);
  `max()` has no boolean form (`bool_or`); a bare `NULL` in a CTE is text; sqlglot writes function names in upper case
  (fine: Postgres folds them).

## Open work

- **DuckDB 2.0.1 drift (fixed, PR #26, 2026-11-21):** see "Engine differences" above. Five regression tests in `test_text.py`
  (`=PROPER("")`, `=TEXT(0,"0.00")`, `=TEXT(1.005,"0.00")`, one requiring `TEXT` and `ROUND` to agree, one that a
  value too large to scale keeps its digits without losing the batch), three on the Postgres target, and the
  `postgres` rewrite test's `string_split('', '')` case replaced - that input's answer is now version-dependent, so
  asserting the shim equals live DuckDB on it was a test that could only pass on one version. The shim keeps the
  2.0.0 answer and is only ever reached with guarded (non-empty) text, which the new Postgres tests pin.
- **`ROUND`/`ROUNDUP`/`ROUNDDOWN`/`TRUNC` at the double boundary (fixed, PR #28, 2026-11-21):** they scaled with
  `abs(x) * power(10, d)` and guarded afterwards, and Postgres raises on that multiply where DuckDB returns `inf`, so
  `ROUND(1e308, 2)` - also as a literal - turned every cell of its batch UNRESOLVED on `main`. Measuring the fix found
  a second, worse bug on the same line: `power(10, -400)` underflows to `0` on DuckDB (the branch that answers `0`)
  but is `Infinity` through `cellproof_pow` on Postgres, so **`ROUND(5,-400)` returned 5 where Excel says 0** - a wrong
  answer on the Postgres target, not just a dead batch. Both are now decided before any scaling: a non-finite `x` is
  passed through, digits `> 308` give `x` (`10^309` is larger than `MAX_DOUBLE`), digits `<= -309` give `0`, and the
  product is bounded by `abs(x) < MAX_DOUBLE / greatest(p, 10) * (1 - EPS)` before the multiply. `greatest(p, 10)` is
  `p` wherever the guard is reached (`trunc(d) > 0`) and keeps the constant foldable when the digits are a literal:
  Postgres folds a dead `CASE` arm while planning, and `1.8e308 / 1e-5` overflows the planner (`NumericValueOutOfRange`
  on a query whose answer is 1). The bound is tight on purpose - the whole overflow region cannot answer `x`, because
  `ROUND(1.2345678901234567e300, 5)` is `1.2345678901234598e+300` (the 15-digit snap moves the last digits) and
  `ROUND(1e300,-5)` is `9.999999999999999e+299`. The value, the digits and `power()` are bound once in a derived table
  (`_bound`), so a `ROUND` cell got **smaller**: 533 chars per cell against `main`'s 632 (a 20k-cell region: 10.66M vs
  12.64M chars, same compile and run time, 0.16 s + 0.05 s), and region mode answers cell-for-cell what cell mode does.
  Proved equal to the old SQL two ways: 1,020 (x, digits) pairs on DuckDB 2.0.0 covering both ends of the double
  range, subnormals, `inf`, NaN and fractional digits, 0 differences; and 805 whole cells through `compile_book` +
  `execute` (ROUND/ROUNDUP/ROUNDDOWN/TRUNC over 35 values x 23 digit counts), identical answers old vs new.
  `ROUNDUP`/`ROUNDDOWN`/`TRUNC` keep their `#NUM!` for an unscalable value (a documented shortcut) but answer it
  instead of taking the batch down. Pinned by `test_rounding_at_the_double_boundary_on_postgres` and
  `test_roundup_and_trunc_at_the_boundary_keep_their_num_error_on_postgres` - both fail on `main` and pass here - plus
  the existing Excel-example tests in `test_maths.py`.
- **PR #5 (overflow as #NUM!, imshah-id):** the right approach (checked per operation, so IFERROR sees it; regions,
  cells and arrays agree; results stay correct - every measured case is still PASS, so nothing here is a false PASS).
  The two 2026-11-07 requests are met: the COUNTIF/SUMIF error-criterion change is *guarded* (`_criterion_error` maps a
  computed error criterion to `#UNCHECKED`, so it stays UNRESOLVED instead of guessing what Excel counts), and SQL size
  is measured and bounded per batch by `BIND_SQL = 1_500` bindings. Changes requested 2026-11-20 for the remaining
  half: the binder is set only in cell mode, so the region path (the default) has none and the check repeats the
  expression at every level - 10 rows x 240 terms went 33,915 chars / 0.07 s to 3,853,585 chars / 22.5 s against
  `main` (~113x size, ~320x time). `postgres` is red only from a stale base (head `9c1b079` predates #22's
  `cellproof_grouped`; rebase should clear it with no code change). Merging it closes the "arithmetic overflow gives
  `inf`" gap.
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

`main` is 1.1.1; the 1.2 work (below) finishes on this branch as 1.2.0. 1.1 was `cellproof sql` (PR #19) and
`cellproof watch` (PR #20, both merged 2026-11-17) plus `--fast`, the one-pass read path. The order below is what
makes it a tool people reach for daily (1.2) before the function-coverage work (2.0, issue #14).

1. **1.2 - trust.** Spec: `docs/superpowers/specs/2026-11-21-trust-design.md`. Three pieces, each its own PR, in order;
   1.2.0 is tagged after the third.
   - **Piece 1 (merged 2026-11-21, PR #30): `cellproof prove book.xlsx`** - a bundle (`cells.sql`, `report.json`,
     `REPORT.md`, `proof.json`) a stranger checks with DuckDB alone, plus `prove --check DIR [--workbook f]`, which
     re-hashes the files, re-reads the workbook's saved values, and runs `cells.sql` against `report.json`. Exit
     0/1/2/3. The seal detects accidents, not forgery, and `REPORT.md` says so. `Book.calc` carries the workbook's
     `<calcPr>`.
   - **Piece 2 (merged 2026-11-21, PR #31): the FAIL split.** A FAIL whose inputs were all right is a **primary**
     FAIL; one that reads a FAIL (directly or through other formulas) is **inherited**, attributed to the nearest
     primary it reaches. Today both read the same, which teaches people to ignore FAILs. Plus the saved-value
     diagnosis: quote the file's own `<calcPr>`/`last_saved_by` evidence instead of calling a cell "stale", which
     cannot be proven. `report.json` gains a `failures` key; no status is added.
   - **Piece 3 (PR open): `cellproof diff old.xlsx new.xlsx`** - what changed between two saves, and which SQL cells
     moved. Verdicts come from the normal region-mode run; the SQL compared comes from a second cell-mode compile,
     because regions may merge or split differently between the two files. Written and green on `cellproof-diff`
     (18 tests in `tests/test_diff.py`); spikes showed a bounded or whole-column range does not move cell-mode SQL
     but a named range that grows does.
2. **Merge PR #5** (overflow as `#NUM!`) once the region path uses the same `BIND_SQL` binder as cell mode (cell
   mode is already bounded, e.g. 781,211 -> 164,411 chars at 200 terms) - the last known-wrong-answer gap.
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
