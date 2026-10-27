# INFO: handoff for the next coding agent

You are taking over development of **cellproof**. Read this file first, then `AGENTS.md` (rules), `README.md` (what the
tool does and its limits) and `PROGRESS.md` (state and next steps). This file is the short version of all of them plus
the things that are not written down elsewhere.

Last updated: 2026-10-27. Current release: 0.2.2 (PyPI: `cellproof`). Repo: github.com/Nikhil414/cellproof (public).

## What we are building

cellproof turns an Excel workbook (`.xlsx`/`.xlsm`) into SQL and **proves every number matches**.

1. Read the formulas and the values Excel saved in the file (Excel's saved values are a free answer key).
2. Translate every formula to SQL (DuckDB dialect), with Excel's exact rules: type coercions, error codes, error order.
3. Run the SQL on DuckDB and compare each result with Excel's saved value.
4. Emit a dbt project that recomputes the workbook, plus `REPORT.md` / `report.json`.

Every formula cell gets one verdict:

- **PASS**: SQL result equals Excel's saved value.
- **FAIL**: SQL result differs; the report shows both values.
- **UNRESOLVED**: cellproof cannot translate it exactly; the report says why (the text of an `Unsupported` exception).

## End goal

A finance or analytics person who has business logic trapped in a spreadsheet can move it to a database and **show
evidence** that nothing changed. The product is trust, not coverage: an honest UNRESOLVED beats a guessed PASS.

Longer term:

- Cover the Excel features real business workbooks use (array formulas are the biggest missing piece).
- Stay fast on big workbooks (200k formulas verifies in about 26 s today).
- Plan 3 (not started): an AI retry loop for cells that come out UNRESOLVED or FAIL. It may see **formula structure
  only, never cell values**.
- Other warehouses besides DuckDB for the dbt output (not started).

## Non-negotiable rules (from AGENTS.md)

- **Never guess a value.** Anything not translated exactly raises `Unsupported("<reason a user understands>")`.
- **A false PASS is the worst bug.** When unsure, produce an honest FAIL or UNRESOLVED.
- **No DuckDB exception may escape from a supported function.** Guard every domain with `CASE`.
- **Regions must equal cell mode.** `compile_book(book, regions=False)` is the reference; region tests compare to it.
- **Runtime dependencies stay `openpyxl` and `duckdb`.** Anything else goes in an optional extra.
- **Cell values never leave the machine.** No network calls with workbook data.
- **Tests first.** Console output is ASCII-only. Measure before claiming speed.
- **Real Excel is the ground truth.** The `formulas` library is only a labelled fallback answer key.
- Commits made by an agent end with a `Co-Authored-By:` trailer naming the agent.

## Where we are (2026-10-27)

Shipped and on `main` / PyPI 0.2.2:

- Full pipeline: reader, parser, dependency graph, cell-mode compiler, region (table) mode, verifier, dbt emitter, report, CLI.
- 96 Excel functions (list in README). Latest additions: LOOKUP, YEARFRAC (all 5 bases, dynamic basis), array constants
  such as `{0,50,100}` wherever a range is taken.
- Region mode: a column repeating one formula compiles to one set-based statement; running totals become window
  functions; VLOOKUP/HLOOKUP/MATCH/INDEX/XLOOKUP and the *IF(S) family pasted down compile to joins.
- Speed: 200k formulas in 26 s; 50k-row VLOOKUP in 4.6 s.
- Tests: about 430 pass, 8 skip (dbt tests skip when `dbt` is not on PATH). Whole suite takes 4 to 8 minutes.
- Validation: 3 workbooks saved by real Excel all PASS. 85 public test workbooks (pycel, formulas, xlcalculator):
  6,874 PASS, 10 FAIL (all in LibreOffice-saved files, flagged in their reports), 16,764 UNRESOLVED, no crashes.
  Run it with `python tests/corpus.py DIR`.
- CI: Windows, macOS and Linux x Python 3.10 to 3.13. Release = GitHub release, which triggers the PyPI publish workflow.

## What to do next (in order)

1. **Array formulas and array evaluation.** `SUM(A1:A10*B1:B10)`, `MAX(IF(c, A1:A9, 0))`, `A1*{1,2}`. Today these are
   UNRESOLVED. This is the largest remaining source of UNRESOLVED cells in business workbooks. It needs element-wise
   evaluation of ranges, which is a new design piece: write a short spec in `docs/superpowers/specs/` first, then a plan
   in `docs/superpowers/plans/`, then build test-first. Existing building blocks you can reuse:
   - Array constants are parsed into hidden-sheet `RefNode`s; the values live in an `ARRAYS` dict keyed by content hash.
     The hidden sheet name is `#array-<hash>` (see gotchas for why).
   - Regions already compile column-wise statements; element-wise range math may fit there.
2. **Test on more real analyst workbooks.** One user workbook found 4 gaps the 85 public files did not. Every real
   workbook is worth more than another synthetic test. Turn each finding into a test in
   `tests/test_corpus_findings.py`.
3. **Read time.** openpyxl loads the workbook twice (8.3 s of the 26 s). Also the per-cell dependency loop (11.3 s compile).
4. **Plan 3**: AI retry loop (structure only, never values).

Known smaller gaps are listed in `PROGRESS.md` under "Known gaps".

## Code map (`src/cellproof/`)

| Module | Job |
|---|---|
| `model.py` | `Cell`, `Book`, `Verdict`, `Unsupported` (its message becomes the UNRESOLVED reason) |
| `reader.py` | loads the workbook twice with openpyxl (formulas, saved values); refuses zip bombs |
| `parser.py`, `refs.py` | formula text to AST with Excel precedence; reference text to `Area` |
| `r1c1.py` | how a reference behaves when a formula is pasted down |
| `regions.py` | finds regions, running totals, splits on cycles |
| `graph.py`, `fenwick.py`, `rangeindex.py` | dependency order without expanding ranges; cycle detection |
| `expr.py` | `Expr` (SQL value + SQL error code), Excel coercions, sentinels such as `#1900` |
| `compiler.py` | cell mode: formula to SQL in batches |
| `region.py` | a region to one set-based statement |
| `functions.py`, `fn/` | the Excel functions (`fn/lookup.py`, `fn/dates.py`, ...) |
| `verifier.py` | runs everything on DuckDB and compares with Excel's saved values |
| `emitter.py` | writes the dbt project; `--dbt` runs `dbt build` |
| `report.py`, `cli.py`, `compute.py` | REPORT.md/report.json, the CLI, the optional `formulas` answer key |

## How to work here

Setup (Mac/Linux):

```bash
python -m venv .venv && . .venv/bin/activate
pip install -e ".[dev,dbt,compute]"
python -m pytest tests -q          # dbt must be on PATH (venv active) or the dbt tests skip
```

Workflow that has worked for every piece so far:

1. Write the failing tests first (TDD). Add a region-mode test next to every new function; it must equal cell mode.
2. Implement the smallest thing that is exactly right. If it cannot be exact, raise `Unsupported` with a plain reason.
3. Run the whole suite. After a change that can affect results, re-run the corpus and compare PASS/FAIL/UNRESOLVED
   counts. FAIL must not go up.
4. Update `README.md` (function list and limitations), `CHANGELOG.md` (under "Unreleased") and `PROGRESS.md`.
5. When Excel's behaviour is uncertain, **do not guess from memory**. Remove the unverified test case, or find an
   Excel-saved workbook that proves it. (Example: a YEARFRAC test with an unverified expected value was dropped.)

## Gotchas learned the hard way

- **Braces in sheet names break things.** Python `.format()` and dbt's Jinja both treat `{...}` as placeholders. The
  array sheet used to be `{cellproof-array-hash}` and failed in region mode and in dbt; it is now `#array-<hash>`. Never
  put `{` or `}` in generated names. Detect array sheets by membership in the `ARRAYS` dict, not by a string prefix.
- **Error order matters.** Excel reports errors left to right, and errors typed into a formula come before errors inside
  referenced cells. YEARFRAC needed the basis error checked before date errors (52 corpus FAILs until fixed).
- **Parser tests that expect "unparseable" input go stale** when the parser learns new syntax (`{1,2}` was one). Pick a
  form that is still invalid (`{1,A1}`).
- **Excel facts verified against real Excel** are in `PROGRESS.md` ("What Excel taught us") with tests in
  `tests/test_excel_findings.py`. Read them before touching COUNTIF, DAYS360, TEXT, MOD, POWER or date code.
- **Approximate lookups on unsorted data** (LOOKUP, MATCH/VLOOKUP with type 1/-1) stay UNRESOLVED. Excel's answer there
  depends on its binary search.
- **Text to number** only handles plain numbers. Locale-dependent text (`"1,000"`, `"$5"`, `"50%"`) is UNRESOLVED.
- Console output must stay ASCII (the tool is also used on Windows consoles).
- No linter is installed in the dev environment (ruff and pyflakes are missing). Keep the existing style by hand.

## Repo and release facts

- Public repo `Nikhil414/cellproof` (remote `origin`), single-commit-history start. **Never push the old history
  publicly.** A private backup `Nikhil414/unexcel-dev` exists (remote `dev`, branch `public-main` mirrors `main`).
  The project was called `unexcel` until 0.2.0; older plans and specs still use that name.
- Release: bump `version` in `pyproject.toml` and `src/cellproof/__init__.py`, date the CHANGELOG, run the tests,
  commit, push `main`, then `gh release create vX.Y.Z`. The release triggers the PyPI publish workflow. Confirm CI is
  green first.
- Only push, tag or release when the owner asks.
- The owner prefers decisions to questions: decide, build, test, security-check, and explain what you did.

## Safety checklist before every commit

- No workbook data, secrets or tokens in the repo, tests or logs.
- No new runtime dependency.
- No network call that carries workbook data.
- A false PASS is possible? Make it an UNRESOLVED instead.
