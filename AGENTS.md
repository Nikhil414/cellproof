# unexcel: guide for coding agents and contributors

Read `README.md` first (what the tool does, limits), then `PROGRESS.md` (current state and next steps). `DESIGN.md` is
the original spec; `docs/superpowers/specs/` and `docs/superpowers/plans/` hold the design and plan of each later
piece. `PLAN*.md` are the early plans, kept for history.

## Setup

```bash
python -m venv .venv && . .venv/bin/activate       # Windows: .venv\Scripts\activate
pip install -e ".[dev,dbt,compute]"
python -m pytest tests -q                            # dbt must be on PATH (activate the venv) or the dbt tests skip
UNEXCEL_PERF=1 python -m pytest tests/test_perf.py -s  # 200k-formula benchmark
```

## Code map (`src/unexcel/`)

| Module | Job |
|---|---|
| `model.py` | `Cell`, `Book`, `Verdict`, `Unsupported` (its message becomes the UNRESOLVED reason) |
| `reader.py` | loads the workbook twice with openpyxl (formulas, saved values); refuses zip bombs |
| `parser.py`, `refs.py` | formula text -> AST with Excel precedence; reference text -> `Area` |
| `r1c1.py` | how a reference behaves when a formula is pasted down (same row, absolute, previous row, growing range) |
| `regions.py` | finds regions (columns repeating one formula over the same rows), running totals, splits on cycles |
| `graph.py`, `fenwick.py`, `rangeindex.py` | dependency order without expanding ranges; cycle detection |
| `expr.py` | `Expr` (SQL value + SQL error code), Excel coercions |
| `compiler.py` | cell mode: formula -> SQL in batches |
| `region.py` | a region -> one set-based statement (window functions for running totals) |
| `functions.py`, `fn/` | the 93 Excel functions |
| `verifier.py` | runs everything on DuckDB and compares with Excel's saved values |
| `emitter.py` | writes the dbt project; `--dbt` runs `dbt build` |
| `report.py`, `cli.py`, `compute.py` | REPORT.md/report.json, the CLI, the optional `formulas` answer key |

## Rules

- **Never guess a value.** Anything not translated exactly raises `Unsupported("<reason a user understands>")`.
- **A false PASS is the worst bug.** When unsure, produce an honest FAIL or UNRESOLVED.
- **No DuckDB exception may escape from a supported function.** Guard every domain with `CASE`.
- **Regions must equal cell mode.** `compile_book(book, regions=False)` is the reference; region tests compare against it.
- **Runtime dependencies stay `openpyxl` and `duckdb`.** Anything else goes in an optional extra.
- **Cell values never leave the machine.** No network calls with workbook data.
- **Tests first**, console output ASCII-only (Windows consoles), measure before claiming speed.
- **Real Excel is the ground truth**; the `formulas` library is only a labelled fallback answer key.
- Commits made by an agent end with a `Co-Authored-By:` trailer for the agent.
