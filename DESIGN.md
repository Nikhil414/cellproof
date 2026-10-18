# unexcel — Design (v1)

> Status: **ACCEPTED 2026-10-11** (architecture, error handling, testing validated with user).
> Working name. `unexcel` clashes with a CRAN R package / `drhrf/unexcel` repo — rename before launch.
> Date: 2026-10-11

---

## 1. Understanding summary

- **What:** `pip install` CLI. Input: Excel workbook (`.xlsx` / `.xlsm`). Output: dbt project that recomputes the workbook's numbers + a report proving every output cell matches.
- **Why:** Business logic trapped in workbooks one person understands. Manual migration takes weeks, and nobody proves the numbers still match. Verified 2026-10-11: no open-source tool compiles Excel *formulas* to SQL/dbt. Existing tools copy *data*, not logic.
- **Who:** Analysts, finance/ops teams, analytics engineers moving off Excel.
- **Requirements from the user:** runs in anyone's environment, fast, reliable.
- **Non-goals (v1):** VBA/macros, pivot tables, Power Query, charts, external links, Power BI output, web UI, `.xls`/`.xlsb`.

## 2. Assumptions

1. Workbooks up to ~50 sheets / ~200k formula cells. Target: under 5 min on a laptop. Measure it, don't promise it.
2. v1 supports the ~80 most common functions. Every unsupported cell is reported as `UNRESOLVED`, never silently skipped.
3. Match rule: numbers within 1e-9 relative, text exact, blank vs. 0 vs. error compared explicitly.
4. Answer key = the cached values Excel saved in the file. If they're missing (file written by a script), compute them with the `formulas` library and label them `computed`, not `excel-saved`.
5. Deterministic: same workbook in, same output out. AI answers are cached on disk.
6. Solo maintainer, 4 weeks × ~15 h. Interview 5 analysts first; continue if at least 2 want to try it.

## 3. Decision log

| # | Decision | Alternatives | Why |
|---|---|---|---|
| D1 | Build unexcel | MetricProof, sqlport, whydrop-crew, pbi-doctor, onemetric | Verified gap, exact answer key, biggest audience |
| D2 | Output a dbt project | plain SQL, pandas, PBI model | What teams ship; dbt runs on every warehouse |
| D3 | CLI core, Hermes optional | Hermes required; two repos | Anyone can try it in a minute |
| D4 | Rules first, AI for the rest | LLM-first; no AI | Fast, repeatable, free |
| D5 | AI sees structure only | full values; opt-in flag | Finance data never leaves the machine |
| D6 | Laya (local, Apache-2.0) default; Jev optional | Jev default | Jev is a closed, early-access cloud API |
| D7 | Max 3 retries + no-progress stop + `UNRESOLVED` | max 5; loop until pass | Bounded and honest |
| D8 | **Approach A:** deterministic compiler + thin captain loop on failures; cell-mode (C) fallback for non-table regions | B agent-native; C cell-only | B fails fast/reliable/portable; C alone gives unreadable SQL |
| D9 | Verify loop on DuckDB directly; `dbt build` once at the end | dbt on every retry | dbt start-up costs seconds per call; the loop must be fast |

## 4. Architecture

```
 .xlsx ─► reader ─► graph ─► regions ─► compiler ─► emitter ─► verifier ─► report
                                 ▲          ▲                      │
                                 │          └──── captain ◄────────┘ (failed / unsupported cells)
                                 └── Decider adapter (Laya/Jev)  └── Translator adapter (any LLM)
```

Plain-code pipeline. The captain only wakes for cells the rules can't handle. Agents get narrow, typed jobs.

### Components (one module each)

| Module | Job | Built on |
|---|---|---|
| `reader` | Load the workbook twice: formulas + cached values. Resolve named ranges. List hidden sheets and merged cells. | `openpyxl` |
| `graph` | Tokenize formulas → cell dependency graph → topological order. Circular refs → `UNRESOLVED`. | `openpyxl.formula.Tokenizer`, stdlib `graphlib` |
| `regions` | Find tables: contiguous block, header row, same formula shape down a column (R1C1-equal). Ambiguous → ask the Decider. Low confidence → abstain → cell mode. | rules + Decider |
| `compiler` | Formula AST → SQL expression. One registry entry per function (`SUMIFS` → grouped aggregate/join, `IF` → `CASE`, `VLOOKUP`/`XLOOKUP` → left join). | `sqlglot` expression builders |
| `emitter` | Write the dbt project: input tables as seeds, one model per region, `schema.yml` tests, docs, lineage. Transpile to the target dialect. | `sqlglot`, Jinja-free string templates |
| `verifier` | Run the compiled SQL on DuckDB and map result rows/cols back to cell addresses. Compare with the answer key → `PASS` / `FAIL(evidence)` / `UNRESOLVED(reason)`. | `duckdb` |
| `captain` | For each failed group: build a task {formula, neighbour schema, failing evidence} → Translator → recompile → re-verify **only that group**. Max 3 rounds; stop if a round fixes nothing. | plain Python loop |
| `report` | `REPORT.md` + `report.json`: pass/fail/unresolved counts, per-cell status, findings. | stdlib |
| `adapters` | `Decider` (choice + probability, can abstain): Laya local default, Jev optional. `Translator`: one HTTP call to any LLM. Installed only via extras. | `unexcel[laya]`, `unexcel[llm]` |

Hermes: separate optional skill that shells out to `unexcel convert`, returns the zip + report over Telegram. Not in the core package.

## 5. Data flow (one run)

1. `unexcel convert model.xlsx --out model_dbt/ [--target snowflake]`
2. `reader` → in-memory cells (formula, cached value, type).
3. `graph` → order + cycles.
4. `regions` → tables + leftover cells (cell mode).
5. `compiler` → SQL per region; unsupported cells queued.
6. `verifier` → per-cell verdicts on DuckDB (seconds).
7. `captain` → retries on queued/failed groups (≤3 rounds).
8. `emitter` → dbt project; final `dbt build` on DuckDB as an end-to-end check.
9. `report` → summary printed; the user reviews `REPORT.md` (the human gate). Nothing is pushed anywhere.

Privacy: only formula text, headers, sheet names and types go to the adapters. Cell values never leave the process.

## 6. Error handling & edge cases

| Case | Behaviour |
|---|---|
| `.xls`, `.xlsb`, encrypted, corrupt | Clear error naming the format; exit code 2 |
| Volatile `NOW`, `TODAY`, `RAND`, `OFFSET`, `INDIRECT` | `UNRESOLVED` (non-deterministic or dynamic reference) |
| External links `[Book2.xlsx]` | `UNRESOLVED` |
| Circular references | `UNRESOLVED`, cycle listed |
| Excel errors `#N/A`, `#DIV/0!` | SQL `NULL` + error-kind column; compared by kind |
| Numbers stored as text | Copy Excel semantics (`SUMIFS` skips them) + **finding** in the report |
| Hard-coded value inside a formula column | Kept as an `overrides` seed + **finding** |
| Dates | Excel serial → date with the 1900 leap-year bug (serial 60) handled; 1904 date system detected |
| Whole-column refs `C:C` | Bound to the detected table's rows |
| Merged cells, hidden sheets | Read normally; listed in the report |
| No cached values | Compute with `formulas`, label `computed` |

## 7. Testing strategy

- **Fixtures:** small workbooks, one per function/edge case. Cached values come from a real recalculation (LibreOffice headless `--convert-to xlsx`), not hand-typed. LibreOffice differs from Excel on some functions, so a core set of ~20 fixtures is saved once in real Excel and treated as ground truth.
- **Golden test:** for every fixture, `convert` → every cell `PASS` or the expected `UNRESOLVED`.
- **Unit tests:** compiler registry (formula → SQL), region detection, date conversion.
- **Rules-only baseline:** the full suite passes with no AI installed. An AI adapter stays only if it measurably lowers `UNRESOLVED` on the fixtures.
- **Perf check:** one generated 200k-formula workbook, timed in CI.
- **CI:** GitHub Actions on Windows, macOS, Linux × Python 3.10–3.13.

## 8. Risks

1. Excel function long tail (500+) → mitigated by honest `UNRESOLVED` + a contributor-friendly function registry.
2. Region detection on messy sheets → cell-mode fallback; Decider abstains instead of guessing.
3. Dialect drift (verified on DuckDB, shipped as Snowflake/BigQuery) → `--verify-on target` (post-v1).
4. Name clash → rename before the PyPI publish.

## 9. Open questions

- License: MIT or Apache-2.0?
- Demo workbook: a public financial model, or build a messy realistic one?
- Final name.
