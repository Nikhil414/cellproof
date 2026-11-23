# Trust (1.2): prove, saved-value diagnosis, diff

## Goal

Three things that make a cellproof result something a person can act on and someone else can check later, without
having to be in the room when it ran:

1. `cellproof prove book.xlsx` writes a proof bundle: the workbook's hash, the tool versions, one runnable SQL script
   that answers every cell, and the report. `cellproof prove --check dir` re-checks that bundle without trusting the
   machine that made it.
2. A FAIL stops being one undifferentiated pile. The report says which FAILs are worth looking at (their inputs are
   already right) and which only fail because a cell they read is wrong, and it warns when the file itself says its
   saved values cannot be trusted.
3. `cellproof diff old.xlsx new.xlsx` says what changed between two versions of a workbook, and which SQL moved.

## Why

Today a FAIL means "the SQL and Excel's saved value disagree", and that one sentence covers three unrelated problems:
our translation is wrong, the file was never recalculated after an edit, or a cell upstream is wrong and the
difference travelled. A report that shows all three the same way teaches people to ignore FAILs, and the corpus runs
show plenty of each. `prove` and `diff` are the other half: a result you can hand to someone, and a result you can
compare with yesterday's.

## Principles

- **No new status.** PASS, FAIL and UNRESOLVED stay the only three. A stale cached value cannot be proven stale - we
  cannot see the inputs Excel used before the edit - so calling it anything other than FAIL would be a guess. What we
  can prove is stronger and narrower, and that is what gets said (see piece 2).
- **The bundle must be checkable without cellproof's cooperation.** The SQL in it runs on a bare DuckDB; the hash
  proves which workbook it belongs to; `--check` re-derives everything it can.
- **Report changes are additive.** `report.json` gains keys; no existing key changes meaning. Everything already
  reading `report.json` keeps working.
- **DuckDB only in 1.2.** A Postgres proof would have to record translated SQL that is not one runnable script yet;
  that waits until the translation path produces one. The manifest records which engine produced the verdicts, so a
  Postgres proof can be added later without changing the format.
- **Console output stays ASCII.** The bundle's files are UTF-8 like every other output.
- **Measure before claiming speed.** The FAIL split runs on every `convert`, so it has a budget (see Cost).

## Piece 1: `cellproof prove`

### What it writes

`cellproof prove book.xlsx` verifies the workbook exactly as `convert` does (regions on, DuckDB), then writes a
folder (default `<workbook>_proof`):

| File | What it is |
|---|---|
| `cells.sql` | Every input cell's saved value, one statement per batch in dependency order, then a `SELECT` over every cell. The SQL the verifier ran, as one script. |
| `report.json` | What `convert` would have written: each cell's status, Excel's saved value, our answer, the reason. |
| `REPORT.md` | The same for people. |
| `proof.json` | Identity and seal: see below. |

`cells.sql` is what makes the bundle a proof rather than a claim. It needs nothing but DuckDB: the workbook's hash
says which file it belongs to, the SQL says how every number was made, and `report.json` says what came out. The
chain is `query.build(book, compiled, selection(book, None, None))`, which already emits exactly this shape
(`CREATE OR REPLACE TABLE cells`, the seed rows, one `INSERT` per batch, a final `SELECT`) and already seeds every
selected input cell, not only the ones a formula reads. Checked against the three example workbooks before this spec
was written: the script runs on a fresh DuckDB and returns rows identical to `verifier.execute`, cell for cell
(`commission`, `errors`, `functions` - 0 mismatches, 0 absent; the largest is 97 cells and 938 KB of SQL).

When a cell could not be translated it is absent from `cells.sql`, and `report.json` carries its UNRESOLVED reason;
the script's header counts them, so an auditor sees the hole instead of a short answer. The header is a comment, so
the script still runs.

`prove` calls `compute.fill_missing` exactly as `convert` does, because a workbook written by a script has no saved
values at all and would otherwise prove nothing. But an answer the `formulas` library supplied is not Excel's
answer, and a bundle that let the two look alike would be the false PASS this project fears. So the count goes in the
manifest (`"computed_answers": 12`), `report.json` keeps the per-cell `"answer": "computed"` it already writes, and
`REPORT.md` repeats the sentence `convert` already prints. The claim a bundle makes is then exact: *these* values are
Excel's, *those* are a library's.

### `proof.json`

```json
{
  "format": "cellproof-proof/1",
  "workbook": {"file": "book.xlsx", "bytes": 481203, "sha256": "..."},
  "cellproof": "1.2.0",
  "engine": {"name": "duckdb", "version": "2.0.0"},
  "python": "3.12.7",
  "platform": "macOS-15.6-arm64",
  "created": "2026-11-21T09:14:02Z",
  "saved_values": {"calc_mode": "manual", "full_calc_on_load": true, "calc_id": 191029,
                   "last_saved_by": "Microsoft Excel"},
  "counts": {"PASS": 41203, "FAIL": 3, "UNRESOLVED": 811},
  "computed_answers": 12,
  "files": {"cells.sql": "...", "report.json": "...", "REPORT.md": "..."},
  "bundle": "..."
}
```

- `saved_values` is what the file says about its own cache. `last_saved_by` comes from `docProps/app.xml`, which
  `reader._app` already reads for the "saved by LibreOffice" finding; the three `calcPr` attributes are new - one
  regex over `xl/workbook.xml`, the same zip open `_mark_dynamic` already does. Missing attributes are `null`, never
  guessed. Recording it here means an auditor sees the state of the file rather than our opinion of it, and piece 2
  warns from the same values.
- `files` holds the sha256 of each sibling file. `bundle` is the sha256 of the sorted `name:hash` lines, so one
  number covers the set. `proof.json` cannot hash itself; nothing depends on it being sealed.
- `created` is the only value that differs between two runs on the same workbook and the same version, so the
  manifest of two runs is otherwise byte-identical - a property worth having, and worth a test.

### `cellproof prove --check dir [--workbook book.xlsx]`

1. Read `proof.json`; refuse an unknown `format`.
2. Recompute every sibling's sha256 and `bundle`; any mismatch is a hard stop ("the bundle was changed after it was
   written").
3. With `--workbook`, hash it; it must match `workbook.sha256`. Then re-read the workbook and re-derive every
   formula cell's saved value from it, comparing with `report.json`'s `expected` - so a bundle cannot pair a report
   with a workbook that did not produce it.
4. Run `cells.sql` on a fresh DuckDB and compare each formula cell's row with `report.json`'s `actual`: same kind,
   and numbers equal within the tolerance `verifier.matches` already uses.
5. Exit 0 when every check passed; 1 when a cell's value moved; 2 when a file, hash or format is wrong. Print one
   line per kind of check (`4 files intact`, `cells.sql reproduces 41,206 of 41,206 answers`), ASCII only.

Step 4 is the part that does not need cellproof at all: anyone with DuckDB, the workbook and the bundle can run
`cells.sql` themselves and compare against `report.json` with their own eyes. `--check` is the automatable version of
that same act, and it is written to do nothing more - so a disagreement between the two would be a bug in `--check`,
not a difference of opinion.

Comparing with the same tolerance as `matches` matters: two notions of "equal" would let a bundle pass `--check` and
fail `convert`. So `matches` is split into a cell-shaped wrapper and a plain `same(kind, num, txt, other_kind, other)`
that both callers use. The wrapper keeps its current signature and behaviour, so nothing that calls it today changes.

### What `prove` does not do

- It does not run `dbt build` and does not write a dbt project. A proof is about the workbook, not about a delivery.
- It does not verify on Postgres in 1.2 (see Principles).
- It does not sign anything. The seal detects accidental and casual tampering, not a determined forger with
  write access to both the bundle and the report; the spec says so in one line at the top of `REPORT.md` rather than
  letting the word "proof" overpromise.

## Piece 2: saved values that do not match their own inputs

### Three kinds of FAIL, told apart

A FAIL is a disagreement. What decides how much it means is whether the cells it read were themselves right:

- **Primary FAIL** - no cell it reads, directly or through other formula cells, is a FAIL. Its inputs were right, so
  the disagreement is about this cell: either our translation is wrong, or the value Excel saved for it never matched
  its formula. These are the ones to look at first.
- **Inherited FAIL** - it reads at least one FAIL. The difference may have travelled; the cell may be perfectly
  translated. Attributed to the primary FAILs it reads (nearest first), so fixing the top one is the move.
- **Not a FAIL at all, and never counted here** - a cell whose expected value came from the `formulas` library
  (`compute.fill_missing`) has no saved value to be stale. `report.py` already knows these cells (`computed`); the
  split excludes them and says why when asked.

The definition is exact and cheap to check. `impact.readers_of(book, keys)` already resolves every formula's
references in one pass over the lexed formulas (with a bisect per column, so `A:A` costs O(log n)); it just filters
its output down to the keys it was given. Called once with every formula cell's key, it returns the whole read graph
inverted - each cell to the cells that read it. One breadth-first walk out from the FAIL set then marks exactly the
inherited cells; the rest of the FAILs are primary.

The cost is the honest risk here: that one call keeps an entry per formula cell, where today's "Wrong answers" section
calls it with a handful of FAIL keys. So the PR measures memory and time on the 200k benchmark before claiming the
2 s budget, and if the full inverted graph is too heavy, the fallback is a memoised walk upstream from each FAIL over
`query.formula_reads` edges, which touches only what the FAILs actually read.

### What the report says

`report.json` gains one key:

```json
"failures": {"primary": [{"cell": "Model!F42", "expected": 100.0, "actual": 99.5, "dependents": 240}],
             "inherited": 241, "computed_answers_excluded": 12}
```

`REPORT.md` replaces the flat "Failed cells" table with two: "Failures to look at" (primary, with how many cells read
each) and one line saying how many other FAILs only read those. The existing "Wrong answers other cells read" section
stays - it is the same idea from the other side, and it is already in the wild.

### The file's own evidence

When `xl/workbook.xml` says the cache cannot be trusted, the report says so above the failure tables:

- `calcMode="manual"` - Excel was told not to recalculate; the saved values are whatever the last full calculation
  left.
- `fullCalcOnLoad="1"` - Excel itself will recalculate everything when the file is next opened.
- `calcId` missing or `0` - the file does not carry a calculation-engine stamp, so it was probably not last written by
  Excel.
- `last_saved_by` naming a program that is not Microsoft Excel. `compile_book` already emits this as a finding;
  here it is turned into an instruction ("its saved values follow that program's rules"), not a second finding.

When any of these hold, `REPORT.md` says: *the file says its saved values may be out of date; recalculate (F9) and
save in Excel, then run again. Until then a FAIL is not evidence that the SQL is wrong.* When none hold, nothing is
printed - no speculation about files that are fine.

### Never the word "stale" on a cell

We cannot prove a cell is stale, and the cost of pretending is the worst bug we have: a real translation error
labelled "the file's fault" would be a wrong answer wearing a reassuring badge. So no cell is ever marked STALE; the
words used are "Excel saved X, the formula gives Y", plus the file-level evidence above. The improvement is not a new
label, it is the split - 3 primary FAILs and 240 inherited ones reads very differently from 243 FAILs.

## Piece 3: `cellproof diff old.xlsx new.xlsx`

Both workbooks are loaded and verified the way `convert` does it (region mode), because those are the verdicts a
person would act on. Only the SQL comparison uses a second compile in **cell mode**
(`compile_book(book, regions=False)`): there, a formula's SQL depends on its own text and its inputs, not on which
neighbouring columns happened to group into a region, so a region reshuffle does not read as 200 statements changing.
Both compiles reuse the same loaded book, so the extra cost is compile time only.

`diff` prints and writes (`<out>/diff.json`, plus a `DIFF.md` when the difference is more than a summary):

- **Cells added and removed**, by A1 key, with counts.
- **Input cells whose saved value changed**, with both values.
- **Formula cells whose text changed.**
- **Formula cells whose SQL changed while their text did not** - the formula looks the same but what it reads moved.
  This is the one that surprises people, and it is why comparing SQL is worth the compile.
- **Verdicts that changed**, per cell: PASS -> FAIL, FAIL -> PASS, UNRESOLVED -> PASS, and so on, with the reason.
  A new FAIL is the headline; it is printed first and sets exit code 1.
- **Region changes**: the columns that verified as a region in one version and not the other, from
  `compiled.notes` (which already names each rejected region and why) - not from diffing region SQL, which moves for
  reasons nobody asked about.

Cells are matched by sheet name and A1 address only. A renamed sheet reads as everything removed and everything
added, which is the truth and is said in the output rather than guessed around.

`diff` needs the cell-mode compile for the SQL comparison only; the verdicts still come from the normal region-mode
run, because that is the run `convert` makes and the run a person would trust. So `diff` reports what will happen
when they re-run `convert`, not what would happen in a mode nobody uses.

Exit codes follow `convert`: 1 when the new version has a FAIL the old one did not, 0 when it does not. A difference
by itself is not a failure. `diff` is a question, not a gate.

## Cost

- Piece 2 runs on every `convert`: one extra pass over the lexed formulas (`readers_of`, already paid for by the
  "Wrong answers" section) plus one inversion and two walks. Budget: under 2 s added on the 200k-formula benchmark
  (`tests/test_perf.py`), measured before the PR, and skipped entirely when there are no FAILs.
- Piece 1 adds no verification work (it reuses the run `convert` already does) and writes one file the size of the
  compiled SQL: roughly 10 MB for 200k cells. Stated in the command's help; no compression, because a proof should be
  readable.
- Piece 3 is deliberately two loads, two compiles and two verifications. It is a deliberate command, not a hook.

## Testing

Tests come first, and each piece lands as its own PR.

Piece 1:
- `cells.sql` from a workbook with inputs, a formula, a running total and an untranslatable cell runs on a fresh
  DuckDB and returns the same rows as the verifier's `execute` (`tests/xlsxkit.book_of` fixtures).
- A bundle's four files exist, every hash matches, and `bundle` matches the sorted list.
- `--check` passes on a fresh bundle; fails with exit 2 when one byte of `cells.sql` changes; fails with exit 1 when
  a recorded `actual` is edited to a different number; fails with exit 2 when `--workbook` names a different file.
- Two `prove` runs on the same workbook differ only in `created`.
- An apostrophe in a sheet name and a non-ASCII text value survive the round trip through `cells.sql`.
- `same()` gives the same verdicts as the current `matches()` on the existing test table (error codes, booleans,
  strings, blank, the tolerance edges).

Piece 2:
- A FAIL reading only inputs is primary; a cell reading it is inherited and attributed to it; a three-deep chain
  attributes to the first. A cell whose expected value is computed is excluded and counted separately.
- A `#DIV/0!` FAIL that other cells read is primary even though many cells read it.
- `calcMode="manual"`, `fullCalcOnLoad="1"`, absent `calcId` and a non-Excel `last_saved_by` each produce the warning
  sentence; a clean Excel file produces none. Fixtures go through `tests/xlsxkit.write_xlsx`, extended to write a
  `<calcPr>` (it currently writes none).
- A workbook with warnings and zero FAILs prints no failure tables and no stale sentence.
- The 200k benchmark's extra time is measured, in the PR body, against the 2 s budget.

Piece 3:
- Added, removed and changed-input cells on a two-version fixture.
- A formula whose text changed and one whose text did not but whose SQL did (an input moving under a range).
- A new FAIL sets exit code 1 and appears first.
- Renaming a sheet reports everything removed and added, and does not crash.

All three: the full suite stays green on DuckDB 2.0.0 locally and 2.0.1 in CI, the corpus counts in PROGRESS move only
where the new information explains the move, and `--check` is listed in `README.md`.

## Order

1. Piece 1 (`prove`) - no change to any existing path.
2. Piece 2 (FAIL split) - touches `impact.py`, `report.py`, the verifier's tolerance helper; no CLI change beyond a
   `--stats` line.
3. Piece 3 (`diff`) - a new module and a new subcommand, reusing what 1 and 2 built.

Version 1.2.0 after the third. `PROGRESS.md` and `README.md` are updated with each piece, not at the end.
