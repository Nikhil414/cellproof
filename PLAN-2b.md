# unexcel — Plan 2b: The function library Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** unexcel translates the functions real workbooks use: 85 new Excel functions, chosen from published usage data. It also adds the run-time-typed values and the type conversions that make `IFERROR(..., "")` and `"Total: "&A1` work.

**Architecture:** Each function family is one module in a new `unexcel/fn/` package that registers into the existing `FUNCS` table, so `compiler.py` is unchanged apart from one import. Two foundations come first:
- **Run-time types.** An `Expr` of type `"any"` carries a number part, a text part and a kind decided by DuckDB, so a result can be text in one row and a number in another, as Excel's can.
- **The criteria engine.** One SQL implementation of Excel's `">5"` / `"<>x"` / wildcard / empty-cell rules, shared by all eight `*IF(S)` functions.

Every function returns Excel's error codes through Plan 2a's `Expr.err`.

**Setup:** every command assumes the venv is active (`source .venv/bin/activate` on macOS/Linux, `.venv\Scripts\activate` on Windows), with `python -m pip install -e ".[dev,dbt]"`.

**Tech Stack:** Python ≥3.10, openpyxl 3.1, DuckDB ≥1.1 (verified on 1.5.6), pytest. No new dependencies.

**Spec:** `DESIGN.md` §2.2 ("v1 supports the ~80 most common functions; every unsupported cell is UNRESOLVED") and the Plan 2 row of `PLAN.md`'s roadmap. Builds on Plan 2a (branch `plan-2a`, PR #2).

### Which functions, and why

Measured, not guessed. Two published corpora agree on the core:
- **Sheetpedia** (NeurIPS 2025, 300k spreadsheets). The top formula patterns after deduplication: plain arithmetic (most of all), IF, SUM, VLOOKUP, AND+IF, COUNTIF, IFERROR+VLOOKUP, TEXT, INDEX+MATCH, AVERAGE, COUNTIFS, IFERROR, SUMIFS, SUMIF, CONCATENATE, the CHAR/CLEAN/SUBSTITUTE/TRIM cleanup chain, IF+ISBLANK, LEFT, the COLUMNS/MID/REPT chain, XLOOKUP and MONTH, with WEEKDAY just behind.
- **Aivaloglou, Hoepelman & Hermans** (J. Softw. Evol. Proc. 2017; Enron, EUSES, Fuse and WikiLeaks, 8.5M unique formulas). Top 15: CONCATENATE, IF, SUM, ROUND, SUMIF, VLOOKUP, DAYS360, SUBTOTAL, COUNTIF, AND, SUMIFS, ISNA, ISERROR; AVERAGE and INDIRECT are top-5 in single corpora.

This plan covers every function on both lists except INDIRECT (volatile, deliberately UNRESOLVED) and the "reference" uses of INDEX (`A1:INDEX(...)`). It then completes each family. After it, unexcel knows 93 functions:

| Family | Functions | Task |
|---|---|---|
| Conditional | COUNTIF, COUNTIFS, SUMIF, SUMIFS (rewritten), AVERAGEIF, AVERAGEIFS, MINIFS, MAXIFS | 2 |
| Aggregate | AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, PRODUCT, MEDIAN, SUMPRODUCT, SUBTOTAL | 3 |
| Lookup | VLOOKUP (approximate added), HLOOKUP, MATCH, INDEX, XLOOKUP, CHOOSE | 4 |
| Logic & info | IFS, SWITCH, NOT, XOR, IFNA, ISNA, ISERROR, ISERR, ISNUMBER, ISTEXT, ISLOGICAL, ISBLANK, NA, TRUE, FALSE | 5 |
| Math | ABS, INT, SIGN, PI, SQRT, EXP, LN, LOG10, LOG, POWER, MOD, TRUNC, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, MROUND | 6 |
| Text | CONCATENATE, CONCAT, TEXTJOIN, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, PROPER, SUBSTITUTE, REPLACE, FIND, SEARCH, REPT, VALUE, TEXT, EXACT, CHAR, CLEAN | 7 |
| Date | DATE, YEAR, MONTH, DAY, WEEKDAY, DAYS, EDATE, EOMONTH, DAYS360, DATEDIF | 8 |
| (from Plans 1/2a) | SUM, IF, ROUND, AND, OR, IFERROR | — |

### Measured while planning (2026-10-11, DuckDB 1.5.6)

- **October 12 local branch validation:** 146 passed, 4 skipped, including dbt. The 200k benchmark passed in 263.3 s (read 1.8 s, compile 5.6 s, verify 255.3 s, emit 0.7 s; 801 batches). Three skipped checks require Excel-saved files; the fourth is the separately executed opt-in benchmark.
- **The original draft passed 97 tests, including a real `dbt build`.** The code below includes the October 12 review fixes and a separate regression gate after Task 8. The historical 200k timing predates these fixes; Task 9 must re-measure it.
- **The `formulas` library is not a usable second opinion.** On 20 probes it contradicted Microsoft's documentation several times: it says TRIM keeps inner spaces, that AVERAGE fails on text, that COUNTIF with `"<>5"` and `">b"` counts nothing, and it writes `1E+15` as `1000000000000000`. Uncertain behaviour therefore goes to a real-Excel workbook (Task 9), not to `formulas`.
- **DuckDB throws on `ln(0)`, `sqrt(-1)` and `log10(0)`, and on `make_date(2026, 14, 1)`.** One throw stops the whole batch, so every function with a domain is a guarded CASE. A CASE never evaluates the branch it skips, even for constants.
- **`printf('%.15g')` matches how Excel writes numbers in text.** `1/3` gives `0.333333333333333`, `0.1+0.2` gives `0.3`, `1e18` gives `1E+18` after `upper()`.
- **A text criterion built from cells (`">"&B1`) and inlined into every comparison** grew a COUNTIF to 137k characters of SQL and broke DuckDB's expression depth. Criteria are therefore parsed once, in a one-row table.
- **Text-to-number uses a strict pattern, not `TRY_CAST`.** `TRY_CAST('inf' AS DOUBLE)` is `inf` in DuckDB, where Excel's answer is `#VALUE!`.
- **There is no `initcap`, so PROPER is built from list functions.** `regexp_escape` escapes `~`, so SEARCH handles `~*` and `~?` after escaping.

### Deliberate limits (each UNRESOLVED with a reason, or an honest FAIL)
- Array expressions: `SUMPRODUCT((A1:A9="x")*B1:B9)`, array formulas, and reference results (`A1:INDEX(...)`). These are Plan 2c/later.
- Text-to-number reads plain numbers only. Excel also reads `"1,000"`, `"$5"`, `"50%"` and dates typed as text; unexcel gives `#VALUE!` there, which shows as a FAIL, never a false PASS.
- Not yet supported: criteria taken from an empty cell; SUBTOTAL 101-111 (hidden rows); TEXT formats outside the table; DATEDIF `"MD"`/`"YM"`/`"YD"`; WEEKDAY types 11-17; the 1904 date system; TEXTJOIN keeping the empty cells of a range; COUNT and MEDIAN of values typed into the formula.
- The quadratic slowdowns found in Plan 2a's review (running `SUM($B$1:B{r})`, per-row lookups) are Plan 2c.

## Global Constraints

- Python `>=3.10`. Runtime dependencies: **only** `openpyxl>=3.1`, `duckdb>=1.1`. No new dependency in this plan.
- Match rule (unchanged): numbers `math.isclose(rel_tol=1e-9, abs_tol=1e-12)`; text exact (case-sensitive); booleans exact; an Excel error matches only the same error code.
- Never guess: any function, argument form or format not covered → `UNRESOLVED` with a human-readable reason naming it.
- No DuckDB exception may escape from a supported function on any input: guard every domain with CASE.
- Every numeric literal is `CAST(x AS DOUBLE)`; division is `a / NULLIF(b, 0)` plus `'#DIV/0!'`.
- Cell table schema unchanged: `sheet, row_num, col_num, kind (n/s/b/e), num, txt`. A text cell's `txt` may come back NULL from the seed CSV when it was `""`: always read it as `COALESCE(txt, '')`.
- The verifier and the dbt project run the same batch SQL.
- Performance target (DESIGN §2.1): 200k formula cells under 5 minutes; Task 9 re-measures.

## Review Focus

1. **A formula returning text in some rows and a number in others** (`IFERROR(VLOOKUP(...), "")`, `IF(A2="", "", A2*2)`) feeding another formula and a SUM range → each cell gets its own type; SUM skips the text. Test in Task 1.
2. **A criterion built from cells** (`COUNTIF(A:A, ">"&B1)`) → parsed at run time, no DuckDB depth error. Test in Task 2 (`S!D1&S!D2`).
3. **Text that only looks like a number** (`"1,000"+1`, `" 3 "*2`) → Excel's reading or `#VALUE!`, never a silently wrong number. Test in Task 1.
4. **Math on bad input** (`LN(0)`, `SQRT(-1)`, `MOD(5, 0)`) → Excel's error code, never a DuckDB exception that stops the batch. Test in Task 6.
5. **Dates on Excel's odd calendar** (serial 60 = the fake 1900-02-29, serial 0, DATE month roll-over) and 1904-system workbooks → Excel's values, or UNRESOLVED. Test in Task 8.

---

## File structure

```
src/unexcel/
  expr.py          + "any" type, conversions, unify_all, pick, compare                 (Task 1)
  compiler.py      any-typed cells in batches; `import unexcel.fn`                      (Tasks 1, 2)
  functions.py     registry + SUM IF ROUND AND OR IFERROR; SUMIFS/VLOOKUP move out       (Tasks 1, 2, 4)
  refs.py          C:C means every row                                                 (Task 1)
  model.py         Book.date1904                                                       (Task 8)
  reader.py        sets date1904                                                       (Task 8)
  compute.py       keeps date1904                                                      (Task 8)
  fn/__init__.py   imports every module below                                          (Tasks 2-8)
  fn/criteria.py   criteria engine + *IF(S)                                            (Task 2)
  fn/aggregate.py  AVERAGE ... SUBTOTAL                                                (Task 3)
  fn/lookup.py     VLOOKUP ... CHOOSE                                                  (Task 4)
  fn/logic.py      IFS ... FALSE                                                       (Task 5)
  fn/maths.py      ABS ... MROUND                                                      (Task 6)
  fn/text.py       CONCATENATE ... CLEAN, TEXT number formats                          (Task 7)
  fn/dates.py      DATE ... DATEDIF, TEXT date formats                                 (Task 8)
tests/
  test_types.py test_criteria.py test_aggregate.py test_lookups.py test_logic.py
  test_maths.py test_text.py test_dates.py                                             NEW
examples/
  make_functions.py   the real-Excel ground-truth workbook for this plan                (Task 9)
```

---

### Task 1: Run-time types and Excel's type conversions

**Files:**
- Replace: `src/unexcel/expr.py`, `src/unexcel/compiler.py`, `src/unexcel/functions.py`
- Modify: `src/unexcel/refs.py`, `tests/test_compiler.py`, `tests/test_functions.py`, `tests/test_refs.py`
- Test: `tests/test_types.py`

**Interfaces:**
- Consumes: Plan 2a's `Expr(sql, type, err)`, `errs`, `error`, `num`, `txt`, `as_bool`, `unify`, `render_batch`, `Compiler.cell_value`.
- Produces:
  - `Expr(sql, type, err=NO_ERR, txt="NULL", kind="NULL")`; new type `"any"`: `sql` is its number part (TRUE/FALSE as 1/0), `txt` its text part, `kind` SQL giving `'n'`/`'s'`/`'b'`. `expr.any_of(kind, num_sql, txt_sql, err)`, `expr.as_any(e)`.
  - Excel conversions: `num(txt)` reads `' 3 '`, `-1.5e3` (else `#VALUE!`); `txt(num)` writes 15 significant digits; `txt(bool)` is `TRUE`/`FALSE`; `as_bool(txt)` reads `"TRUE"`/`"FALSE"` (else `#VALUE!`). Helpers `expr.parse_num(sql)`, `expr.num_text(sql)`.
  - `expr.unify(a, b)` / `expr.unify_all(list)`: one type, or `"any"` when types differ (never raises "mixes").
  - `expr.pick(cond_sql, yes, no, err)`: a CASE over every part of two unified results.
  - `expr.compare(a, b, op)`: Excel comparison (text ignores case; across types number < text < TRUE/FALSE; an empty cell acts as 0, "" or FALSE).
  - Batch CTEs gain columns `t` (text part) and `k` (kind) for `"any"` cells; `KIND["any"] = "a"`; `Compiler.kind_of` returns `"a"` for any-typed formula cells.
  - `refs.MAX_ROW = 1_048_576`: `C:C` now means every row, as in Excel (`Area.present` keeps it fast).

- [ ] **Step 1: Write the failing tests `tests/test_types.py`**

```python
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute, verify


def run(sheets):
    book = book_of(sheets)
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def test_if_and_iferror_may_return_text_or_a_number():  # the most common mixed-type patterns
    got, c = run({"S": {
        "A1": 5, "A2": -1, "A3": 0,
        "B1": ('=IF(A1>0, "yes", 0)', "yes"), "B2": ('=IF(A2>0, "yes", 0)', 0),
        "B3": ('=IFERROR(1/A3, "")', ""), "B4": ('=IFERROR(1/A1, "")', 0.2),
        "C1": ('=B1&"!"', "yes!"), "C2": ("=B2+1", 1), "C4": ("=SUM(B1:B4)", 0.2),
    }})
    assert c.unresolved == {}
    assert got[("S", 1, 2)] == ("s", None, "yes") and got[("S", 2, 2)] == ("n", 0.0, None)
    assert got[("S", 3, 2)] == ("s", None, "") and got[("S", 4, 2)] == ("n", 0.2, None)
    assert got[("S", 1, 3)] == ("s", None, "yes!") and got[("S", 2, 3)] == ("n", 1.0, None)
    assert got[("S", 4, 3)] == ("n", 0.2, None)  # SUM skips the text results


def test_text_to_number_and_number_to_text():
    got, c = run({"S": {
        "A1": 1e18,
        "B1": ('="3"+1', 4), "B2": ('=" 3 "*2', 6), "B3": ('="abc"+1', "#VALUE!"), "B4": ('="1,000"+1', "#VALUE!"),
        "B5": ('="x"&1/3', "x0.333333333333333"), "B6": ('="x"&0.1+0.2', "x0.3"), "B7": ('="x"&TRUE', "xTRUE"),
        "B8": ('=A1&""', "1E+18"), "B9": ('=IF("TRUE", 1, 2)', 1), "B10": ('=IF("x", 1, 2)', "#VALUE!"),
    }})
    assert c.unresolved == {}
    assert [got[("S", r, 2)] for r in range(1, 11)] == [
        ("n", 4.0, None), ("n", 6.0, None), ("e", None, "#VALUE!"), ("e", None, "#VALUE!"),
        ("s", None, "x0.333333333333333"), ("s", None, "x0.3"), ("s", None, "xTRUE"), ("s", None, "1E+18"),
        ("n", 1.0, None), ("e", None, "#VALUE!")]


def test_comparisons_across_types_follow_excel_order():
    got, c = run({"S": {
        "A1": 5, "A2": "5", "B1": ('=1<"a"', True), "B2": ('="a"<TRUE', True), "B3": ("=A1=A2", False),
        "B4": ('=Z9=""', True), "B5": ("=Z9=0", True), "B6": ('=IF(A1>0,"x",1)>2', True),  # text > any number
    }})
    assert c.unresolved == {}
    assert [got[("S", r, 2)] for r in range(1, 7)] == [
        ("b", 1.0, None), ("b", 1.0, None), ("b", 0.0, None), ("b", 1.0, None), ("b", 1.0, None), ("b", 1.0, None)]


def test_mixed_vlookup_column_and_verdicts_on_mixed_cells():
    book = book_of({"R": {"A1": "a", "B1": 1, "A2": "b", "B2": "two"}, "S": {
        "A1": ("=VLOOKUP(\"a\", R!A1:B2, 2, FALSE)", 1), "A2": ("=VLOOKUP(\"b\", R!A1:B2, 2, FALSE)", "two"),
        "A3": ('=IF(A1=1, "", 5)', "")}})
    assert {v.status for v in verify(book, compile_book(book))} == {"PASS"}
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_types.py -v`
Expected: FAIL. Mixed IF branches are UNRESOLVED ("mixes txt and num"), `"3"+1` is UNRESOLVED ("text used as a number"), `"x"&1/3` is UNRESOLVED ("number used as text").

- [ ] **Step 3: Replace `src/unexcel/expr.py`**

```python
"""Typed SQL expressions, Excel-style type coercion and Excel error codes."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from unexcel.model import Key, Unsupported

CELLS = "__unexcel_cells__"  # every cell known before the current batch; replaced when SQL is rendered
NO_ERR = "NULL"  # the err of an expression that cannot fail
Type = Literal["num", "txt", "bool", "blank", "err", "any"]


@dataclass(frozen=True)
class Expr:
    sql: str  # the value; for "any" its number part (numbers, TRUE/FALSE as 1/0); meaningless when err is set
    type: Type
    err: str = NO_ERR  # SQL giving the Excel error code ('#DIV/0!') or NULL
    txt: str = "NULL"  # "any" only: the text part
    kind: str = "NULL"  # "any" only: SQL giving 'n', 's' or 'b' - the value's type, known only when the SQL runs


ZERO = Expr("CAST(0 AS DOUBLE)", "num")
BLANK = Expr("NULL", "blank")  # an empty cell; always coerced before use, never rendered raw


def lit_num(v: float) -> Expr:
    return Expr(f"CAST({v!r} AS DOUBLE)", "num")  # bare literals would be DECIMAL in DuckDB


def lit_str(v: str) -> Expr:
    return Expr("'" + v.replace("'", "''") + "'", "txt")


def error(code: str) -> Expr:
    """A value that is always this Excel error."""
    return Expr("NULL", "err", lit_str(code).sql)


def errs(*parts: str) -> str:
    """Excel reports the first error it meets, left to right."""
    live = [p for p in parts if p != NO_ERR]
    if not live:
        return NO_ERR
    return live[0] if len(live) == 1 else f"COALESCE({', '.join(live)})"


def any_of(kind: str, num_sql: str, txt_sql: str, err: str = NO_ERR) -> Expr:
    return Expr(num_sql, "any", err, txt_sql, kind)


def parse_num(t: str) -> str:
    """Text -> number the way Excel's arithmetic reads it (' 3 ', '-1.5e3'), else NULL. 'inf', '1,000' -> NULL."""
    return (f"(CASE WHEN regexp_matches(trim({t}), '^[+-]?([0-9]+[.]?[0-9]*|[.][0-9]+)([eE][+-]?[0-9]+)?$') "
            f"THEN CAST(trim({t}) AS DOUBLE) END)")


def num_text(x: str) -> str:
    """Number -> text the way Excel writes it in a formula: up to 15 significant digits (0.1+0.2 -> '0.3')."""
    return f"(CASE WHEN {x} = 0 THEN '0' ELSE upper(printf('%.15g', {x})) END)"


def _value_error(bad: str) -> str:
    return f"(CASE WHEN {bad} THEN '#VALUE!' END)"


def num(e: Expr) -> Expr:
    if e.type == "num":
        return e
    if e.type == "blank":
        return ZERO
    if e.type == "bool":
        return Expr(f"CAST({e.sql} AS DOUBLE)", "num", e.err)
    if e.type == "err":
        return Expr("CAST(NULL AS DOUBLE)", "num", e.err)
    if e.type == "txt":
        n = parse_num(e.sql)
        return Expr(n, "num", errs(e.err, _value_error(f"{n} IS NULL")))
    n = parse_num(e.txt)
    return Expr(f"(CASE WHEN {e.kind} = 's' THEN {n} ELSE {e.sql} END)", "num",
                errs(e.err, _value_error(f"{e.kind} = 's' AND {n} IS NULL")))


def _bool_text(b: str) -> str:
    return f"(CASE WHEN {b} THEN 'TRUE' ELSE 'FALSE' END)"


def txt(e: Expr) -> Expr:
    if e.type == "txt":
        return e
    if e.type == "blank":
        return Expr("''", "txt")
    if e.type == "err":
        return Expr("CAST(NULL AS VARCHAR)", "txt", e.err)
    if e.type == "num":
        return Expr(num_text(e.sql), "txt", e.err)
    if e.type == "bool":
        return Expr(_bool_text(e.sql), "txt", e.err)
    return Expr(f"(CASE {e.kind} WHEN 's' THEN {e.txt} WHEN 'b' THEN {_bool_text(f'{e.sql} <> 0')} "
                f"ELSE {num_text(e.sql)} END)", "txt", e.err)


def _text_bool(t: str) -> str:
    return f"(CASE upper({t}) WHEN 'TRUE' THEN TRUE WHEN 'FALSE' THEN FALSE END)"


def as_bool(e: Expr) -> Expr:
    if e.type == "bool":
        return e
    if e.type == "blank":
        return Expr("FALSE", "bool")
    if e.type == "num":
        return Expr(f"({e.sql} <> 0)", "bool", e.err)
    if e.type == "err":
        return Expr("CAST(NULL AS BOOLEAN)", "bool", e.err)
    if e.type == "txt":
        b = _text_bool(e.sql)
        return Expr(b, "bool", errs(e.err, _value_error(f"{b} IS NULL")))
    b = _text_bool(e.txt)
    return Expr(f"(CASE WHEN {e.kind} = 's' THEN {b} ELSE {e.sql} <> 0 END)", "bool",
                errs(e.err, _value_error(f"{e.kind} = 's' AND {b} IS NULL")))


def as_any(e: Expr) -> Expr:
    """Any value as the run-time-typed form."""
    if e.type == "any":
        return e
    if e.type in ("blank", "num"):
        e = num(e)
        return any_of("'n'", e.sql, "CAST(NULL AS VARCHAR)", e.err)
    if e.type == "txt":
        return any_of("'s'", "CAST(NULL AS DOUBLE)", e.sql, e.err)
    if e.type == "bool":
        return any_of("'b'", f"CAST({e.sql} AS DOUBLE)", "CAST(NULL AS VARCHAR)", e.err)
    return any_of("'n'", "CAST(NULL AS DOUBLE)", "CAST(NULL AS VARCHAR)", e.err)  # err


COERCE = {"num": num, "txt": txt, "bool": as_bool, "any": as_any}
FLEX = ("blank", "err")  # types that take on the other side's type


def unify(a: Expr, b: Expr) -> tuple[Expr, Expr]:
    """Give two results one type (IF branches, IFERROR). Different types become 'any', as in Excel."""
    return tuple(unify_all([a, b]))


def unify_all(es: list[Expr]) -> list[Expr]:
    fixed = {e.type for e in es} - set(FLEX)
    if not fixed:
        return [num(e) for e in es]
    target = fixed.pop() if len(fixed) == 1 else "any"
    return [COERCE[target](e) for e in es]


def pick(cond: str, yes: Expr, no: Expr, err: str = NO_ERR) -> Expr:
    """CASE WHEN cond THEN yes ELSE no, for two results of one type (see unify), keeping every part."""
    def case(a: str, b: str) -> str:
        return a if a == b else f"(CASE WHEN {cond} THEN {a} ELSE {b} END)"
    return Expr(case(yes.sql, no.sql), yes.type, err, case(yes.txt, no.txt), case(yes.kind, no.kind))


RANK = {"num": 0, "txt": 1, "bool": 2}  # Excel orders numbers < text < TRUE/FALSE across types


def _rank(kind: str) -> str:
    return f"(CASE {kind} WHEN 'n' THEN 0 WHEN 's' THEN 1 ELSE 2 END)"


def _blank_like(other: Expr) -> Expr:
    """An empty cell compared with something acts as 0, "" or FALSE - whatever the other side is."""
    if other.type == "txt":
        return Expr("''", "txt")
    if other.type == "bool":
        return Expr("FALSE", "bool")
    if other.type == "any":
        return any_of(other.kind, "CAST(0 AS DOUBLE)", "''")
    return ZERO


def compare(a: Expr, b: Expr, op: str) -> Expr:
    """Excel comparison: text ignores case; different types compare by type (number < text < logical)."""
    if a.type == "blank" and b.type == "blank":
        a, b = ZERO, ZERO
    elif a.type == "blank":
        a = _blank_like(b)
    elif b.type == "blank":
        b = _blank_like(a)
    if a.type == "err":
        a = num(a)
    if b.type == "err":
        b = num(b)
    err = errs(a.err, b.err)
    if a.type == b.type and a.type != "any":
        if a.type == "txt":
            return Expr(f"(LOWER({a.sql}) {op} LOWER({b.sql}))", "bool", err)
        return Expr(f"({a.sql} {op} {b.sql})", "bool", err)
    if a.type != "any" and b.type != "any":
        return Expr(f"({RANK[a.type]} {op} {RANK[b.type]})", "bool", err)
    a, b = as_any(a), as_any(b)
    ra, rb = _rank(a.kind), _rank(b.kind)
    return Expr(f"(CASE WHEN {ra} <> {rb} THEN {ra} {op} {rb} WHEN {ra} = 1 THEN LOWER({a.txt}) {op} LOWER({b.txt}) "
                f"ELSE {a.sql} {op} {b.sql} END)", "bool", err)


def finalize(e: Expr) -> Expr:
    return ZERO if e.type == "blank" else e  # =A1 on an empty cell shows 0 in Excel


def where_cell(key: Key) -> str:
    sheet, row, col = key
    return f"sheet = {lit_str(sheet).sql} AND row_num = {row} AND col_num = {col}"


def where_area(area, alias: str = "") -> str:
    p = f"{alias}." if alias else ""
    return (f"{p}sheet = {lit_str(area.sheet).sql} AND {p}row_num BETWEEN {area.r1} AND {area.r2} "
            f"AND {p}col_num BETWEEN {area.c1} AND {area.c2}")
```

- [ ] **Step 4: Replace `src/unexcel/compiler.py`**

```python
"""Formula AST -> SQL. Cell mode, run in batches.

Every cell known so far (inputs, then the results of earlier batches) lives in one table with the
columns of INPUT_COLUMNS: kind is n/s/b/e, num holds numbers and TRUE/FALSE as 1/0, txt holds text
and Excel error codes. A batch is up to BATCH consecutive formula cells in dependency order. It
becomes one SQL statement: the cells it computes are materialized CTEs (v = value, e = error code
or NULL), appended to the table it read from. One statement per workbook broke DuckDB at ~1k
formula cells; batches keep every statement small.
"""
from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from unexcel.expr import (BLANK, CELLS, NO_ERR, Expr, any_of, compare, error, errs, finalize, lit_num, lit_str, num,
                          txt, where_cell)
from unexcel.functions import FUNCS
from unexcel.graph import formula_order
from unexcel.model import Book, Cell, Key, Unsupported, a1
from unexcel.parser import Binary, Bool, Call, ErrLit, Num, Postfix, RefNode, Str, Unary, parse
from unexcel.refs import Area, resolve

INPUT_COLUMNS = "sheet VARCHAR, row_num INTEGER, col_num INTEGER, kind VARCHAR, num DOUBLE, txt VARCHAR"
SEED_HEADER = ["sheet", "row_num", "col_num", "kind", "num", "txt"]
READ_SEED = ("read_csv(?, header = true, quote = '\"', escape = '\"', columns = {'sheet': 'VARCHAR', "
             "'row_num': 'INTEGER', 'col_num': 'INTEGER', 'kind': 'VARCHAR', 'num': 'DOUBLE', 'txt': 'VARCHAR'})")
VOLATILE = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT"}
ARITH = {"+", "-", "*"}
KIND = {"num": "n", "txt": "s", "bool": "b", "err": "e", "any": "a"}  # a: decided when the SQL runs
TYPE = {"n": "num", "s": "txt", "b": "bool"}
NO_VALUE = "no saved value in the file - open and save it in Excel"
DEEP = "formula is nested too deeply to translate"
TOO_BIG = "formula's SQL would be too large to run"
# IFERROR/division repeat their arguments in the value and the error SQL, so SQL can double per
# nesting level; checking every node keeps memory near this bound (200k-cell workbooks stay < 10 KB/cell)
MAX_SQL = 1_000_000
# ponytail: fixed batch size. Measured on DuckDB 1.5.6: 200k flat cells 71 s, a 100k-row running total 50 s.
# Table mode (Plan 2c) is the real speed-up; change this only with a new measurement.
BATCH = 250


@dataclass(frozen=True)
class Step:
    key: Key
    expr: Expr
    cte: str  # its CTE name inside its batch


@dataclass
class Compiled:
    batches: list[list[Step]]
    unresolved: dict[Key, str]
    findings: list[str]

    @property
    def steps(self) -> list[Step]:
        return [s for batch in self.batches for s in batch]


class NewBatch(Exception):
    """A range reads a cell of the open batch. Ranges read the table, so the batch must close first."""


class Compiler:
    def __init__(self, book: Book):
        self.book = book
        self.done: dict[Key, tuple[str, bool]] = {}  # formula cell in a closed batch -> (type, may be an error)
        self.open: dict[Key, tuple[str, str, bool]] = {}  # formula cell in the open batch -> (cte, type, may err)
        self.findings: list[str] = []

    def compile_cell(self, cell: Cell) -> Expr:
        if cell.array:
            raise Unsupported("array formula")
        return finalize(self.expr(parse(cell.formula), cell.sheet))

    def expr(self, node, sheet: str) -> Expr:
        e = self._expr(node, sheet)
        if sum(len(part) for part in (e.sql, e.err, e.txt, e.kind)) > MAX_SQL:
            raise Unsupported(TOO_BIG)
        return e

    def _expr(self, node, sheet: str) -> Expr:
        if node is None:
            return BLANK
        if isinstance(node, Num):
            return lit_num(node.value)
        if isinstance(node, Str):
            return lit_str(node.value)
        if isinstance(node, Bool):
            return Expr("TRUE" if node.value else "FALSE", "bool")
        if isinstance(node, ErrLit):
            return error(node.value)
        if isinstance(node, RefNode):
            area = resolve(node.text, sheet, self.book)
            if not area.is_single:
                raise Unsupported(f"range {node.text} used as a single value")
            return self.cell_value(area.key)
        if isinstance(node, Unary):
            x = num(self.expr(node.operand, sheet))
            return Expr(f"(-{x.sql})", "num", x.err)
        if isinstance(node, Postfix):
            x = num(self.expr(node.operand, sheet))
            return Expr(f"({x.sql} / 100)", "num", x.err)
        if isinstance(node, Binary):
            return self.binary(node, sheet)
        if isinstance(node, Call):
            if node.name in VOLATILE:
                raise Unsupported(f"volatile function {node.name}")
            fn = FUNCS.get(node.name)
            if fn is None:
                raise Unsupported(f"function {node.name} not supported yet")
            return fn(self, node.args, sheet)
        raise Unsupported(f"cannot translate {node!r}")

    def binary(self, node: Binary, sheet: str) -> Expr:
        a, b, op = self.expr(node.left, sheet), self.expr(node.right, sheet), node.op
        if op in ARITH or op in ("/", "^"):
            x, y = num(a), num(b)
            if op in ARITH:
                return Expr(f"({x.sql} {op} {y.sql})", "num", errs(x.err, y.err))
            if op == "/":  # DuckDB gives inf, Excel #DIV/0!
                return Expr(f"({x.sql} / NULLIF({y.sql}, 0))", "num",
                            errs(x.err, y.err, f"(CASE WHEN {y.sql} = 0 THEN '#DIV/0!' END)"))
            # DuckDB POWER returns inf or nan where Excel returns an error (verified on 1.5.6)
            bad = (f"(CASE WHEN {x.sql} = 0 AND {y.sql} < 0 THEN '#DIV/0!' "
                   f"WHEN {x.sql} = 0 AND {y.sql} = 0 THEN '#NUM!' "
                   f"WHEN {x.sql} < 0 AND {y.sql} <> FLOOR({y.sql}) THEN '#NUM!' END)")
            return Expr(f"POWER({x.sql}, {y.sql})", "num", errs(x.err, y.err, bad))
        if op == "&":
            x, y = txt(a), txt(b)
            return Expr(f"({x.sql} || {y.sql})", "txt", errs(x.err, y.err))
        return compare(a, b, op)

    def cell_value(self, key: Key) -> Expr:
        cell = self.book.cells.get(key)
        if cell is None or (cell.formula is None and cell.kind == "blank"):
            return BLANK
        if key in self.open:
            cte, typ, may_err = self.open[key]
            err = f"(SELECT e FROM {cte})" if may_err else NO_ERR
            if typ == "any":
                return any_of(f"(SELECT k FROM {cte})", f"(SELECT v FROM {cte})", f"(SELECT t FROM {cte})", err)
            return Expr(f"(SELECT v FROM {cte})", typ, err)
        if cell.formula is None:
            if cell.kind == "e":
                return error(cell.value)
            typ, may_err = TYPE[cell.kind], False
        elif key in self.done:
            typ, may_err = self.done[key]
        else:
            raise Unsupported(f"depends on unresolved cell {a1(key)}")
        where = f"FROM {CELLS} WHERE {where_cell(key)}"
        err = f"(SELECT CASE WHEN kind = 'e' THEN txt END {where})" if may_err else NO_ERR
        if typ == "num":
            return Expr(f"(SELECT num {where})", "num", err)
        if typ == "txt":
            return Expr(f"(SELECT COALESCE(txt, '') {where})", "txt", err)  # the seed CSV reads '' back as NULL
        if typ == "bool":
            return Expr(f"((SELECT num {where}) <> 0)", "bool", err)
        if typ == "any":
            return any_of(f"(SELECT kind {where})", f"(SELECT num {where})", f"(SELECT COALESCE(txt, '') {where})", err)
        return Expr("NULL", "err", err)

    def ready(self, area: Area) -> list[Key]:
        """The non-empty cells of a range that SQL will read from the table. All must be computed already."""
        keys = area.present(self.book)
        for key in keys:
            if self.book.cells[key].formula is None:
                continue
            if key in self.open:
                raise NewBatch
            if key not in self.done:
                raise Unsupported(f"depends on unresolved cell {a1(key)}")
        return keys

    def kind_of(self, key: Key) -> str:
        """Value kind (n/s/b/e) a cell has in the table when it isn't an error. Call ready() first."""
        cell = self.book.cells[key]
        return KIND[self.done[key][0]] if cell.formula is not None else cell.kind

    def may_err(self, key: Key) -> bool:
        """Whether the table can hold an error for this cell. Call ready() first."""
        cell = self.book.cells[key]
        return self.done[key][1] if cell.formula is not None else cell.kind == "e"


def _close(comp: Compiler, batches: list[list[Step]], current: list[Step]) -> None:
    """Seal the open batch: later cells read its results from the table."""
    if current:
        batches.append(list(current))
        comp.done.update((s.key, (s.expr.type, s.expr.err != NO_ERR)) for s in current)
        comp.open.clear()
        current.clear()


def compile_book(book: Book) -> Compiled:
    order, unresolved = formula_order(book)
    comp = Compiler(book)
    batches: list[list[Step]] = []
    current: list[Step] = []
    for key in order:
        cell = book.cells[key]
        if cell.kind == "blank":
            unresolved[key] = NO_VALUE
            continue
        if len(current) >= BATCH:
            _close(comp, batches, current)
        try:
            try:
                e = comp.compile_cell(cell)
            except NewBatch:
                _close(comp, batches, current)
                e = comp.compile_cell(cell)  # the open batch is empty now, so this can't raise NewBatch
        except Unsupported as err:
            unresolved[key] = str(err)
            continue
        except RecursionError:  # Python's stack, not DuckDB's: DuckDB took 900 nesting levels in tests
            unresolved[key] = DEEP
            continue
        step = Step(key, e, f"c{len(current) + 1}")
        comp.open[key] = (step.cte, e.type, e.err != NO_ERR)
        current.append(step)
    _close(comp, batches, current)
    return Compiled(batches, unresolved, comp.findings)


def render_batch(steps: list[Step], prev: str) -> str:
    """One batch as a SELECT: every row of `prev` plus one row per cell this batch computes."""
    ctes = ",\n".join(f"{s.cte} AS MATERIALIZED (SELECT {s.expr.sql} AS v, CAST({s.expr.err} AS VARCHAR) AS e, "
                      f"CAST({s.expr.txt} AS VARCHAR) AS t, CAST({s.expr.kind} AS VARCHAR) AS k)" for s in steps)
    rows = []
    for s in steps:
        sheet, row, col = s.key
        typ = s.expr.type
        v, e, t, k = (f"(SELECT {c} FROM {s.cte})" for c in "vetk")
        if typ == "any":  # the kind is only known when the SQL runs
            kind = f"CASE WHEN {e} IS NULL THEN {k} ELSE 'e' END"
            num_sql = f"CASE WHEN {e} IS NULL AND {k} <> 's' THEN CAST({v} AS DOUBLE) END"
            txt_sql = f"CASE WHEN {e} IS NOT NULL THEN {e} WHEN {k} = 's' THEN {t} END"
        else:
            kind = f"CASE WHEN {e} IS NULL THEN '{KIND[typ]}' ELSE 'e' END"
            num_sql = ("CAST(NULL AS DOUBLE)" if typ in ("txt", "err")
                       else f"CASE WHEN {e} IS NULL THEN CAST({v} AS DOUBLE) END")
            txt_sql = f"COALESCE({e}, CAST({v} AS VARCHAR))" if typ == "txt" else e
        rows.append(f"SELECT {lit_str(sheet).sql}, {row}, {col}, {kind}, {num_sql}, {txt_sql}")
    sql = f"WITH\n{ctes}\nSELECT * FROM {CELLS}\nUNION ALL\n" + "\nUNION ALL\n".join(rows)
    return sql.replace(CELLS, prev)


def input_rows(book: Book) -> list[tuple]:
    """Every non-formula cell as a row of the cells table (see INPUT_COLUMNS)."""
    rows = []
    for (sheet, row, col), cell in sorted(book.cells.items()):
        if cell.formula is not None or cell.kind == "blank":
            continue
        if cell.kind == "b":
            num_v, txt_v = (1.0 if cell.value else 0.0), None
        elif cell.kind == "n":
            num_v, txt_v = cell.value, None
        else:
            num_v, txt_v = None, cell.value
        rows.append((sheet, row, col, cell.kind, num_v, txt_v))
    return rows


def write_seed(book: Book, path: Path) -> None:
    """The input cells as CSV: the dbt seed, and what the verifier loads (0.1 s for 200k rows)."""
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(SEED_HEADER)
        for row in input_rows(book):
            w.writerow(["" if x is None else x for x in row])
```

- [ ] **Step 5: Replace `src/unexcel/functions.py`** (IF and IFERROR keep every part of an any-typed result; VLOOKUP returns `"any"` for a mixed column)

```python
"""Excel function -> SQL. Each entry: fn(compiler, args, sheet) -> Expr.

Ranges are read from the cells table (inputs plus the formula cells of earlier batches).
Compiler.ready() checks every formula cell in a range was computed in an earlier batch.
Every function returns its Excel error in Expr.err; see expr.errs for the order rule.
"""
from __future__ import annotations

from unexcel.expr import CELLS, NO_ERR, ZERO, Expr, any_of, as_bool, errs, lit_str, num, pick, unify, where_area
from unexcel.model import Unsupported, a1
from unexcel.parser import Bool, Num, RefNode, Str
from unexcel.refs import Area, resolve

FUNCS = {}


def register(name):
    def deco(fn):
        FUNCS[name] = fn
        return fn

    return deco


def arity(name, args, lo, hi):
    if not lo <= len(args) <= hi:
        raise Unsupported(f"{name} needs {lo}-{hi} arguments, got {len(args)}")


def _looks_numeric(text: str) -> bool:
    try:
        float(text.replace(",", "").strip())
        return True
    except ValueError:
        return False


def note_text_numbers(c, keys):
    """Excel skips text in SUM-like ranges, even text that looks like a number. Tell the user."""
    for key in keys:
        cell = c.book.cells[key]
        if cell.formula is None and cell.kind == "s" and _looks_numeric(cell.value):
            c.findings.append(f"{a1(key)}: number stored as text ({cell.value!r}) is skipped by Excel")


def first_error(c, area: Area, keys) -> str:
    """SQL for the first error in a range, row by row, or NO_ERR when no cell in it can hold one."""
    if not any(c.may_err(k) for k in keys):
        return NO_ERR
    return f"(SELECT txt FROM {CELLS} WHERE {where_area(area)} AND kind = 'e' ORDER BY row_num, col_num LIMIT 1)"


def range_sum(c, area: Area) -> Expr:
    """SUM over a range: text, TRUE/FALSE and blanks are skipped; the first error wins."""
    keys = c.ready(area)
    note_text_numbers(c, keys)
    return Expr(f"(SELECT COALESCE(SUM(num), CAST(0 AS DOUBLE)) FROM {CELLS} WHERE {where_area(area)} AND kind = 'n')",
                "num", first_error(c, area, keys))


@register("SUM")
def sum_(c, args, sheet):
    parts = [range_sum(c, resolve(arg.text, sheet, c.book)) if isinstance(arg, RefNode) else num(c.expr(arg, sheet))
             for arg in args]
    if not parts:
        return ZERO
    return Expr("(" + " + ".join(p.sql for p in parts) + ")", "num", errs(*(p.err for p in parts)))


@register("IF")
def if_(c, args, sheet):
    arity("IF", args, 2, 3)
    cond = as_bool(c.expr(args[0], sheet))
    yes = c.expr(args[1], sheet)
    no = c.expr(args[2], sheet) if len(args) == 3 else Expr("FALSE", "bool")
    yes, no = unify(yes, no)
    taken = NO_ERR if yes.err == no.err == NO_ERR else f"(CASE WHEN {cond.sql} THEN {yes.err} ELSE {no.err} END)"
    return pick(cond.sql, yes, no, errs(cond.err, taken))


@register("ROUND")
def round_(c, args, sheet):
    arity("ROUND", args, 2, 2)
    x, digits = num(c.expr(args[0], sheet)), num(c.expr(args[1], sheet))
    # DuckDB ROUND on DOUBLE rounds half away from zero like Excel (2.675 -> 2.68, verified on 1.5.6)
    return Expr(f"ROUND({x.sql}, CAST(TRUNC({digits.sql}) AS INTEGER))", "num", errs(x.err, digits.err))


def _logic(name, joiner):
    def fn(c, args, sheet):
        if not args:
            raise Unsupported(f"{name} needs arguments")
        parts = [as_bool(c.expr(a, sheet)) for a in args]
        return Expr("(" + f" {joiner} ".join(p.sql for p in parts) + ")", "bool", errs(*(p.err for p in parts)))

    return fn


FUNCS["AND"] = _logic("AND", "AND")
FUNCS["OR"] = _logic("OR", "OR")


@register("IFERROR")
def iferror(c, args, sheet):
    arity("IFERROR", args, 2, 2)
    value, fallback = unify(c.expr(args[0], sheet), c.expr(args[1], sheet))
    if value.err == NO_ERR:
        return value
    ok = f"(({value.err}) IS NULL)"
    err = NO_ERR if fallback.err == NO_ERR else f"(CASE WHEN {ok} THEN NULL ELSE {fallback.err} END)"
    return pick(ok, value, fallback, err)


def _column(c, node, sheet, fname) -> tuple[Area, list]:
    if not isinstance(node, RefNode):
        raise Unsupported(f"{fname} needs a cell range")
    area = resolve(node.text, sheet, c.book)
    return area, c.ready(area)


def _match(alias: str, crit: Expr) -> str:
    if crit.type == "txt":  # Excel matching ignores case
        return f"{alias}.kind = 's' AND LOWER({alias}.txt) = LOWER({crit.sql})"
    if crit.type == "num":
        return f"{alias}.kind = 'n' AND {alias}.num = {crit.sql}"
    raise Unsupported(f"criteria of type {crit.type}")


@register("SUMIFS")
def sumifs(c, args, sheet):
    if len(args) < 3 or len(args) % 2 == 0:
        raise Unsupported("SUMIFS needs a sum range and criteria pairs")
    total, keys = _column(c, args[0], sheet, "SUMIFS")
    if total.c1 != total.c2:
        raise Unsupported("SUMIFS sum range wider than one column")
    keys = list(keys)
    joins, conds = [], []
    for i in range(1, len(args), 2):
        area, area_keys = _column(c, args[i], sheet, "SUMIFS")
        keys += area_keys
        if area.c1 != area.c2 or area.r2 - area.r1 != total.r2 - total.r1:
            raise Unsupported("SUMIFS ranges must be single columns of the same height")
        node = args[i + 1]
        if isinstance(node, Str) and node.value and (node.value[0] in "<>=" or any(ch in node.value for ch in "*?~")):
            raise Unsupported("SUMIFS comparison/wildcard criteria not supported yet")
        crit = c.expr(node, sheet)
        if crit.err != NO_ERR:
            raise Unsupported("SUMIFS criteria that can be an error not supported yet")
        k = f"k{len(joins) + 1}"
        joins.append(f"JOIN {CELLS} {k} ON {k}.sheet = {lit_str(area.sheet).sql} AND {k}.col_num = {area.c1} "
                     f"AND {k}.row_num = s.row_num + {area.r1 - total.r1}")
        conds.append(_match(k, crit))
    if any(c.may_err(k) for k in keys):
        raise Unsupported("SUMIFS over cells that hold errors not supported yet")
    note_text_numbers(c, [k for k in keys if k[2] == total.c1 and k[0] == total.sheet])
    where = " AND ".join([where_area(total, "s"), "s.kind = 'n'", *conds])
    return Expr(f"(SELECT COALESCE(SUM(s.num), CAST(0 AS DOUBLE)) FROM {CELLS} s {' '.join(joins)} WHERE {where})",
                "num")


@register("VLOOKUP")
def vlookup(c, args, sheet):
    if len(args) == 3:
        raise Unsupported("VLOOKUP approximate match not supported yet")
    arity("VLOOKUP", args, 4, 4)
    exact = args[3]
    if not (exact is None or (isinstance(exact, Bool) and not exact.value)
            or (isinstance(exact, Num) and exact.value == 0)):
        raise Unsupported("VLOOKUP approximate match not supported yet")
    if not isinstance(args[2], Num):
        raise Unsupported("VLOOKUP column index must be a number")
    table, keys = _column(c, args[1], sheet, "VLOOKUP")
    col = table.c1 + int(args[2].value) - 1
    if not table.c1 <= col <= table.c2:
        raise Unsupported("VLOOKUP column index outside the table")
    kinds = [c.kind_of(k) for k in keys if k[2] == col]
    values = [k for k in kinds if k != "e"]  # an error in the result column is returned as that error
    if len(values) > 1 and values[0] == "s" and set(values[1:]) == {"n"}:
        values = values[1:]  # text header above a numeric column
    if set(values) <= {"n"}:
        column, typ, empty = "num", "num", "CAST(0 AS DOUBLE)"
    elif set(values) == {"s"}:
        column, typ, empty = "txt", "txt", "''"
    else:
        column, typ, empty = None, "any", None  # mixed column: the type is known only when the SQL runs
    lookup = c.expr(args[0], sheet)
    first_col = Area(table.sheet, table.r1, table.c1, table.r2, table.c1)
    src = (f"FROM {CELLS} k LEFT JOIN {CELLS} r ON r.sheet = k.sheet AND r.row_num = k.row_num "
           f"AND r.col_num = {col} WHERE {where_area(first_col, 'k')} AND {_match('k', lookup)} "
           f"ORDER BY k.row_num LIMIT 1")
    found = f"(SELECT CASE WHEN r.kind = 'e' THEN r.txt ELSE '' END {src})"  # NULL = no row matched
    err = f"(CASE WHEN {found} IS NULL THEN '#N/A' WHEN {found} = '' THEN NULL ELSE {found} END)"
    if typ == "any":
        return any_of(f"(SELECT COALESCE(r.kind, 'n') {src})", f"(SELECT COALESCE(r.num, 0) {src})",
                      f"(SELECT COALESCE(r.txt, '') {src})", errs(lookup.err, err))
    return Expr(f"(SELECT COALESCE(r.{column}, {empty}) {src})", typ, errs(lookup.err, err))
```

- [ ] **Step 6: Whole columns mean every row, in `src/unexcel/refs.py`**

In `src/unexcel/refs.py` replace
```python
from unexcel.model import Book, Key, Unsupported
```
with
```python
from unexcel.model import Book, Key, Unsupported

MAX_ROW = 1_048_576
```
In `src/unexcel/refs.py` replace
```python
    return Area(sheet, r1 or 1, c1, r2 or book.max_row.get(sheet, 1), c2)
```
with
```python
    return Area(sheet, r1 or 1, c1, r2 or MAX_ROW, c2)  # C:C is every row, as in Excel; Area.present skips empties
```

- [ ] **Step 7: Update the three Plan 1 expectations this changes**

In `tests/test_compiler.py` replace
```python
    assert bad[("S", 2, 1)] == "text used as a number"
```
with
```python
    assert ("S", 2, 1) not in bad  # text in arithmetic is Excel's #VALUE!, computed like any other error
```
In `tests/test_functions.py` replace
```python
def test_if_with_different_branch_types_is_unresolved():
    _, c = run({"S": {"A1": 1, "B1": ('=IF(A1>0, "yes", 0)', "yes"), "B2": ("=IF(A1)", 0)}})
    assert c.unresolved[("S", 1, 2)] == "mixes txt and num"
```
with
```python
def test_if_with_wrong_argument_count_is_unresolved():
    _, c = run({"S": {"A1": 1, "B1": ('=IF(A1>0, "yes", 0)', "yes"), "B2": ("=IF(A1)", 0)}})
    assert ("S", 1, 2) not in c.unresolved  # mixed branch types are fine since Plan 2b
```
In `tests/test_refs.py` replace
```python
    assert resolve("C:C", "S", BOOK) == Area("S", 1, 3, 10, 3)
```
with
```python
    assert resolve("C:C", "S", BOOK) == Area("S", 1, 3, 1_048_576, 3)  # every row, as in Excel
```

- [ ] **Step 8: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 9: Commit**

```bash
git add src/unexcel/expr.py src/unexcel/compiler.py src/unexcel/functions.py src/unexcel/refs.py tests
git commit -m "feat: run-time types and Excel's text/number/logical conversions" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: The criteria engine: COUNTIF, COUNTIFS, SUMIF, SUMIFS, AVERAGEIF, AVERAGEIFS, MINIFS, MAXIFS

**Files:**
- Create: `src/unexcel/fn/__init__.py`, `src/unexcel/fn/criteria.py`
- Modify: `src/unexcel/compiler.py` (import the library), `src/unexcel/functions.py` (SUMIFS moves out), `tests/test_errors.py`, `tests/test_lookup.py`
- Test: `tests/test_criteria.py`

**Interfaces:**
- Consumes: `Compiler.ready/kind_of/may_err/expr`, `expr.parse_num/txt/lit_str/where_area/errs`, `functions.arity/note_text_numbers/register`.
- Produces:
  - Package `unexcel.fn`: importing it registers the library. `compiler.py` does `import unexcel.fn`. Later tasks add one module name to `fn/__init__.py`.
  - `criteria.match_sql(alias, crit, q) -> (match_sql, blank_sql, table_sql)`: Excel's criteria rules for one cell; `table_sql` is the one-row table `q` the caller puts in its FROM clause (a text criterion is parsed there once).
  - `criteria.like_pattern(sql) -> sql`: Excel wildcards to a `LIKE ... ESCAPE '\'` pattern (used by Task 4).
- Rules: a number criterion matches equal numbers and text that reads as that number; text may start with `= <> < > <= >=`; `*` `?` are wildcards and `~` escapes them; text compares ignore case; empty cells match `""`, `"="` and `"<>x"`; a matched error in the sum range is the result.

- [ ] **Step 1: Write the failing tests `tests/test_criteria.py`**

```python
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute

# A1:A9 - 5, "5", apple, Banana, TRUE, (blank), "", a*b, #N/A ; B: amounts ; C: regions
DATA = {
    "A1": 5, "A2": "5", "A3": "apple", "A4": "Banana", "A5": True, "A7": "", "A8": "a*b", "A9": "#N/A",
    "B1": 10, "B2": 20, "B3": 30, "B4": 40, "B5": 50, "B6": 60, "B7": 70, "B8": 80, "B9": 90,
    "C1": "W", "C2": "E", "C3": "W", "C4": "E", "C5": "W", "C6": "E", "C7": "W", "C8": "E", "C9": "W",
    "D1": ">", "D2": 4,
}


def run(formulas):
    book = book_of({"S": DATA, "F": formulas})
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def N(v):
    return ("n", float(v), None)


def test_countif_criteria_forms():
    cases = {'"5"': 2, '">4"': 1, '"<>5"': 7, '"a*"': 2, '"a~*b"': 1, '">b"': 1, '"TRUE"': 1, '""': 2, '"="': 2,
             '"<>"': 7, '"?pple"': 1, '"%"': 0, 'S!D1&S!D2': 1, '"<=5"': 1, "5": 2, "TRUE": 1}
    formulas = {f"A{i}": (f"=COUNTIF(S!A1:A9, {crit})", n) for i, (crit, n) in enumerate(cases.items(), 1)}
    got, c = run(formulas)
    assert c.unresolved == {}
    assert [got[("F", i, 1)] for i in range(1, len(cases) + 1)] == [N(n) for n in cases.values()]


def test_countifs_counts_blank_rows_too():
    got, c = run({"A1": ('=COUNTIFS(S!A1:A9, "", S!C1:C9, "E")', 1),  # A6 blank & E
                  "A2": ('=COUNTIFS(S!A1:A9, "<>5", S!C1:C9, "W")', 4)})
    assert c.unresolved == {}
    assert got[("F", 1, 1)] == N(1) and got[("F", 2, 1)] == N(4)


def test_sumif_averageif_and_the_ifs_family():
    got, c = run({
        "A1": ('=SUMIF(S!C1:C9, "W", S!B1:B9)', 250), "A2": ('=SUMIF(S!C1:C9, "W", S!B1)', 250),  # resized
        "A3": ('=SUMIFS(S!B1:B9, S!C1:C9, "E", S!B1:B9, ">"&S!D2*10)', 140),
        "A4": ('=AVERAGEIF(S!C1:C9, "E", S!B1:B9)', 50), "A5": ('=AVERAGEIFS(S!B1:B9, S!C1:C9, "X")', "#DIV/0!"),
        "A6": ('=MINIFS(S!B1:B9, S!C1:C9, "W")', 10), "A7": ('=MAXIFS(S!B1:B9, S!C1:C9, "W")', 90),
        "A8": ('=SUMIF(S!B1:B9, ">=50")', 350),
    })
    assert c.unresolved == {}
    assert [got[("F", r, 1)] for r in range(1, 9)] == [
        N(250), N(250), N(140), N(50), ("e", None, "#DIV/0!"), N(10), N(90), N(350)]


def test_a_matched_error_makes_the_sum_an_error():
    got, c = run({"A1": ('=SUMIF(S!C1:C2, "W", S!A1:A2)', 5), "A2": ('=SUMIF(S!C8:C9, "W", S!A8:A9)', "#N/A")})
    assert c.unresolved == {}
    assert got[("F", 1, 1)] == N(5)  # A2 is the text "5": skipped, like Excel
    assert got[("F", 2, 1)] == ("e", None, "#N/A")


def test_criteria_from_an_empty_cell_is_unresolved():
    _, c = run({"A1": ("=COUNTIF(S!A1:A9, S!Z99)", 0)})
    assert c.unresolved[("F", 1, 1)] == "criteria from an empty cell not supported yet"
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_criteria.py -v`
Expected: FAIL. `KeyError` on `got[...]` because the cells are UNRESOLVED with "function COUNTIF not supported yet".

- [ ] **Step 3: Create `src/unexcel/fn/__init__.py`**

```python
"""The Excel function library. Importing this package registers every function in unexcel.functions.FUNCS."""
from unexcel.fn import criteria  # noqa: F401
```

- [ ] **Step 4: Create `src/unexcel/fn/criteria.py`**

```python
"""COUNTIF, COUNTIFS, SUMIF, SUMIFS, AVERAGEIF, AVERAGEIFS, MINIFS, MAXIFS.

A criterion is a number (matches equal numbers, and text that reads as that number), TRUE/FALSE, or text
that may start with =, <>, <, >, <= or >= and may use the wildcards * and ? (~ escapes them). Text criteria
can be built while the sheet runs (">"&B1), so they are parsed in SQL, not in Python.
Empty cells have no row in the cells table. They match "" and "=" and "<>x", so they are counted from
the range's size.
"""
from __future__ import annotations

from unexcel.expr import CELLS, NO_ERR, Expr, errs, lit_str, parse_num, txt, where_area
from unexcel.functions import arity, note_text_numbers, register
from unexcel.model import Unsupported
from unexcel.parser import RefNode
from unexcel.refs import Area, resolve

OPS = (">=", "<=", "<>", ">", "<", "=")


def like_pattern(rhs: str) -> str:
    """Excel wildcards -> a LIKE pattern (ESCAPE '\\'): * -> %, ? -> _, ~* and ~? are literal."""
    p = f"lower({rhs})"
    for a, b in (("'\\'", "'\\\\'"), ("'%'", "'\\%'"), ("'_'", "'\\_'"), ("'~*'", "chr(1)"), ("'~?'", "chr(2)"),
                 ("'*'", "'%'"), ("'?'", "'_'"), ("chr(1)", "'*'"), ("chr(2)", "'?'")):
        p = f"replace({p}, {a}, {b})"
    return p


def match_sql(alias: str, crit: Expr, q: str) -> tuple[str, str, str]:
    """(SQL true when cell `alias` meets the criterion, SQL true when an empty cell does, a one-row table `q`).

    A text criterion is parsed once, in the one-row table `q` the caller puts in its FROM clause:
    repeating the criterion SQL inside every comparison overflowed DuckDB's expression depth.
    """
    if crit.err != NO_ERR:
        raise Unsupported("criteria that can be an error not supported yet")
    if crit.type == "blank":
        raise Unsupported("criteria from an empty cell not supported yet")
    k, n, t = f"{alias}.kind", f"{alias}.num", f"COALESCE({alias}.txt, '')"  # the seed CSV reads '' back as NULL
    if crit.type == "num":
        return f"COALESCE(({k} = 'n' AND {n} = {crit.sql}) OR ({k} = 's' AND {parse_num(t)} = {crit.sql}), FALSE)", \
            "FALSE", f"(SELECT 1 AS one) {q}"
    if crit.type == "bool":
        return f"COALESCE({k} = 'b' AND {n} = CAST({crit.sql} AS DOUBLE), FALSE)", "FALSE", f"(SELECT 1 AS one) {q}"
    ops = "(CASE " + " ".join(f"WHEN c LIKE '{o}%' THEN '{o}'" for o in OPS) + " ELSE '' END)"
    parsed = (f"(SELECT op, rhs, {parse_num('rhs')} AS rn, "
              f"(CASE upper(rhs) WHEN 'TRUE' THEN 1.0 WHEN 'FALSE' THEN 0.0 END) AS rb, {like_pattern('rhs')} AS pat "
              f"FROM (SELECT op, substr(c, length(op) + 1) AS rhs FROM (SELECT c, {ops} AS op "
              f"FROM (SELECT {txt(crit).sql} AS c)))) {q}")
    op, rhs, rn, rb = f"{q}.op", f"{q}.rhs", f"{q}.rn", f"{q}.rb"

    def order(a: str, b: str) -> str:
        return " ".join(["(CASE", *(f"WHEN {op} = '{o}' THEN {a} {o} {b}" for o in (">", "<", ">=", "<=")), "END)"])

    eq = (f"COALESCE(CASE WHEN {rn} IS NOT NULL THEN ({k} = 'n' AND {n} = {rn}) OR ({k} = 's' AND {parse_num(t)} = {rn}) "
          f"WHEN {rb} IS NOT NULL THEN {k} = 'b' AND {n} = {rb} "
          f"ELSE {k} = 's' AND lower({t}) LIKE {q}.pat ESCAPE '\\' END, FALSE)")
    ordered = (f"COALESCE(CASE WHEN {rn} IS NOT NULL THEN {k} = 'n' AND {order(n, rn)} "
               f"WHEN {rb} IS NOT NULL THEN {k} = 'b' AND {order(n, rb)} "
               f"ELSE {k} = 's' AND {order(f'lower({t})', f'lower({rhs})')} END, FALSE)")
    match = f"(CASE WHEN {op} = '<>' THEN NOT {eq} WHEN {op} IN ('', '=') THEN {eq} ELSE {ordered} END)"
    blank = f"(({op} IN ('', '=') AND {rhs} = '') OR ({op} = '<>' AND {rhs} <> ''))"
    return match, blank, parsed


def _range(c, node, sheet: str, fname: str) -> tuple[Area, list]:
    if not isinstance(node, RefNode):
        raise Unsupported(f"{fname} needs a cell range")
    area = resolve(node.text, sheet, c.book)
    return area, c.ready(area)


def _column(c, node, sheet: str, fname: str, height: int | None = None) -> tuple[Area, list]:
    area, keys = _range(c, node, sheet, fname)
    if area.c1 != area.c2 or (height is not None and area.r2 - area.r1 != height):
        raise Unsupported(f"{fname} ranges must be single columns of the same height")
    return area, keys


def _pairs(c, args, sheet: str, fname: str, height: int | None) -> list[tuple[Area, Expr]]:
    pairs = []
    for i in range(0, len(args), 2):
        area, _ = _column(c, args[i], sheet, fname, height)
        height = area.r2 - area.r1
        pairs.append((area, c.expr(args[i + 1], sheet)))
    return pairs


def _joined(pairs, base: str, base_r1) -> tuple[str, str, str]:
    """(one-row criteria tables, LEFT JOINs of each criteria range to the driving rows, AND of conditions, AND of blanks)."""
    parsed, joins, conds, blanks = [], [], [], []
    for i, (area, crit) in enumerate(pairs, 1):
        k = f"k{i}"
        joins.append(f"LEFT JOIN {CELLS} {k} ON {k}.sheet = {lit_str(area.sheet).sql} "
                     f"AND {k}.col_num = {area.c1} AND {k}.row_num = {base} + {area.r1 - base_r1}")
        match, blank, source = match_sql(k, crit, f"q{i}")
        parsed.append(source)
        conds.append(f"(CASE WHEN {k}.kind IS NULL THEN {blank} ELSE {match} END)")
        blanks.append(blank)
    return parsed, " ".join(joins), " AND ".join(conds), " AND ".join(blanks)


@register("COUNTIF")
def countif(c, args, sheet):
    arity("COUNTIF", args, 2, 2)
    area, _ = _range(c, args[0], sheet, "COUNTIF")
    match, blank, source = match_sql("x", c.expr(args[1], sheet), "q1")
    size = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1)
    inside = f"FROM {CELLS} x WHERE {where_area(area, 'x')}"
    return Expr(f"(SELECT (SELECT COUNT(*) FILTER (WHERE {match}) {inside}) "
                f"+ CASE WHEN {blank} THEN {size} - (SELECT COUNT(*) {inside}) ELSE 0 END FROM {source})", "num")


@register("COUNTIFS")
def countifs(c, args, sheet):
    if not args or len(args) % 2:
        raise Unsupported("COUNTIFS needs criteria pairs")
    pairs = _pairs(c, args, sheet, "COUNTIFS", None)
    height = pairs[0][0].r2 - pairs[0][0].r1 + 1
    rows = " UNION ".join(f"SELECT row_num - {a.r1} AS off FROM {CELLS} WHERE {where_area(a)}" for a, _ in pairs)
    tables, joins, cond, blanks = _joined(pairs, "o.off", 0)
    return Expr(f"(WITH o AS ({rows}) SELECT (SELECT COUNT(*) FILTER (WHERE {cond}) FROM o {joins}) "
                f"+ CASE WHEN {blanks} THEN {height} - (SELECT COUNT(*) FROM o) ELSE 0 END FROM {', '.join(tables)})", "num")


def _aggregate(c, target: Area, pairs, how: str) -> Expr:
    """SUM / AVERAGE / MIN / MAX of the target cells whose rows meet every criterion."""
    keys = c.ready(target)
    note_text_numbers(c, keys)
    tables, joins, cond, _ = _joined(pairs, "s.row_num", target.r1)
    crossed = " ".join(f"CROSS JOIN {t}" for t in tables)
    src = f"FROM {CELLS} s {crossed} {joins} WHERE {where_area(target, 's')} AND {cond}"
    numbers = "FILTER (WHERE s.kind = 'n')"
    err = NO_ERR
    if any(c.may_err(k) for k in keys):  # a matched error cell makes the result that error
        err = f"(SELECT s.txt {src} AND s.kind = 'e' ORDER BY s.row_num LIMIT 1)"
    if how == "SUM":
        return Expr(f"(SELECT COALESCE(SUM(s.num) {numbers}, CAST(0 AS DOUBLE)) {src})", "num", err)
    if how in ("MIN", "MAX"):
        return Expr(f"(SELECT COALESCE({how}(s.num) {numbers}, CAST(0 AS DOUBLE)) {src})", "num", err)
    count = f"(SELECT COUNT(*) {numbers} {src})"
    return Expr(f"(SELECT SUM(s.num) {numbers} / NULLIF(COUNT(*) {numbers}, 0) {src})", "num",
                errs(err, f"(CASE WHEN {count} = 0 THEN '#DIV/0!' END)"))


def _resized(area: Area, like: Area) -> Area:
    """Excel's SUMIF uses the sum range's top-left cell and the criteria range's size."""
    return Area(area.sheet, area.r1, area.c1, area.r1 + (like.r2 - like.r1), area.c1)


def _if(how: str):
    def fn(c, args, sheet):
        fname = f"{how}IF" if how in ("SUM", "AVERAGE") else how
        arity(fname, args, 2, 3)
        area, _ = _column(c, args[0], sheet, fname)
        target = area
        if len(args) == 3:
            if not isinstance(args[2], RefNode):
                raise Unsupported(f"{fname} needs a cell range")
            target = _resized(resolve(args[2].text, sheet, c.book), area)
        return _aggregate(c, target, [(area, c.expr(args[1], sheet))], how)
    return fn


def _ifs(how: str):
    def fn(c, args, sheet):
        fname = f"{how}IFS"
        if len(args) < 3 or len(args) % 2 == 0:
            raise Unsupported(f"{fname} needs a sum range and criteria pairs")
        target, _ = _column(c, args[0], sheet, fname)
        pairs = _pairs(c, args[1:], sheet, fname, target.r2 - target.r1)
        return _aggregate(c, target, pairs, how)
    return fn


register("SUMIF")(_if("SUM"))
register("AVERAGEIF")(_if("AVERAGE"))
register("SUMIFS")(_ifs("SUM"))
register("AVERAGEIFS")(_ifs("AVERAGE"))
register("MINIFS")(_ifs("MIN"))
register("MAXIFS")(_ifs("MAX"))
```

- [ ] **Step 5: Wire it in.** In `src/unexcel/compiler.py` replace
```python
from pathlib import Path

from unexcel.expr import
```
with
```python
from pathlib import Path

import unexcel.fn  # noqa: F401  (registers the function library)
from unexcel.expr import
```
In `src/unexcel/functions.py` delete the whole `sumifs` function: from the line `@register("SUMIFS")` down to the blank lines before `@register("VLOOKUP")`.

- [ ] **Step 6: Update the two Plan 2a tests this changes** (SUMIFS now supports comparisons and errors)

In `tests/test_errors.py` replace
```python
def test_sumifs_over_error_cells_is_unresolved():
    _, c = run({"S": {"A1": "x", "B1": "#N/A", "C1": ('=SUMIFS(B1:B1, A1:A1, "x")', "#N/A")}})
    assert c.unresolved[("S", 1, 3)] == "SUMIFS over cells that hold errors not supported yet"
```
with
```python
def test_sumifs_over_a_matched_error_cell_returns_that_error():
    got, c = run({"S": {"A1": "x", "B1": "#N/A", "C1": ('=SUMIFS(B1:B1, A1:A1, "x")', "#N/A")}})
    assert c.unresolved == {} and got[("S", 1, 3)] == E("#N/A")
```
In `tests/test_lookup.py` replace
```python
    assert "comparison/wildcard" in c.unresolved[("X", 1, 1)]
```
with
```python
    assert ("X", 1, 1) not in c.unresolved  # comparison criteria are supported since Plan 2b
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 8: Commit**

```bash
git add src/unexcel/fn src/unexcel/compiler.py src/unexcel/functions.py tests
git commit -m "feat: Excel criteria engine for COUNTIF(S), SUMIF(S), AVERAGEIF(S), MINIFS, MAXIFS" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Aggregates: AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, PRODUCT, MEDIAN, SUMPRODUCT, SUBTOTAL

**Files:**
- Create: `src/unexcel/fn/aggregate.py`
- Modify: `src/unexcel/fn/__init__.py`
- Test: `tests/test_aggregate.py`

**Interfaces:**
- Consumes: `Compiler.ready/may_err/expr`, `expr.num/errs/lit_str/where_area`, `functions.arity/note_text_numbers/register`.
- Produces: `aggregate.Rng(area, keys, skip)` with `.where(alias)`, `.agg(fn, kinds)`, `.err(c)`. Rules: ranges count numbers only; typed-in values are converted; an error in a range is the result; MIN/MAX of nothing is 0; AVERAGE of nothing is `#DIV/0!`; MEDIAN of nothing is `#NUM!`; SUMPRODUCT treats non-numbers as 0 and takes plain ranges only; SUBTOTAL 1-6 and 9 leave out other SUBTOTAL cells, 101-111 (hidden rows) are UNRESOLVED.

- [ ] **Step 1: Write the failing tests `tests/test_aggregate.py`**

```python
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute

DATA = {"A1": 4, "A2": "x", "A3": True, "A4": 8, "A5": "15,000", "B1": 2, "B2": 3, "B3": "", "C1": "#N/A"}


def run(formulas):
    book = book_of({"S": DATA, "F": formulas})
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def N(v):
    return ("n", float(v), None)


def test_aggregates_skip_text_and_logicals_in_ranges():
    got, c = run({
        "A1": ("=AVERAGE(S!A1:A9)", 6), "A2": ('=AVERAGE(S!A1:A2, "3", TRUE)', 2.6667), "A3": ("=MIN(S!A1:A9)", 4),
        "A4": ("=MAX(S!A1:A9, 20)", 20), "A5": ("=MAX(S!Z1:Z9)", 0), "A6": ("=COUNT(S!A1:B9)", 4),
        "A7": ("=COUNTA(S!A1:B9)", 8), "A8": ("=COUNTBLANK(S!B1:B4)", 2), "A9": ("=PRODUCT(S!A1:B9)", 192),
        "A10": ("=MEDIAN(S!A1:B9)", 3.5), "A11": ("=SUMPRODUCT(S!A1:A2, S!B1:B2)", 8),
        "A12": ("=AVERAGE(S!Z1:Z9)", "#DIV/0!"), "A13": ("=MAX(S!A1:A9, S!C1)", "#N/A"), "A14": ("=PRODUCT(S!Z1:Z2)", 0),
    })
    assert c.unresolved == {}
    assert got[("F", 2, 1)][1] == 8 / 3
    assert [got[("F", r, 1)] for r in (1, 3, 4, 5, 6, 7, 8, 9, 10, 11)] == [
        N(6), N(4), N(20), N(0), N(4), N(8), N(2), N(192), N(3.5), N(8)]
    assert got[("F", 12, 1)] == ("e", None, "#DIV/0!") and got[("F", 13, 1)] == ("e", None, "#N/A")
    assert got[("F", 14, 1)] == N(0)


def test_subtotal_leaves_out_other_subtotals():
    book = book_of({"S": {"A1": 1, "A2": 2, "A3": ("=SUBTOTAL(9, A1:A2)", 3), "A4": 10,
                          "A5": ("=SUBTOTAL(9, A1:A4)", 13), "A6": ("=SUBTOTAL(1, A1:A4)", 13 / 3),
                          "A7": ("=SUBTOTAL(109, A1:A4)", 13)}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert got[("S", 5, 1)] == N(13) and got[("S", 6, 1)][1] == 13 / 3
    assert compiled.unresolved[("S", 7, 1)] == "SUBTOTAL that ignores hidden rows not supported yet"
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_aggregate.py -v`
Expected: FAIL. the cells are UNRESOLVED ("function AVERAGE not supported yet" ...), so `c.unresolved` isn't empty.

- [ ] **Step 3: Create `src/unexcel/fn/aggregate.py`**

```python
"""AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, PRODUCT, MEDIAN, SUMPRODUCT, SUBTOTAL.

Like SUM: in a range only numbers count (text, TRUE/FALSE and empty cells are skipped); a value typed
straight into the arguments is converted (AVERAGE(1, "3") = 2). An error in a range is the result.
"""
from __future__ import annotations

from dataclasses import dataclass, replace

from unexcel.expr import BLANK, CELLS, NO_ERR, Expr, errs, lit_str, num, where_area
from unexcel.functions import arity, note_text_numbers, register
from unexcel.model import Unsupported
from unexcel.parser import Num, RefNode
from unexcel.refs import Area, resolve


@dataclass(frozen=True)
class Rng:
    area: Area
    keys: tuple
    skip: tuple = ()  # cells left out (SUBTOTAL leaves out other SUBTOTALs)

    def where(self, alias: str = "") -> str:
        p = f"{alias}." if alias else ""
        return where_area(self.area, alias) + "".join(
            f" AND NOT ({p}row_num = {r} AND {p}col_num = {col})" for _, r, col in self.skip)

    def agg(self, fn: str, kinds: str = "('n')") -> str:
        return f"(SELECT {fn} FROM {CELLS} WHERE {self.where()} AND kind IN {kinds})"

    def err(self, c) -> str:
        if not any(c.may_err(k) for k in self.keys if k not in self.skip):
            return NO_ERR
        return (f"(SELECT txt FROM {CELLS} WHERE {self.where()} AND kind = 'e' "
                f"ORDER BY row_num, col_num LIMIT 1)")


def _rng(c, node, sheet: str, skip_subtotals: bool = False) -> Rng:
    area = resolve(node.text, sheet, c.book)
    keys = tuple(c.ready(area))
    note_text_numbers(c, keys)
    skip = ()
    if skip_subtotals:
        skip = tuple(k for k in keys if "SUBTOTAL(" in (c.book.cells[k].formula or "").upper())
    return Rng(area, keys, skip)


def _split(c, args, sheet: str, skip_subtotals: bool = False) -> tuple[list[Rng], list[Expr]]:
    ranges, values = [], []
    for a in args:
        if isinstance(a, RefNode):
            ranges.append(_rng(c, a, sheet, skip_subtotals))
        else:
            values.append(num(c.expr(a, sheet) if a is not None else BLANK))
    return ranges, values


def _errs(c, ranges, values) -> str:
    return errs(*(r.err(c) for r in ranges), *(v.err for v in values))


def _average(c, ranges, values) -> Expr:
    total = " + ".join([*(r.agg("COALESCE(SUM(num), 0)") for r in ranges), *(v.sql for v in values)]) or "0"
    count = " + ".join([*(r.agg("COUNT(*)") for r in ranges), *("1" for _ in values)]) or "0"
    return Expr(f"(({total}) / NULLIF({count}, 0))", "num",
                errs(_errs(c, ranges, values), f"(CASE WHEN ({count}) = 0 THEN '#DIV/0!' END)"))


def _extreme(c, ranges, values, how: str) -> Expr:
    parts = [*(r.agg(f"{how}(num)") for r in ranges), *(v.sql for v in values)]
    fn = "LEAST" if how == "MIN" else "GREATEST"  # both skip NULLs in DuckDB; no numbers at all -> 0, like Excel
    return Expr(f"COALESCE({fn}({', '.join(parts)}), CAST(0 AS DOUBLE))", "num", _errs(c, ranges, values))


def _count(c, ranges, values) -> Expr:
    if values:
        raise Unsupported("COUNT of values typed into the formula not supported yet")
    return Expr("(" + " + ".join(r.agg("CAST(COUNT(*) AS DOUBLE)") for r in ranges) + ")", "num")


def _counta(c, ranges, values) -> Expr:
    parts = [r.agg("CAST(COUNT(*) AS DOUBLE)", "('n', 's', 'b', 'e')") for r in ranges] + ["1" for _ in values]
    return Expr("CAST((" + " + ".join(parts) + ") AS DOUBLE)", "num")


def _product(c, ranges, values) -> Expr:
    parts = [f"COALESCE({r.agg('product(num)')}, 1)" for r in ranges] + [v.sql for v in values]
    count = " + ".join([*(r.agg("COUNT(*)") for r in ranges), *("1" for _ in values)])
    return Expr(f"(CASE WHEN ({count}) = 0 THEN CAST(0 AS DOUBLE) ELSE {' * '.join(parts)} END)", "num",
                _errs(c, ranges, values))


def _sum(c, ranges, values) -> Expr:
    parts = [*(r.agg("COALESCE(SUM(num), 0)") for r in ranges), *(v.sql for v in values)]
    return Expr("(" + " + ".join(parts) + ")" if parts else "CAST(0 AS DOUBLE)", "num", _errs(c, ranges, values))


def _plain(name, how):
    def fn(c, args, sheet):
        if not args:
            raise Unsupported(f"{name} needs arguments")
        ranges, values = _split(c, args, sheet)
        result = how(c, ranges, values)
        if name not in ("COUNT", "COUNTA"):
            r, v = iter(ranges), iter(values)
            ordered = [next(r).err(c) if isinstance(a, RefNode) else next(v).err for a in args]
            result = replace(result, err=errs(*ordered, result.err))
        return result
    register(name)(fn)


_plain("AVERAGE", _average)
_plain("MIN", lambda c, r, v: _extreme(c, r, v, "MIN"))
_plain("MAX", lambda c, r, v: _extreme(c, r, v, "MAX"))
_plain("COUNT", _count)
_plain("COUNTA", _counta)
_plain("PRODUCT", _product)


@register("COUNTBLANK")
def countblank(c, args, sheet):
    arity("COUNTBLANK", args, 1, 1)
    if not isinstance(args[0], RefNode):
        raise Unsupported("COUNTBLANK needs a cell range")
    r = _rng(c, args[0], sheet)
    a = r.area
    size = (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1)
    filled = (f"(SELECT COUNT(*) FROM {CELLS} WHERE {r.where()} "
              f"AND NOT (kind = 's' AND COALESCE(txt, '') = ''))")  # "" counts as blank here, like Excel
    return Expr(f"CAST({size} - {filled} AS DOUBLE)", "num")


@register("MEDIAN")
def median(c, args, sheet):
    if not args or not all(isinstance(a, RefNode) for a in args):
        raise Unsupported("MEDIAN of values typed into the formula not supported yet")
    ranges = [_rng(c, a, sheet) for a in args]
    rows = " UNION ALL ".join(f"SELECT num FROM {CELLS} WHERE {r.where()} AND kind = 'n'" for r in ranges)
    values = f"FROM ({rows}) median_values"
    return Expr(f"(SELECT CAST(median(num) AS DOUBLE) {values})", "num",
                errs(_errs(c, ranges, []), f"(CASE WHEN (SELECT COUNT(*) {values}) = 0 THEN '#NUM!' END)"))


@register("SUMPRODUCT")
def sumproduct(c, args, sheet):
    if not args or not all(isinstance(a, RefNode) for a in args):
        raise Unsupported("SUMPRODUCT of array expressions not supported yet")
    ranges = [_rng(c, a, sheet) for a in args]
    first = ranges[0].area
    shape = (first.r2 - first.r1, first.c2 - first.c1)
    if any((r.area.r2 - r.area.r1, r.area.c2 - r.area.c1) != shape for r in ranges):
        return Expr("NULL", "num", "'#VALUE!'")  # Excel: ranges of different sizes
    joins = "".join(
        f" JOIN {CELLS} a{i} ON a{i}.sheet = {lit_str(r.area.sheet).sql} "
        f"AND a{i}.row_num = a0.row_num + {r.area.r1 - first.r1} AND a{i}.col_num = a0.col_num + {r.area.c1 - first.c1} "
        f"AND a{i}.kind = 'n'" for i, r in enumerate(ranges[1:], 1))
    product = " * ".join(f"a{i}.num" for i in range(len(ranges)))  # blanks, text and TRUE/FALSE count as 0
    return Expr(f"(SELECT COALESCE(SUM({product}), CAST(0 AS DOUBLE)) FROM {CELLS} a0{joins} "
                f"WHERE {where_area(first, 'a0')} AND a0.kind = 'n')", "num", _errs(c, ranges, []))


SUBTOTALS = {1: _average, 2: _count, 3: _counta, 4: lambda c, r, v: _extreme(c, r, v, "MAX"),
             5: lambda c, r, v: _extreme(c, r, v, "MIN"), 6: _product, 9: _sum}


@register("SUBTOTAL")
def subtotal(c, args, sheet):
    if len(args) < 2 or not isinstance(args[0], Num):
        raise Unsupported("SUBTOTAL needs a function number and ranges")
    code = int(args[0].value)
    if code > 100:
        raise Unsupported("SUBTOTAL that ignores hidden rows not supported yet")
    if code not in SUBTOTALS:
        raise Unsupported(f"SUBTOTAL function {code} not supported yet")
    if not all(isinstance(a, RefNode) for a in args[1:]):
        raise Unsupported("SUBTOTAL needs cell ranges")
    ranges = [_rng(c, a, sheet, skip_subtotals=True) for a in args[1:]]
    return SUBTOTALS[code](c, ranges, [])
```

- [ ] **Step 4: Register it: replace `src/unexcel/fn/__init__.py`**

```python
"""The Excel function library. Importing this package registers every function in unexcel.functions.FUNCS."""
from unexcel.fn import aggregate, criteria  # noqa: F401
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 6: Commit**

```bash
git add src/unexcel/fn tests/test_aggregate.py
git commit -m "feat: AVERAGE MIN MAX COUNT COUNTA COUNTBLANK PRODUCT MEDIAN SUMPRODUCT SUBTOTAL" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: Lookups: VLOOKUP (approximate too), HLOOKUP, MATCH, INDEX, XLOOKUP, CHOOSE

**Files:**
- Create: `src/unexcel/fn/lookup.py`
- Replace: `src/unexcel/functions.py` (VLOOKUP and its helpers move to `fn/lookup.py`)
- Modify: `src/unexcel/fn/__init__.py`, `tests/test_lookup.py`
- Test: `tests/test_lookups.py`

**Interfaces:**
- Consumes: `criteria.like_pattern`, `expr.any_of/unify_all/num/errs/lit_str/where_area`, `Compiler.ready/kind_of/expr`.
- Produces: `lookup.int_literal(node) -> int | None` (a whole number written in the formula; `-1` parses as minus applied to 1). Used by Task 8.
- Rules: the found value's type comes from the cells of the result range (one type; else `"any"`); an empty cell found reads as 0; exact text matching ignores case and honours wildcards (XLOOKUP only with match_mode 2); approximate match picks the largest value <= x (MATCH -1: smallest >= x); not found is `#N/A`; INDEX out of range is `#REF!`; CHOOSE out of range is `#VALUE!`.

- [ ] **Step 1: Write the failing tests `tests/test_lookups.py`**

```python
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute

# Rates: brackets sorted ascending (A), rate (B), label (C, one blank)
DATA = {"A1": 0, "B1": 0.1, "C1": "low", "A2": 1000, "B2": 0.2, "C2": "mid", "A3": 5000, "B3": 0.3,
        "A4": "x", "B4": "#N/A", "E1": "Jan", "F1": "Feb", "G1": "Mar", "E2": 1, "F2": 2, "G2": 3}


def run(formulas):
    book = book_of({"R": DATA, "F": formulas})
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def N(v):
    return ("n", float(v), None)


def test_approximate_lookups_find_the_bracket():
    got, c = run({"A1": ("=VLOOKUP(2500, R!A1:B3, 2)", 0.2), "A2": ("=VLOOKUP(2500, R!A1:B3, 2, TRUE)", 0.2),
                  "A3": ("=MATCH(5000, R!A1:A3)", 3), "A4": ("=MATCH(-1, R!A1:A3, 1)", "#N/A"),
                  "A5": ("=MATCH(999, R!A1:A3, -1)", 2)})
    assert c.unresolved == {}
    assert [got[("F", r, 1)] for r in range(1, 6)] == [N(0.2), N(0.2), N(3), ("e", None, "#N/A"), N(2)]


def test_exact_lookups_wildcards_and_errors():
    got, c = run({"A1": ('=MATCH("m*", R!C1:C3, 0)', 2), "A2": ('=MATCH("Feb", R!E1:G1, 0)', 2),
                  "A3": ('=HLOOKUP("mar", R!E1:G2, 2, FALSE)', 3), "A4": ('=VLOOKUP("x", R!A1:B4, 2, FALSE)', "#N/A"),
                  "A5": ('=XLOOKUP("mid", R!C1:C3, R!A1:A3)', 1000), "A6": ('=XLOOKUP("m*", R!C1:C3, R!A1:A3)', "#N/A"),
                  "A7": ('=XLOOKUP("zz", R!C1:C3, R!A1:A3, "none")', "none"), "A8": ("=XLOOKUP(3000, R!A1:A3, R!B1:B3, 0, -1)", 0.2)})
    assert c.unresolved == {}
    assert [got[("F", r, 1)] for r in range(1, 9)] == [
        N(2), N(2), N(3), ("e", None, "#N/A"), N(1000), ("e", None, "#N/A"), ("s", None, "none"), N(0.2)]


def test_index_choose_and_mixed_results():
    got, c = run({"A1": ("=INDEX(R!A1:B3, 2, 2)", 0.2), "A2": ("=INDEX(R!E1:G1, 3)", "Mar"), "A3": ("=INDEX(R!A1:A3, 9)", "#REF!"),
                  "A4": ("=INDEX(R!C1:C3, 3)", 0), "A5": ('=INDEX(R!A1:A3, MATCH("mid", R!C1:C3, 0))', 1000),
                  "A6": ('=CHOOSE(2, "a", 5, TRUE)', 5), "A7": ('=CHOOSE(4, "a", 5)', "#VALUE!"),
                  "A8": ("=INDEX(R!C1:C3, 1)", "low")})
    assert c.unresolved == {}
    assert [got[("F", r, 1)] for r in range(1, 9)] == [
        N(0.2), ("s", None, "Mar"), ("e", None, "#REF!"), N(0), N(1000), N(5), ("e", None, "#VALUE!"), ("s", None, "low")]
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_lookups.py -v`
Expected: FAIL. The cells are UNRESOLVED ("VLOOKUP approximate match not supported yet", "function MATCH not supported yet" ...).

- [ ] **Step 3: Create `src/unexcel/fn/lookup.py`**

```python
"""VLOOKUP, HLOOKUP, MATCH, INDEX, XLOOKUP, CHOOSE.

The value found is read from the cells table. Its type comes from the cells the range holds: one type ->
that type; mixed, or text with empty cells -> 'any' (known when the SQL runs). An empty cell found reads
as 0, like Excel. Exact matching of text ignores case and (except XLOOKUP's default) honours wildcards.
"""
from __future__ import annotations

from unexcel.expr import CELLS, NO_ERR, Expr, any_of, errs, lit_str, num, unify_all, where_area
from unexcel.fn.criteria import like_pattern
from unexcel.functions import arity, register
from unexcel.model import Unsupported
from unexcel.parser import Bool, Num, RefNode, Str, Unary
from unexcel.refs import Area, resolve


def _area(c, node, sheet: str, fname: str) -> tuple[Area, list]:
    if not isinstance(node, RefNode):
        raise Unsupported(f"{fname} needs a cell range")
    area = resolve(node.text, sheet, c.book)
    return area, c.ready(area)


def _type_of(c, area: Area, keys) -> str:
    kinds = {c.kind_of(k) for k in keys} - {"e"}  # an error found is returned as that error, whatever the type
    full = len(keys) == (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1)
    if kinds <= {"n"}:
        return "num"  # an empty cell found reads as 0
    if full and kinds == {"s"}:
        return "txt"
    if full and kinds == {"b"}:
        return "bool"
    return "any"


def _read(src: str, typ: str, err: str) -> Expr:
    """The cell `r` that the one-row query `src` finds; r is NULL when the cell is empty."""
    if typ == "num":
        return Expr(f"(SELECT COALESCE(r.num, 0) {src})", "num", err)
    if typ == "txt":
        return Expr(f"(SELECT COALESCE(r.txt, '') {src})", "txt", err)
    if typ == "bool":
        return Expr(f"(SELECT r.num <> 0 {src})", "bool", err)
    return any_of(f"(SELECT COALESCE(r.kind, 'n') {src})", f"(SELECT COALESCE(r.num, 0) {src})",
                  f"(SELECT COALESCE(r.txt, '') {src})", err)


def _found_error(src: str, missing: str) -> str:
    """'missing' when src finds no row, the cell's own error when it holds one, else NULL."""
    found = f"(SELECT CASE WHEN r.kind = 'e' THEN r.txt ELSE '' END {src})"
    return f"(CASE WHEN {found} IS NULL THEN '{missing}' WHEN {found} = '' THEN NULL ELSE {found} END)"


def _static_text(node) -> bool:
    return isinstance(node, Str) and not any(ch in node.value for ch in "*?~")


def _exact(k: str, x: Expr, wild: bool) -> str:
    """SQL true when cell `k` equals the lookup value x (text ignores case; wildcards when `wild`)."""
    text = f"lower(COALESCE({k}.txt, ''))"
    t = (f"{text} LIKE {like_pattern(x.txt if x.type == 'any' else x.sql)} ESCAPE '\\'" if wild
         else f"{text} = lower({x.txt if x.type == 'any' else x.sql})")
    if x.type == "num":
        return f"({k}.kind = 'n' AND {k}.num = {x.sql})"
    if x.type == "txt":
        return f"({k}.kind = 's' AND {t})"
    if x.type == "bool":
        return f"({k}.kind = 'b' AND {k}.num = CAST({x.sql} AS DOUBLE))"
    if x.type == "any":
        return (f"(CASE {x.kind} WHEN 'n' THEN {k}.kind = 'n' AND {k}.num = {x.sql} "
                f"WHEN 's' THEN {k}.kind = 's' AND {t} ELSE {k}.kind = 'b' AND {k}.num = {x.sql} END)")
    raise Unsupported(f"lookup of a value of type {x.type}")


def _near(k: str, x: Expr, cmp: str) -> tuple[str, str]:
    """(condition, sort key) for approximate match: values of x's type that are <= x (cmp '<=') or >= x."""
    if x.type == "num":
        return f"({k}.kind = 'n' AND {k}.num {cmp} {x.sql})", f"{k}.num"
    if x.type == "txt":
        return f"({k}.kind = 's' AND lower({k}.txt) {cmp} lower({x.sql}))", f"lower({k}.txt)"
    raise Unsupported("approximate match on this kind of value not supported yet")


def _search(c, x: Expr, line: Area, mode: str, wild: bool, last: bool | None = None) -> tuple[str, str]:
    """(FROM...LIMIT 1 tail selecting matching cell `k` of the 1-D range `line`, its position SQL)."""
    by_row = line.c1 == line.c2
    pos = "k.row_num" if by_row else "k.col_num"
    where = f"FROM {CELLS} k WHERE {where_area(line, 'k')}"
    if mode == "exact":
        cond, order = _exact("k", x, wild), f"{pos} {'DESC' if last else 'ASC'}"
    elif mode == "le":  # largest value <= x; the last of equal ones, as Excel's search on sorted data
        cond, key = _near("k", x, "<=")
        order = f"{key} DESC, {pos} {'ASC' if last is False else 'DESC'}"
    else:  # "ge": smallest value >= x
        cond, key = _near("k", x, ">=")
        order = f"{key} ASC, {pos} {'DESC' if last else 'ASC'}"
    return f"{where} AND {cond} ORDER BY {order} LIMIT 1", pos


def int_literal(node) -> int | None:
    """A whole number written in the formula (-1 parses as minus applied to 1), else None."""
    if isinstance(node, Num):
        return int(node.value)
    if isinstance(node, Unary) and isinstance(node.operand, Num):
        return -int(node.operand.value)
    return None


def _exact_flag(node) -> bool:
    """VLOOKUP/HLOOKUP 4th argument: FALSE or 0 -> exact; omitted, TRUE or 1 -> approximate."""
    if node is None:
        return True
    if isinstance(node, Bool):
        return not node.value
    if isinstance(node, Num):
        return node.value == 0
    raise Unsupported("lookup match type must be TRUE/FALSE written in the formula")


def _xlookup_x(c, args, sheet) -> Expr:
    x = c.expr(args[0], sheet)
    if x.type == "blank":
        raise Unsupported("lookup of an empty cell not supported yet")
    return x


def _table_lookup(fname: str, by_row: bool):
    def fn(c, args, sheet):
        arity(fname, args, 3, 4)
        exact = len(args) == 4 and _exact_flag(args[3])
        if not isinstance(args[2], Num):
            raise Unsupported(f"{fname} index must be a number written in the formula")
        table, _ = _area(c, args[1], sheet, fname)
        offset = int(args[2].value) - 1
        if by_row:
            line = Area(table.sheet, table.r1, table.c1, table.r2, table.c1)
            target = Area(table.sheet, table.r1, table.c1 + offset, table.r2, table.c1 + offset)
            ok = table.c1 <= target.c1 <= table.c2
            hit = f"r.row_num = k.row_num AND r.col_num = {target.c1}"
        else:
            line = Area(table.sheet, table.r1, table.c1, table.r1, table.c2)
            target = Area(table.sheet, table.r1 + offset, table.c1, table.r1 + offset, table.c2)
            ok = table.r1 <= target.r1 <= table.r2
            hit = f"r.col_num = k.col_num AND r.row_num = {target.r1}"
        if not ok:
            raise Unsupported(f"{fname} {'column' if by_row else 'row'} index outside the table")
        x = _xlookup_x(c, args, sheet)
        tail, _ = _search(c, x, line, "exact" if exact else "le", wild=exact)
        tail = tail.replace(f"FROM {CELLS} k WHERE",
                            f"FROM {CELLS} k LEFT JOIN {CELLS} r ON r.sheet = k.sheet AND {hit} WHERE", 1)
        keys = c.ready(target)
        return _read(tail, _type_of(c, target, keys), errs(x.err, _found_error(tail, "#N/A")))
    return fn


register("VLOOKUP")(_table_lookup("VLOOKUP", by_row=True))
register("HLOOKUP")(_table_lookup("HLOOKUP", by_row=False))


def _line(c, node, sheet: str, fname: str) -> Area:
    area, _ = _area(c, node, sheet, fname)
    if area.r1 != area.r2 and area.c1 != area.c2:
        raise Unsupported(f"{fname} needs a single row or column")
    return area


@register("MATCH")
def match(c, args, sheet):
    arity("MATCH", args, 2, 3)
    kind = 1 if len(args) == 2 or args[2] is None else int_literal(args[2])
    if kind not in (-1, 0, 1):
        raise Unsupported("MATCH type must be -1, 0 or 1 written in the formula")
    line = _line(c, args[1], sheet, "MATCH")
    x = _xlookup_x(c, args, sheet)
    tail, pos = _search(c, x, line, {0: "exact", 1: "le", -1: "ge"}[kind], wild=True)
    first = line.r1 if line.c1 == line.c2 else line.c1
    return Expr(f"(SELECT CAST({pos} - {first} + 1 AS DOUBLE) {tail})", "num",
                errs(x.err, f"(CASE WHEN (SELECT 1 {tail}) IS NULL THEN '#N/A' END)"))


@register("XLOOKUP")
def xlookup(c, args, sheet):
    arity("XLOOKUP", args, 3, 6)
    modes = []
    for i, default in ((4, 0), (5, 1)):
        node = args[i] if len(args) > i else None
        mode = default if node is None else int_literal(node)
        if mode is None:
            raise Unsupported("XLOOKUP modes must be numbers written in the formula")
        modes.append(mode)
    match_mode, search_mode = modes
    if match_mode not in (0, -1, 1, 2) or search_mode not in (1, -1):
        raise Unsupported("XLOOKUP binary search modes not supported yet")
    look = _line(c, args[1], sheet, "XLOOKUP")
    back = _line(c, args[2], sheet, "XLOOKUP")
    by_row = look.c1 == look.c2
    if (by_row and (back.c1 != back.c2 or back.r2 - back.r1 != look.r2 - look.r1)) or \
            (not by_row and (back.r1 != back.r2 or back.c2 - back.c1 != look.c2 - look.c1)):
        raise Unsupported("XLOOKUP return range must be one row or column the size of the lookup range")
    x = _xlookup_x(c, args, sheet)
    mode = {0: "exact", 2: "exact", -1: "le", 1: "ge"}[match_mode]
    tail, _ = _search(c, x, look, mode, wild=match_mode == 2, last=search_mode == -1)
    hit = (f"r.row_num = k.row_num + {back.r1 - look.r1} AND r.col_num = {back.c1}" if by_row
           else f"r.col_num = k.col_num + {back.c1 - look.c1} AND r.row_num = {back.r1}")
    tail = tail.replace(f"FROM {CELLS} k WHERE",
                        f"FROM {CELLS} k LEFT JOIN {CELLS} r ON r.sheet = {lit_str(back.sheet).sql} AND {hit} WHERE", 1)
    keys = c.ready(back)
    found = _read(tail, _type_of(c, back, keys), errs(x.err, _found_error(tail, "#N/A")))
    if len(args) < 4 or args[3] is None:
        return found
    missing = c.expr(args[3], sheet)
    found, missing = unify_all([found, missing])
    hit_sql = f"((SELECT 1 {tail}) IS NOT NULL)"
    pick_err = f"(CASE WHEN {hit_sql} THEN {found.err} ELSE {missing.err} END)"

    def case(a: str, b: str) -> str:
        return f"(CASE WHEN {hit_sql} THEN {a} ELSE {b} END)"
    return Expr(case(found.sql, missing.sql), found.type, errs(x.err, pick_err), case(found.txt, missing.txt),
                case(found.kind, missing.kind))


@register("INDEX")
def index(c, args, sheet):
    arity("INDEX", args, 2, 3)
    area, keys = _area(c, args[0], sheet, "INDEX")
    height, width = area.r2 - area.r1 + 1, area.c2 - area.c1 + 1
    first = num(c.expr(args[1], sheet))
    second = num(c.expr(args[2], sheet)) if len(args) == 3 and args[2] is not None else None
    if second is None:
        if height == 1:
            row, col = Expr("CAST(1 AS DOUBLE)", "num"), first  # one row: the number picks the column
        elif width == 1:
            row, col = first, Expr("CAST(1 AS DOUBLE)", "num")
        else:
            raise Unsupported("INDEX of a block needs a row and a column")
    else:
        row, col = first, second
    r = f"(CASE WHEN trunc({row.sql}) BETWEEN 1 AND {height} THEN CAST(trunc({row.sql}) AS BIGINT) END)"
    k = f"(CASE WHEN trunc({col.sql}) BETWEEN 1 AND {width} THEN CAST(trunc({col.sql}) AS BIGINT) END)"
    src = (f"FROM (SELECT 1) one LEFT JOIN {CELLS} r ON r.sheet = {lit_str(area.sheet).sql} "
           f"AND r.row_num = {area.r1} - 1 + {r} AND r.col_num = {area.c1} - 1 + {k}")
    bounds = f"(CASE WHEN {r} IS NULL OR {k} IS NULL OR {r} < 1 OR {r} > {height} OR {k} < 1 OR {k} > {width} THEN '#REF!' END)"
    found = f"(SELECT CASE WHEN r.kind = 'e' THEN r.txt END {src})"
    return _read(src, _type_of(c, area, keys), errs(row.err, col.err, bounds, found))


@register("CHOOSE")
def choose(c, args, sheet):
    if len(args) < 2:
        raise Unsupported("CHOOSE needs an index and values")
    i = num(c.expr(args[0], sheet))
    values = unify_all([c.expr(a, sheet) for a in args[1:]])
    n = f"TRY_CAST(trunc({i.sql}) AS BIGINT)"

    def case(parts: list[str]) -> str:
        if len(set(parts)) == 1:
            return parts[0]
        return "(CASE " + " ".join(f"WHEN {n} = {j} THEN {p}" for j, p in enumerate(parts, 1)) + " END)"
    bad = f"(CASE WHEN {n} IS NULL OR {n} < 1 OR {n} > {len(values)} THEN '#VALUE!' END)"
    err = errs(i.err, bad, case([v.err for v in values]) if any(v.err != NO_ERR for v in values) else NO_ERR)
    return Expr(case([v.sql for v in values]), values[0].type, err, case([v.txt for v in values]),
                case([v.kind for v in values]))
```

- [ ] **Step 4: Replace `src/unexcel/functions.py`** (registry, helpers, SUM, IF, ROUND, AND, OR, IFERROR stay here)

```python
"""Excel function -> SQL. Each entry: fn(compiler, args, sheet) -> Expr.

Ranges are read from the cells table (inputs plus the formula cells of earlier batches).
Compiler.ready() checks every formula cell in a range was computed in an earlier batch.
Every function returns its Excel error in Expr.err; see expr.errs for the order rule.
"""
from __future__ import annotations

from unexcel.expr import CELLS, NO_ERR, ZERO, Expr, any_of, as_bool, errs, lit_str, num, pick, unify, where_area
from unexcel.model import Unsupported, a1
from unexcel.parser import Bool, Num, RefNode, Str
from unexcel.refs import Area, resolve

FUNCS = {}


def register(name):
    def deco(fn):
        FUNCS[name] = fn
        return fn

    return deco


def arity(name, args, lo, hi):
    if not lo <= len(args) <= hi:
        raise Unsupported(f"{name} needs {lo}-{hi} arguments, got {len(args)}")


def _looks_numeric(text: str) -> bool:
    try:
        float(text.replace(",", "").strip())
        return True
    except ValueError:
        return False


def note_text_numbers(c, keys):
    """Excel skips text in SUM-like ranges, even text that looks like a number. Tell the user."""
    for key in keys:
        cell = c.book.cells[key]
        if cell.formula is None and cell.kind == "s" and _looks_numeric(cell.value):
            c.findings.append(f"{a1(key)}: number stored as text ({cell.value!r}) is skipped by Excel")


def first_error(c, area: Area, keys) -> str:
    """SQL for the first error in a range, row by row, or NO_ERR when no cell in it can hold one."""
    if not any(c.may_err(k) for k in keys):
        return NO_ERR
    return f"(SELECT txt FROM {CELLS} WHERE {where_area(area)} AND kind = 'e' ORDER BY row_num, col_num LIMIT 1)"


def range_sum(c, area: Area) -> Expr:
    """SUM over a range: text, TRUE/FALSE and blanks are skipped; the first error wins."""
    keys = c.ready(area)
    note_text_numbers(c, keys)
    return Expr(f"(SELECT COALESCE(SUM(num), CAST(0 AS DOUBLE)) FROM {CELLS} WHERE {where_area(area)} AND kind = 'n')",
                "num", first_error(c, area, keys))


@register("SUM")
def sum_(c, args, sheet):
    parts = [range_sum(c, resolve(arg.text, sheet, c.book)) if isinstance(arg, RefNode) else num(c.expr(arg, sheet))
             for arg in args]
    if not parts:
        return ZERO
    return Expr("(" + " + ".join(p.sql for p in parts) + ")", "num", errs(*(p.err for p in parts)))


@register("IF")
def if_(c, args, sheet):
    arity("IF", args, 2, 3)
    cond = as_bool(c.expr(args[0], sheet))
    yes = c.expr(args[1], sheet)
    no = c.expr(args[2], sheet) if len(args) == 3 else Expr("FALSE", "bool")
    yes, no = unify(yes, no)
    taken = NO_ERR if yes.err == no.err == NO_ERR else f"(CASE WHEN {cond.sql} THEN {yes.err} ELSE {no.err} END)"
    return pick(cond.sql, yes, no, errs(cond.err, taken))


@register("ROUND")
def round_(c, args, sheet):
    arity("ROUND", args, 2, 2)
    x, digits = num(c.expr(args[0], sheet)), num(c.expr(args[1], sheet))
    # DuckDB ROUND on DOUBLE rounds half away from zero like Excel (2.675 -> 2.68, verified on 1.5.6)
    return Expr(f"ROUND({x.sql}, CAST(TRUNC({digits.sql}) AS INTEGER))", "num", errs(x.err, digits.err))


def _logic(name, joiner):
    def fn(c, args, sheet):
        if not args:
            raise Unsupported(f"{name} needs arguments")
        parts = [as_bool(c.expr(a, sheet)) for a in args]
        return Expr("(" + f" {joiner} ".join(p.sql for p in parts) + ")", "bool", errs(*(p.err for p in parts)))

    return fn


FUNCS["AND"] = _logic("AND", "AND")
FUNCS["OR"] = _logic("OR", "OR")


@register("IFERROR")
def iferror(c, args, sheet):
    arity("IFERROR", args, 2, 2)
    value, fallback = unify(c.expr(args[0], sheet), c.expr(args[1], sheet))
    if value.err == NO_ERR:
        return value
    ok = f"(({value.err}) IS NULL)"
    err = NO_ERR if fallback.err == NO_ERR else f"(CASE WHEN {ok} THEN NULL ELSE {fallback.err} END)"
    return pick(ok, value, fallback, err)
```

- [ ] **Step 5: Register it: replace `src/unexcel/fn/__init__.py`**

```python
"""The Excel function library. Importing this package registers every function in unexcel.functions.FUNCS."""
from unexcel.fn import aggregate, criteria, lookup  # noqa: F401
```

- [ ] **Step 6: Update the Plan 1 test this changes** (approximate VLOOKUP is supported now)

In `tests/test_lookup.py` replace
```python
    assert c.unresolved[("X", 3, 2)] == "VLOOKUP approximate match not supported yet"
```
with
```python
    assert got[("X", 3, 2)] == ("n", 0.02, None)  # approximate match: largest value <= "south"
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 8: Commit**

```bash
git add src/unexcel/fn src/unexcel/functions.py tests
git commit -m "feat: lookups - approximate VLOOKUP, HLOOKUP, MATCH, INDEX, XLOOKUP, CHOOSE" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: Logic and information: IFS, SWITCH, NOT, XOR, IFNA, ISNA, ISERROR, ISERR, ISNUMBER, ISTEXT, ISLOGICAL, ISBLANK, NA, TRUE, FALSE

**Files:**
- Create: `src/unexcel/fn/logic.py`
- Modify: `src/unexcel/fn/__init__.py`
- Test: `tests/test_logic.py`

**Interfaces:**
- Consumes: `expr.as_bool/compare/error/errs/pick/unify/unify_all`.
- Produces: functions only. Rules: IFS/SWITCH with no match is `#N/A`; an error in a condition stops the chain; the IS functions never fail (an error in their argument is their answer); ISBLANK is TRUE only for an empty cell (a formula's "" is not blank).

- [ ] **Step 1: Write the failing tests `tests/test_logic.py`**

```python
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute


def run(cells):
    book = book_of({"S": cells})
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def B(v):
    return ("b", 1.0 if v else 0.0, None)


def test_ifs_switch_not_xor():
    got, c = run({"A1": 75, "A2": 0,
                  "B1": ('=IFS(A1>=90, "A", A1>=70, "B", TRUE, "C")', "B"), "B2": ('=IFS(A1>100, "X")', "#N/A"),
                  "B3": ('=IFS(1/A2>1, "X", TRUE, "Y")', "#DIV/0!"), "B4": ('=SWITCH(A1, 50, "half", 75, "most", "?")', "most"),
                  "B5": ('=SWITCH(A1, 1, "one")', "#N/A"), "B6": ('=SWITCH("b", "A", 1, "B", 2)', 2),
                  "B7": ("=NOT(A1>5)", False), "B8": ("=XOR(TRUE, A1>5, FALSE)", False), "B9": ("=XOR(TRUE, A2>5)", True)})
    assert c.unresolved == {}
    assert [got[("S", r, 2)] for r in range(1, 10)] == [
        ("s", None, "B"), ("e", None, "#N/A"), ("e", None, "#DIV/0!"), ("s", None, "most"), ("e", None, "#N/A"),
        ("n", 2.0, None), B(False), B(False), B(True)]


def test_ifna_and_the_is_functions():
    got, c = run({"A1": 0, "A2": "5", "A3": ("=1/A1", "#DIV/0!"), "A4": ("=NA()", "#N/A"), "A5": ('=IF(A1=0, "", 1)', ""),
                  "B1": ('=IFNA(A4, "gone")', "gone"), "B2": ('=IFNA(A3, "gone")', "#DIV/0!"),
                  "B3": ("=ISNA(A4)", True), "B4": ("=ISNA(A3)", False), "B5": ("=ISERROR(A3)", True),
                  "B6": ("=ISERR(A4)", False), "B7": ("=ISERR(A3)", True), "B8": ("=ISNUMBER(A2)", False),
                  "B9": ("=ISNUMBER(A1)", True), "B10": ("=ISTEXT(A5)", True), "B11": ("=ISBLANK(Z9)", True),
                  "B12": ("=ISBLANK(A5)", False), "B13": ("=ISLOGICAL(TRUE())", True), "B14": ("=ISNUMBER(A3)", False)})
    assert c.unresolved == {}
    assert [got[("S", r, 2)] for r in range(1, 15)] == [
        ("s", None, "gone"), ("e", None, "#DIV/0!"), B(True), B(False), B(True), B(False), B(True), B(False),
        B(True), B(True), B(True), B(False), B(True), B(False)]
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_logic.py -v`
Expected: FAIL. the cells are UNRESOLVED ("function IFS not supported yet" ...).

- [ ] **Step 3: Create `src/unexcel/fn/logic.py`**

```python
"""IFS, SWITCH, NOT, XOR, IFNA, ISNA, ISERROR, ISERR, ISNUMBER, ISTEXT, ISLOGICAL, ISBLANK, NA, TRUE, FALSE.

The IS functions look at a value without failing: an error in their argument is their answer, not their error.
"""
from __future__ import annotations

from unexcel.expr import NO_ERR, Expr, as_bool, compare, error, errs, pick, unify, unify_all
from unexcel.functions import arity, register
from unexcel.model import Unsupported
from unexcel.parser import RefNode
from unexcel.refs import resolve


def _chain(branches: list[tuple[Expr, Expr]], default: Expr | None) -> Expr:
    """CASE WHEN cond1 THEN value1 ... ELSE default (or #N/A): IFS and SWITCH."""
    values = unify_all([v for _, v in branches] + ([default] if default is not None else []))
    fallback = values[-1] if default is not None else None

    def case(attr: str) -> str:
        parts = [getattr(v, attr) for v in values]
        if len(set(parts)) == 1:
            return parts[0]
        whens = " ".join(f"WHEN {cnd.sql} THEN {getattr(v, attr)}" for (cnd, _), v in zip(branches, values))
        return f"(CASE {whens} ELSE {getattr(fallback, attr) if fallback else 'NULL'} END)"
    err_whens = []
    for (cnd, _), v in zip(branches, values):
        if cnd.err != NO_ERR:  # a condition that fails stops the chain, as in Excel
            err_whens.append(f"WHEN ({cnd.err}) IS NOT NULL THEN {cnd.err}")
        err_whens.append(f"WHEN {cnd.sql} THEN {v.err}")
    last = fallback.err if fallback else "'#N/A'"
    return Expr(case("sql"), values[0].type, f"(CASE {' '.join(err_whens)} ELSE {last} END)",
                case("txt"), case("kind"))


@register("IFS")
def ifs(c, args, sheet):
    if not args or len(args) % 2:
        raise Unsupported("IFS needs condition/value pairs")
    pairs = [(as_bool(c.expr(args[i], sheet)), c.expr(args[i + 1], sheet)) for i in range(0, len(args), 2)]
    return _chain(pairs, None)


@register("SWITCH")
def switch(c, args, sheet):
    if len(args) < 3:
        raise Unsupported("SWITCH needs a value and value/result pairs")
    x = c.expr(args[0], sheet)
    rest = args[1:]
    default = c.expr(rest[-1], sheet) if len(rest) % 2 else None
    pairs = [(compare(x, c.expr(rest[i], sheet), "="), c.expr(rest[i + 1], sheet)) for i in range(0, len(rest) - 1, 2)]
    return _chain(pairs, default)


@register("NOT")
def not_(c, args, sheet):
    arity("NOT", args, 1, 1)
    x = as_bool(c.expr(args[0], sheet))
    return Expr(f"(NOT {x.sql})", "bool", x.err)


@register("XOR")
def xor(c, args, sheet):
    if not args or any(isinstance(a, RefNode) and not resolve(a.text, sheet, c.book).is_single for a in args):
        raise Unsupported("XOR over ranges not supported yet")
    parts = [as_bool(c.expr(a, sheet)) for a in args]
    total = " + ".join(f"CAST({p.sql} AS INTEGER)" for p in parts)
    return Expr(f"(({total}) % 2 = 1)", "bool", errs(*(p.err for p in parts)))


@register("IFNA")
def ifna(c, args, sheet):
    arity("IFNA", args, 2, 2)
    value, fallback = unify(c.expr(args[0], sheet), c.expr(args[1], sheet))
    if value.err == NO_ERR:
        return value
    na = f"COALESCE(({value.err}) = '#N/A', FALSE)"
    return pick(f"(NOT {na})", value, fallback, f"(CASE WHEN {na} THEN {fallback.err} ELSE {value.err} END)")


def _is(name: str, test):
    def fn(c, args, sheet):
        arity(name, args, 1, 1)
        return Expr(test(c.expr(args[0], sheet)), "bool")
    register(name)(fn)


_is("ISNA", lambda x: f"COALESCE(({x.err}) = '#N/A', FALSE)")
_is("ISERROR", lambda x: f"(({x.err}) IS NOT NULL)")
_is("ISERR", lambda x: f"COALESCE(({x.err}) <> '#N/A', FALSE)")


def _kind_is(kind: str, typ: str):
    def test(x: Expr) -> str:
        if x.type == "any":
            return f"(({x.err}) IS NULL AND {x.kind} = '{kind}')"
        return f"(({x.err}) IS NULL)" if x.type == typ else "FALSE"
    return test


_is("ISNUMBER", _kind_is("n", "num"))
_is("ISTEXT", _kind_is("s", "txt"))
_is("ISLOGICAL", _kind_is("b", "bool"))


@register("ISBLANK")
def isblank(c, args, sheet):
    arity("ISBLANK", args, 1, 1)
    if not isinstance(args[0], RefNode):
        return Expr("FALSE", "bool")  # ISBLANK("") is FALSE: only an empty cell is blank
    area = resolve(args[0].text, sheet, c.book)
    cell = c.book.cells.get(area.key)
    empty = cell is None or (cell.formula is None and cell.kind == "blank")
    return Expr("TRUE" if empty else "FALSE", "bool")


@register("NA")
def na(c, args, sheet):
    arity("NA", args, 0, 0)
    return error("#N/A")


@register("TRUE")
def true(c, args, sheet):
    arity("TRUE", args, 0, 0)
    return Expr("TRUE", "bool")


@register("FALSE")
def false(c, args, sheet):
    arity("FALSE", args, 0, 0)
    return Expr("FALSE", "bool")
```

- [ ] **Step 4: Register it: replace `src/unexcel/fn/__init__.py`**

```python
"""The Excel function library. Importing this package registers every function in unexcel.functions.FUNCS."""
from unexcel.fn import aggregate, criteria, logic, lookup  # noqa: F401
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 6: Commit**

```bash
git add src/unexcel/fn tests/test_logic.py
git commit -m "feat: IFS SWITCH NOT XOR IFNA and the IS functions" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: Math: ABS, INT, SIGN, PI, SQRT, EXP, LN, LOG10, LOG, POWER, MOD, TRUNC, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, MROUND

**Files:**
- Create: `src/unexcel/fn/maths.py`
- Modify: `src/unexcel/fn/__init__.py`
- Test: `tests/test_maths.py`

**Interfaces:**
- Consumes: `expr.num/errs/ZERO`, `Compiler.binary` (POWER reuses `^`).
- Produces: `maths.xl15(sql)`: a value rounded to 15 significant digits. Rules: DuckDB throws on `ln(0)`, `sqrt(-1)` and `log10(0)`, so each domain is guarded with CASE and gets Excel's code (`#NUM!`, `#DIV/0!`); MOD takes the divisor's sign; ROUNDUP/ROUNDDOWN/TRUNC/CEILING/FLOOR/MROUND round at 15 significant digits first.

- [ ] **Step 1: Write the failing tests `tests/test_maths.py`**

```python
import math

from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute

CASES = [  # formula, Excel's answer (from Microsoft's function reference)
    ("=ABS(-2.5)", 2.5), ("=INT(-2.5)", -3), ("=SIGN(-0.1)", -1), ("=PI()", math.pi), ("=SQRT(16)", 4),
    ("=SQRT(-1)", "#NUM!"), ("=EXP(1)", math.e), ("=EXP(1000)", "#NUM!"), ("=LN(0)", "#NUM!"), ("=LOG10(1000)", 3),
    ("=LOG(8, 2)", 3), ("=LOG(100)", 2), ("=LOG(8, 1)", "#DIV/0!"), ("=POWER(2, 10)", 1024), ("=MOD(-7, 3)", 2),
    ("=MOD(7, -3)", -2), ("=MOD(5, 0)", "#DIV/0!"), ("=TRUNC(-8.9)", -8), ("=TRUNC(3.14159, 2)", 3.14),
    ("=ROUNDUP(0.1*3, 1)", 0.3), ("=ROUNDUP(-3.14159, 1)", -3.2), ("=ROUNDDOWN(3.14159, 3)", 3.141),
    ("=ROUNDUP(31415.9, -2)", 31500), ("=CEILING(2.5, 1)", 3), ("=CEILING(-2.5, -2)", -4), ("=CEILING(-2.5, 2)", -2),
    ("=CEILING(1.5, -1)", "#NUM!"), ("=FLOOR(-2.5, -2)", -2), ("=FLOOR(-2.5, 2)", -4), ("=FLOOR(1.5, 0)", "#DIV/0!"),
    ("=MROUND(10, 3)", 9), ("=MROUND(-10, -3)", -9), ("=MROUND(5, -2)", "#NUM!"), ("=MROUND(1.3, 0.2)", 1.4),
]


def test_math_functions_and_their_errors():
    book = book_of({"S": {f"A{i}": (f, v) for i, (f, v) in enumerate(CASES, 1)}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert compiled.unresolved == {}
    for i, (formula, want) in enumerate(CASES, 1):
        kind, n, t = got[("S", i, 1)]
        if isinstance(want, str):
            assert (kind, t) == ("e", want), formula
        else:
            assert kind == "n" and math.isclose(n, want, rel_tol=1e-12), (formula, n)
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_maths.py -v`
Expected: FAIL. the cells are UNRESOLVED ("function ABS not supported yet" ...).

- [ ] **Step 3: Create `src/unexcel/fn/maths.py`**

```python
"""ABS, INT, SIGN, PI, SQRT, EXP, LN, LOG10, LOG, POWER, MOD, TRUNC, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, MROUND.

DuckDB throws on ln(0) or sqrt(-1), and one throw would stop the whole batch, so every function with a
domain is written as CASE WHEN <ok> THEN f(x) END plus the Excel error for the rest (a CASE never runs the
branch it skips, verified on DuckDB 1.5.6). Rounding to a digit first rounds to 15 significant digits, as
Excel does, so ROUNDUP(0.1*3, 1) is 0.3, not 0.4.
"""
from __future__ import annotations

from unexcel.expr import ZERO, Expr, errs, num
from unexcel.functions import arity, register
from unexcel.parser import Binary


def _guarded(value: str, bad: str, code: str) -> tuple[str, str]:
    return f"(CASE WHEN NOT ({bad}) THEN {value} END)", f"(CASE WHEN {bad} THEN '{code}' END)"


def _one(name: str, value, bad=None, code: str = "#NUM!"):
    def fn(c, args, sheet):
        arity(name, args, 1, 1)
        x = num(c.expr(args[0], sheet))
        if bad is None:
            return Expr(value(x.sql), "num", x.err)
        sql, err = _guarded(value(x.sql), bad(x.sql), code)
        return Expr(sql, "num", errs(x.err, err))
    register(name)(fn)


_one("ABS", lambda x: f"abs({x})")
_one("INT", lambda x: f"floor({x})")
_one("SIGN", lambda x: f"CAST(sign({x}) AS DOUBLE)")
_one("SQRT", lambda x: f"sqrt({x})", lambda x: f"{x} < 0")
_one("EXP", lambda x: f"exp({x})", lambda x: f"{x} > 709.782712893384")
_one("LN", lambda x: f"ln({x})", lambda x: f"{x} <= 0")
_one("LOG10", lambda x: f"log10({x})", lambda x: f"{x} <= 0")


@register("PI")
def pi(c, args, sheet):
    arity("PI", args, 0, 0)
    return Expr("pi()", "num")


@register("LOG")
def log(c, args, sheet):
    arity("LOG", args, 1, 2)
    x = num(c.expr(args[0], sheet))
    base = num(c.expr(args[1], sheet)) if len(args) == 2 and args[1] is not None else Expr("CAST(10 AS DOUBLE)", "num")
    bad = f"(CASE WHEN {x.sql} <= 0 OR {base.sql} <= 0 THEN '#NUM!' WHEN {base.sql} = 1 THEN '#DIV/0!' END)"
    return Expr(f"(CASE WHEN {x.sql} > 0 AND {base.sql} > 0 AND {base.sql} <> 1 THEN ln({x.sql}) / ln({base.sql}) END)",
                "num", errs(x.err, base.err, bad))


@register("POWER")
def power(c, args, sheet):
    arity("POWER", args, 2, 2)
    return c.binary(Binary("^", args[0], args[1]), sheet)


def _two(c, args, sheet, name: str, default=None) -> tuple[Expr, Expr]:
    arity(name, args, 1 if default is not None else 2, 2)
    x = num(c.expr(args[0], sheet))
    y = num(c.expr(args[1], sheet)) if len(args) == 2 and args[1] is not None else (default or ZERO)
    return x, y


def xl15(v: str) -> str:
    """v rounded to 15 significant digits, the precision Excel works at."""
    return f"(CASE WHEN {v} = 0 THEN 0 WHEN NOT isfinite({v}) THEN {v} ELSE round({v}, 14 - CAST(floor(log10(abs({v}))) AS INTEGER)) END)"


@register("MOD")
def mod(c, args, sheet):
    x, y = _two(c, args, sheet, "MOD")
    sql, err = _guarded(f"({x.sql} - {y.sql} * floor({xl15(f'{x.sql} / {y.sql}')}))", f"{y.sql} = 0", "#DIV/0!")
    return Expr(sql, "num", errs(x.err, y.err, err))  # Excel: the result takes the divisor's sign


def _digits(name: str, rnd: str):
    def fn(c, args, sheet):
        x, d = _two(c, args, sheet, name, default=ZERO if name == "TRUNC" else None)
        p = f"power(10, trunc({d.sql}))"
        # shortcut: scaling outside DOUBLE gives #NUM!, revisit with Excel boundary fixtures.
        scaled = f"abs({x.sql}) * {p}"
        ok = f"(isfinite({p}) AND {p} > 0 AND isfinite({scaled}))"
        sql = f"(CASE WHEN {d.sql} > 308 THEN {x.sql} WHEN {ok} THEN sign({x.sql}) * {rnd}({xl15(scaled)}) / {p} END)"
        bad = f"(CASE WHEN {d.sql} <= 308 AND NOT {ok} THEN '#NUM!' END)"
        return Expr(sql, "num", errs(x.err, d.err, bad))
    register(name)(fn)


_digits("ROUNDUP", "ceil")
_digits("ROUNDDOWN", "floor")
_digits("TRUNC", "floor")


def _multiple(name: str, rnd: str, zero_code: str | None):
    def fn(c, args, sheet):
        x, m = _two(c, args, sheet, name)
        bad = f"({x.sql} > 0 AND {m.sql} < 0)" if name != "MROUND" else f"({x.sql} * {m.sql} < 0)"
        if zero_code:
            bad = f"({bad} OR {m.sql} = 0)"
        value = f"(CASE WHEN {m.sql} = 0 THEN 0 ELSE {rnd}({xl15(f'{x.sql} / {m.sql}')}) * {m.sql} END)"
        sql, _ = _guarded(value, bad, "#NUM!")
        code = (f"(CASE WHEN {m.sql} = 0 THEN '{zero_code}' WHEN {bad} THEN '#NUM!' END)" if zero_code
                else f"(CASE WHEN {bad} THEN '#NUM!' END)")
        return Expr(sql, "num", errs(x.err, m.err, code))
    register(name)(fn)


_multiple("CEILING", "ceil", None)
_multiple("FLOOR", "floor", "#DIV/0!")
_multiple("MROUND", "round", None)
```

- [ ] **Step 4: Register it: replace `src/unexcel/fn/__init__.py`**

```python
"""The Excel function library. Importing this package registers every function in unexcel.functions.FUNCS."""
from unexcel.fn import aggregate, criteria, logic, lookup, maths  # noqa: F401
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 6: Commit**

```bash
git add src/unexcel/fn tests/test_maths.py
git commit -m "feat: math functions with Excel's error codes" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: Text: CONCATENATE, CONCAT, TEXTJOIN, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, PROPER, SUBSTITUTE, REPLACE, FIND, SEARCH, REPT, VALUE, TEXT, EXACT, CHAR, CLEAN

**Files:**
- Create: `src/unexcel/fn/text.py`
- Modify: `src/unexcel/fn/__init__.py`
- Test: `tests/test_text.py`

**Interfaces:**
- Consumes: `expr.txt/num/num_text/errs/where_area`, `Compiler.ready/may_err/expr`.
- Produces: `text.FORMATS: dict[str, Callable[[sql], sql]]` (TEXT's formats; Task 8 adds date formats; a format function with `is_date = True` is refused in 1904-system workbooks), `text.NUMBER_FORMATS`. Rules: TRIM also squeezes inner runs of spaces; SEARCH ignores case and honours wildcards, FIND does neither; positions are 1-based characters; CHAR uses Windows-1252; TEXT rounds half away from zero before formatting; a TEXT format outside the table is UNRESOLVED.

- [ ] **Step 1: Write the failing tests `tests/test_text.py`**

```python
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute

CASES = [  # formula, Excel's answer (Microsoft's function reference)
    ('=CONCATENATE("a", 1, TRUE)', "a1TRUE"), ("=CONCAT(S!B1:B3)", "x2.5TRUE"), ('=TEXTJOIN("-", TRUE, S!C1:C3, "z")', "p-q-z"),
    ('=LEFT("Excel", 2)', "Ex"), ('=LEFT("Excel")', "E"), ('=RIGHT("Excel", 3)', "cel"), ('=MID("Excel", 2, 3)', "xce"),
    ('=MID("Excel", 0, 1)', "#VALUE!"), ('=LEN("héllo")', 5), ('=TRIM("  a   b  ")', "a b"), ('=UPPER("aB")', "AB"),
    ('=PROPER("hELLO o\'neil 2nd")', "Hello O'Neil 2Nd"), ('=SUBSTITUTE("a-b-c", "-", "+")', "a+b+c"),
    ('=SUBSTITUTE("a-b-c", "-", "+", 2)', "a-b+c"), ('=SUBSTITUTE("a-b", "-", "+", 5)', "a-b"),
    ('=REPLACE("abcdef", 2, 3, "XY")', "aXYef"), ('=FIND("c", "abcabc")', 3), ('=FIND("C", "abcabc")', "#VALUE!"),
    ('=FIND("c", "abcabc", 4)', 6), ('=SEARCH("C", "abcabc")', 3), ('=SEARCH("b?c", "xaBxCd")', 3),
    ('=SEARCH("~?", "a?b")', 2), ('=REPT("ab", 3)', "ababab"), ('=VALUE(" 12.5 ")', 12.5), ('=VALUE("x")', "#VALUE!"),
    ('=EXACT("a", "A")', False), ("=CHAR(65)", "A"), ("=CHAR(128)", "€"), ("=CHAR(0)", "#VALUE!"),
    ('=CLEAN("a"&CHAR(9)&"b")', "ab"), ('=TEXT(1234.567, "#,##0.00")', "1,234.57"), ('=TEXT(2.675, "0.00")', "2.68"),
    ('=TEXT(0.256, "0.0%")', "25.6%"), ('=TEXT(-1234.5, "#,##0")', "-1,235"), ('=TEXT("abc", "0.00")', "abc"),
]


def test_text_functions():
    book = book_of({"S": {"B1": "x", "B2": 2.5, "B3": True, "C1": "p", "C3": "q"},
                    "F": {f"A{i}": (f, v) for i, (f, v) in enumerate(CASES, 1)}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert compiled.unresolved == {}
    for i, (formula, want) in enumerate(CASES, 1):
        kind, n, t = got[("F", i, 1)]
        if isinstance(want, bool):
            assert (kind, n) == ("b", float(want)), formula
        elif isinstance(want, (int, float)):
            assert (kind, n) == ("n", float(want)), formula
        elif want.startswith("#"):
            assert (kind, t) == ("e", want), formula
        else:
            assert (kind, t) == ("s", want), (formula, t)


def test_unsupported_text_forms_say_why():
    book = book_of({"S": {"A1": "a", "A3": "c", "B1": ('=TEXTJOIN(",", FALSE, A1:A3)', "a,,c"),
                          "B2": ('=TEXT(1, "0.0E+00")', "1.0E+00")}})
    unresolved = compile_book(book).unresolved
    assert unresolved[("S", 1, 2)] == "TEXTJOIN keeping the empty cells of a range not supported yet"
    assert unresolved[("S", 2, 2)] == "TEXT format '0.0E+00' not supported yet"
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_text.py -v`
Expected: FAIL. the cells are UNRESOLVED ("function CONCATENATE not supported yet" ...).

- [ ] **Step 3: Create `src/unexcel/fn/text.py`**

```python
"""CONCATENATE, CONCAT, TEXTJOIN, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, PROPER, SUBSTITUTE, REPLACE,
FIND, SEARCH, REPT, VALUE, TEXT, EXACT, CHAR, CLEAN.

Numbers become text the way Excel writes them (expr.num_text: 15 significant digits). Positions are 1-based
and counted in characters, as in Excel. TEXT supports the number formats in NUMBER_FORMATS; dates.py adds
date formats to FORMATS.
"""
from __future__ import annotations

from unexcel.expr import BLANK, CELLS, NO_ERR, Expr, errs, num, num_text, txt, where_area
from unexcel.functions import arity, register
from unexcel.model import Unsupported
from unexcel.parser import Bool, RefNode, Str
from unexcel.refs import resolve


def _t(c, node, sheet: str) -> Expr:
    return txt(c.expr(node, sheet) if node is not None else BLANK)


def _n(c, node, sheet: str) -> Expr:
    return num(c.expr(node, sheet) if node is not None else BLANK)


def _bad(cond: str) -> str:
    return f"(CASE WHEN {cond} THEN '#VALUE!' END)"


CELL_TEXT = ("(CASE kind WHEN 's' THEN COALESCE(txt, '') WHEN 'b' THEN (CASE WHEN num <> 0 THEN 'TRUE' ELSE 'FALSE' END) "
             f"ELSE {num_text('num')} END)")


def _range_texts(c, node, sheet: str) -> tuple[str, str, int]:
    """(SQL list of the range's cell texts in row order, first error, number of empty cells in the range)."""
    area = resolve(node.text, sheet, c.book)
    keys = c.ready(area)
    where = f"FROM {CELLS} WHERE {where_area(area)}"
    err = NO_ERR
    if any(c.may_err(k) for k in keys):
        err = f"(SELECT txt {where} AND kind = 'e' ORDER BY row_num, col_num LIMIT 1)"
    blanks = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1) - len(keys)
    return (f"(SELECT COALESCE(list({CELL_TEXT} ORDER BY row_num, col_num), []) {where} AND kind <> 'e')", err,
            blanks)


@register("CONCATENATE")
def concatenate(c, args, sheet):
    if not args:
        raise Unsupported("CONCATENATE needs arguments")
    parts = [_t(c, a, sheet) for a in args]
    return Expr("(" + " || ".join(p.sql for p in parts) + ")", "txt", errs(*(p.err for p in parts)))


def _pieces(c, args, sheet: str, keep_blanks: bool) -> tuple[list[str], list[str]]:
    lists, err = [], []
    for a in args:
        if isinstance(a, RefNode) and not resolve(a.text, sheet, c.book).is_single:
            items, e, blanks = _range_texts(c, a, sheet)
            if keep_blanks and blanks:
                raise Unsupported("TEXTJOIN keeping the empty cells of a range not supported yet")
            lists.append(items)
            err.append(e)
        else:
            p = _t(c, a, sheet)
            lists.append(f"[{p.sql}]")
            err.append(p.err)
    return lists, err


@register("CONCAT")
def concat(c, args, sheet):
    if not args:
        raise Unsupported("CONCAT needs arguments")
    lists, err = _pieces(c, args, sheet, keep_blanks=False)
    return Expr(f"array_to_string(list_concat({', '.join(lists)}), '')" if len(lists) > 1
                else f"array_to_string({lists[0]}, '')", "txt", errs(*err))


@register("TEXTJOIN")
def textjoin(c, args, sheet):
    if len(args) < 3 or not isinstance(args[1], Bool):
        raise Unsupported("TEXTJOIN needs a delimiter, TRUE/FALSE written in the formula, and values")
    delim = _t(c, args[0], sheet)
    lists, err = _pieces(c, args[2:], sheet, keep_blanks=not args[1].value)
    items = f"list_concat({', '.join(lists)})" if len(lists) > 1 else lists[0]
    if args[1].value:
        items = f"list_filter({items}, x -> x <> '')"
    return Expr(f"array_to_string({items}, {delim.sql})", "txt", errs(delim.err, *err))


def _count_arg(c, args, sheet: str, i: int) -> Expr:
    if len(args) > i and args[i] is not None:
        return _n(c, args[i], sheet)
    return Expr("CAST(1 AS DOUBLE)", "num")


def _side(name: str, fn: str):
    def impl(c, args, sheet):
        arity(name, args, 1, 2)
        t, n = _t(c, args[0], sheet), _count_arg(c, args, sheet, 1)
        return Expr(f"{fn}({t.sql}, CAST(trunc(least(greatest({n.sql}, 0), length({t.sql}))) AS BIGINT))", "txt",
                    errs(t.err, n.err, _bad(f"{n.sql} < 0")))
    register(name)(impl)


_side("LEFT", "left")
_side("RIGHT", "right")


@register("MID")
def mid(c, args, sheet):
    arity("MID", args, 3, 3)
    t, start, n = _t(c, args[0], sheet), _n(c, args[1], sheet), _n(c, args[2], sheet)
    return Expr(f"substr({t.sql}, CAST(trunc(least(greatest({start.sql}, 1), length({t.sql}) + 1)) AS BIGINT), "
                f"CAST(trunc(least(greatest({n.sql}, 0), length({t.sql}))) AS BIGINT))", "txt",
                errs(t.err, start.err, n.err, _bad(f"{start.sql} < 1 OR {n.sql} < 0")))


def _text1(name: str, sql, typ: str = "txt"):
    def impl(c, args, sheet):
        arity(name, args, 1, 1)
        t = _t(c, args[0], sheet)
        return Expr(sql(t.sql), typ, t.err)
    register(name)(impl)


_text1("LEN", lambda t: f"CAST(length({t}) AS DOUBLE)", "num")
_text1("TRIM", lambda t: f"regexp_replace(trim({t}, ' '), ' +', ' ', 'g')")  # Excel also squeezes inner spaces
_text1("UPPER", lambda t: f"upper({t})")
_text1("LOWER", lambda t: f"lower({t})")
_text1("PROPER", lambda t: (f"array_to_string(list_transform(range(length({t})), i -> CASE WHEN i = 0 OR NOT "
                            f"regexp_matches({t}[i], '[[:alpha:]]') THEN upper({t}[i + 1]) ELSE lower({t}[i + 1]) END), '')"))
_text1("CLEAN", lambda t: f"regexp_replace({t}, '[\\x00-\\x1f]', '', 'g')")


@register("SUBSTITUTE")
def substitute(c, args, sheet):
    arity("SUBSTITUTE", args, 3, 4)
    t, old, new = (_t(c, a, sheet) for a in args[:3])
    if len(args) == 3:
        return Expr(f"(CASE WHEN {old.sql} = '' THEN {t.sql} ELSE replace({t.sql}, {old.sql}, {new.sql}) END)", "txt",
                    errs(t.err, old.err, new.err))
    n = _n(c, args[3], sheet)
    parts = f"string_split({t.sql}, {old.sql})"
    k = f"CAST(trunc(least(greatest({n.sql}, 0), len({parts}))) AS BIGINT)"
    return Expr(f"(CASE WHEN {old.sql} = '' OR len({parts}) <= {k} THEN {t.sql} ELSE "
                f"array_to_string({parts}[1:{k}], {old.sql}) || {new.sql} || array_to_string({parts}[{k} + 1:], {old.sql}) END)",
                "txt", errs(t.err, old.err, new.err, n.err, _bad(f"{n.sql} < 1")))


@register("REPLACE")
def replace_(c, args, sheet):
    arity("REPLACE", args, 4, 4)
    t, start, n, new = _t(c, args[0], sheet), _n(c, args[1], sheet), _n(c, args[2], sheet), _t(c, args[3], sheet)
    s = f"CAST(trunc(least(greatest({start.sql}, 1), length({t.sql}) + 1)) AS BIGINT)"
    k = f"CAST(trunc(least(greatest({n.sql}, 0), length({t.sql}))) AS BIGINT)"
    return Expr(f"(substr({t.sql}, 1, {s} - 1) || {new.sql} || substr({t.sql}, {s} + {k}))", "txt",
                errs(t.err, start.err, n.err, new.err, _bad(f"{start.sql} < 1 OR {n.sql} < 0")))


def _find(name: str, sensitive: bool):
    def impl(c, args, sheet):
        arity(name, args, 2, 3)
        what, within = _t(c, args[0], sheet), _t(c, args[1], sheet)
        start = _count_arg(c, args, sheet, 2)
        s = f"TRY_CAST(trunc({start.sql}) AS BIGINT)"
        rest = f"substr({within.sql}, {s})"
        if sensitive:
            pos = f"strpos({rest}, {what.sql})"
            found = f"{pos} > 0"
        else:  # SEARCH: ignores case, * and ? are wildcards, ~ escapes them
            pat = f"regexp_escape(lower({what.sql}))"  # escaped: * -> \\*, ? -> \\?, ~ -> \\~
            for a, b in (("'\\~\\*'", "chr(1)"), ("'\\~\\?'", "chr(2)"), ("'\\*'", "'.*'"), ("'\\?'", "'.'"),
                         ("chr(1)", "'\\*'"), ("chr(2)", "'\\?'")):
                pat = f"replace({pat}, {a}, {b})"
            found = f"regexp_matches(lower({rest}), {pat})"
            pos = f"(length(regexp_extract(lower({rest}), '(?s)^(.*?)(' || {pat} || ')', 1)) + 1)"
        bad = f"{s} IS NULL OR {s} < 1 OR {s} > length({within.sql}) + 1 OR NOT ({found})"
        return Expr(f"(CASE WHEN NOT ({bad}) THEN CAST({pos} + {s} - 1 AS DOUBLE) END)", "num",
                    errs(what.err, within.err, start.err, _bad(bad)))
    register(name)(impl)


_find("FIND", sensitive=True)
_find("SEARCH", sensitive=False)


@register("REPT")
def rept(c, args, sheet):
    arity("REPT", args, 2, 2)
    t, n = _t(c, args[0], sheet), _n(c, args[1], sheet)
    bad = f"({n.sql} < 0 OR length({t.sql}) * trunc({n.sql}) > 32767)"
    sql = (f"(CASE WHEN NOT {bad} THEN CASE WHEN {t.sql} = '' THEN '' "
           f"ELSE repeat({t.sql}, CAST(trunc({n.sql}) AS BIGINT)) END END)")
    return Expr(sql, "txt", errs(t.err, n.err, _bad(bad)))


@register("VALUE")
def value(c, args, sheet):
    arity("VALUE", args, 1, 1)
    return num(c.expr(args[0], sheet))


@register("EXACT")
def exact(c, args, sheet):
    arity("EXACT", args, 2, 2)
    a, b = _t(c, args[0], sheet), _t(c, args[1], sheet)
    return Expr(f"({a.sql} = {b.sql})", "bool", errs(a.err, b.err))


CP1252 = bytes(range(128, 160)).decode("cp1252", errors="replace")


@register("CHAR")
def char(c, args, sheet):
    arity("CHAR", args, 1, 1)
    n = _n(c, args[0], sheet)
    k = f"TRY_CAST(trunc({n.sql}) AS INTEGER)"
    table = "'" + CP1252.replace("'", "''") + "'"
    return Expr(f"(CASE WHEN {k} BETWEEN 128 AND 159 THEN substr({table}, {k} - 127, 1) "
                f"WHEN {k} BETWEEN 1 AND 255 THEN chr({k}) END)", "txt",
                errs(n.err, _bad(f"COALESCE({k} < 1 OR {k} > 255, TRUE)")))  # Windows-1252, the character set of Excel for Windows


def _fixed(decimals: int, grouped: bool, percent: bool):
    def sql(x: str) -> str:
        v = f"round({x} * 100, {decimals})" if percent else f"round({x}, {decimals})"
        body = f"format('{{:,.{decimals}f}}', {v})" if grouped else f"printf('%.{decimals}f', {v})"
        return f"({body} || '%')" if percent else body
    return sql


NUMBER_FORMATS = {"General": num_text}
for d in range(5):
    tail = "." + "0" * d if d else ""
    NUMBER_FORMATS[f"0{tail}"] = _fixed(d, False, False)
    NUMBER_FORMATS[f"#,##0{tail}"] = _fixed(d, True, False)
    NUMBER_FORMATS[f"0{tail}%"] = _fixed(d, False, True)
FORMATS = dict(NUMBER_FORMATS)  # dates.py adds the date formats


@register("TEXT")
def text(c, args, sheet):
    arity("TEXT", args, 2, 2)
    if not isinstance(args[1], Str):
        raise Unsupported("TEXT format must be text written in the formula")
    fmt = FORMATS.get(args[1].value) or FORMATS.get(args[1].value.lower())
    if fmt is None:
        raise Unsupported(f"TEXT format {args[1].value!r} not supported yet")
    if getattr(fmt, "is_date", False) and c.book.date1904:
        raise Unsupported("1904 date system not supported yet")
    x = c.expr(args[0], sheet)
    if x.type == "txt":  # TEXT("abc", "0.00") passes text through
        return x
    x = num(x)
    return Expr(fmt(x.sql), "txt", x.err)
```

- [ ] **Step 4: Register it: replace `src/unexcel/fn/__init__.py`**

```python
"""The Excel function library. Importing this package registers every function in unexcel.functions.FUNCS."""
from unexcel.fn import aggregate, criteria, logic, lookup, maths, text  # noqa: F401
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 6: Commit**

```bash
git add src/unexcel/fn tests/test_text.py
git commit -m "feat: text functions and TEXT number formats" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 8: Dates: DATE, YEAR, MONTH, DAY, WEEKDAY, DAYS, EDATE, EOMONTH, DAYS360, DATEDIF, and TEXT's date formats

**Files:**
- Create: `src/unexcel/fn/dates.py`
- Modify: `src/unexcel/model.py`, `src/unexcel/reader.py`, `src/unexcel/compute.py`, `src/unexcel/fn/__init__.py`
- Test: `tests/test_dates.py`

**Interfaces:**
- Consumes: `lookup.int_literal`, `text.FORMATS`, `expr.num/errs`.
- Produces: `Book.date1904: bool = False` (set by the reader from the workbook's date system); `dates.to_date(sql)`, `dates.from_date(sql)`. Rules: serial 1 = 1900-01-01, Excel's fake 1900-02-29 is serial 60, serial 0 is "1900-01-00"; outside 0..2958465 is `#NUM!`; DATE rolls months and days over and adds 1900 to years below 1900; DAYS360 follows Microsoft's documented U.S. method (FALSE) and European method (TRUE); 1904-system workbooks are UNRESOLVED.

- [ ] **Step 1: Write the failing tests `tests/test_dates.py`**

```python
from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute

CASES = [  # formula, Excel's answer (serial numbers: 46023 = 2026-01-01)
    ("=DATE(2026, 1, 1)", 46023), ("=DATE(2026, 14, 31)", 46449), ("=DATE(26, 1, 1)", 9498), ("=DATE(2026, 2, 29)", 46082),
    ("=DATE(-1, 1, 1)", "#NUM!"), ("=YEAR(46023)", 2026), ("=MONTH(46023.75)", 1), ("=DAY(46081)", 28),
    ("=MONTH(60)", 2), ("=DAY(60)", 29), ("=DAY(0)", 0), ("=YEAR(-1)", "#NUM!"), ("=WEEKDAY(46023)", 5),
    ("=WEEKDAY(46023, 2)", 4), ("=WEEKDAY(46023, 3)", 3), ("=WEEKDAY(1)", 1), ("=DAYS(46081, 46023)", 58),
    ("=EDATE(46053, 1)", 46081), ("=EOMONTH(46023, 1)", 46081), ("=EOMONTH(46023, -1)", 46022),
    ("=DAYS360(DATE(2011,1,30), DATE(2011,12,31))", 330), ("=DAYS360(DATE(2011,1,31), DATE(2011,3,31), TRUE)", 60),
    ("=DAYS360(DATE(2026,1,15), DATE(2026,3,15))", 60), ('=DATEDIF(DATE(2020,5,10), DATE(2026,5,9), "Y")', 5),
    ('=DATEDIF(DATE(2026,1,31), DATE(2026,2,28), "M")', 0), ('=DATEDIF(46023, 46081, "D")', 58),
    ('=DATEDIF(46081, 46023, "D")', "#NUM!"), ('=TEXT(46023, "yyyy-mm-dd")', "2026-01-01"),
    ('=TEXT(46023, "dddd d mmmm")', None), ('=TEXT(46087, "mmm yyyy")', "Mar 2026"), ('=TEXT(46087, "dd/mm/yyyy")', "06/03/2026"),
]


def test_date_functions():
    cases = [(f, v) for f, v in CASES if v is not None]
    book = book_of({"S": {f"A{i}": (f, v) for i, (f, v) in enumerate(cases, 1)}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert compiled.unresolved == {}
    for i, (formula, want) in enumerate(cases, 1):
        kind, n, t = got[("S", i, 1)]
        if isinstance(want, str) and want.startswith("#"):
            assert (kind, t) == ("e", want), formula
        elif isinstance(want, str):
            assert (kind, t) == ("s", want), (formula, t)
        else:
            assert (kind, n) == ("n", float(want)), (formula, n)


def test_1904_workbooks_and_odd_forms_say_why():
    book = book_of({"S": {"A1": ("=YEAR(1)", 1904), "A2": ('=TEXT(1, "yyyy")', "1904")}})
    book.date1904 = True
    assert set(compile_book(book).unresolved.values()) == {"1904 date system not supported yet"}
    other = book_of({"S": {"A1": ('=DATEDIF(1, 2, "MD")', 1), "A2": ("=WEEKDAY(1, 11)", 7)}})
    assert compile_book(other).unresolved == {
        ("S", 1, 1): 'DATEDIF unit must be "Y", "M" or "D" written in the formula',
        ("S", 2, 1): "WEEKDAY return type must be 1, 2 or 3"}
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_dates.py -v`
Expected: FAIL. The cells are UNRESOLVED ("function DATE not supported yet" ...).

- [ ] **Step 3: Record the date system.** In `src/unexcel/model.py`, inside `class Book`, after the `max_row` field add
```python
    date1904: bool = False  # the workbook counts dates from 1904 (old Mac Excel)
```
In `src/unexcel/reader.py` replace
```python
from openpyxl.utils.datetime import to_excel
```
with
```python
from openpyxl.utils.datetime import CALENDAR_MAC_1904, to_excel
```
In `src/unexcel/reader.py` replace
```python
    return Book(cells, [ws.title for ws in fwb.worksheets], names, max_row)
```
with
```python
    return Book(cells, [ws.title for ws in fwb.worksheets], names, max_row, fwb.epoch == CALENDAR_MAC_1904)
```
In `src/unexcel/compute.py` replace
```python
    return Book(cells, book.sheets, book.names, book.max_row), computed, ""
```
with
```python
    return replace(book, cells=cells), computed, ""  # keeps date1904
```

- [ ] **Step 4: Create `src/unexcel/fn/dates.py`**

```python
"""DATE, YEAR, MONTH, DAY, WEEKDAY, DAYS, EDATE, EOMONTH, DAYS360, DATEDIF, and TEXT's date formats.

Excel stores a date as a serial number: 1 = 1900-01-01, and it counts a 29 February 1900 that never existed
(serial 60), so serials from 61 on are days since 1899-12-30. Serial 0 is "1900-01-00". Workbooks using the
1904 date system are not supported yet.
"""
from __future__ import annotations

from unexcel.expr import Expr, errs, num
from unexcel.fn.lookup import int_literal
from unexcel.fn.text import FORMATS
from unexcel.functions import arity, register
from unexcel.model import Unsupported
from unexcel.parser import Bool, Str

MAX_SERIAL = 2958465  # 9999-12-31


def _check(c) -> None:
    if c.book.date1904:
        raise Unsupported("1904 date system not supported yet")


def _serial(c, node, sheet: str) -> Expr:
    _check(c)
    s = num(c.expr(node, sheet))
    return Expr(f"floor({s.sql})", "num", errs(s.err, f"(CASE WHEN {s.sql} < 0 OR {s.sql} > {MAX_SERIAL} THEN '#NUM!' END)"))


def to_date(s: str) -> str:
    """A whole serial (61 and up, or 1..59) as a DATE."""
    return (f"(CASE WHEN {s} >= 61 THEN DATE '1899-12-30' + CAST({s} AS INTEGER) "
            f"ELSE DATE '1899-12-31' + CAST({s} AS INTEGER) END)")


def from_date(d: str) -> str:
    return (f"CAST(CASE WHEN {d} >= DATE '1900-03-01' THEN {d} - DATE '1899-12-30' "
            f"ELSE {d} - DATE '1899-12-31' END AS DOUBLE)")


def _ok(s: Expr) -> str:
    return f"(({s.err}) IS NULL)"


def _part(name: str, fn: str, at0: int, at60: int):
    def impl(c, args, sheet):
        arity(name, args, 1, 1)
        s = _serial(c, args[0], sheet)
        return Expr(f"(CASE WHEN NOT {_ok(s)} THEN NULL WHEN {s.sql} = 0 THEN {at0} WHEN {s.sql} = 60 THEN {at60} "
                    f"ELSE CAST({fn}({to_date(s.sql)}) AS DOUBLE) END)", "num", s.err)
    register(name)(impl)


_part("YEAR", "year", 1900, 1900)
_part("MONTH", "month", 1, 2)
_part("DAY", "day", 0, 29)


@register("DATE")
def date(c, args, sheet):
    arity("DATE", args, 3, 3)
    _check(c)
    y, m, d = (num(c.expr(a, sheet)) for a in args)
    year = f"(CASE WHEN trunc({y.sql}) < 1900 THEN trunc({y.sql}) + 1900 ELSE trunc({y.sql}) END)"
    ok = (f"(trunc({y.sql}) BETWEEN 0 AND 9999 AND abs({m.sql}) < 120000 AND abs({d.sql}) < 4000000)")
    day = (f"CAST(make_date(CAST({year} AS INTEGER), 1, 1) + to_months(CAST(trunc({m.sql}) - 1 AS INTEGER)) "
           f"+ to_days(CAST(trunc({d.sql}) - 1 AS INTEGER)) AS DATE)")
    serial = f"(CASE WHEN {ok} THEN {from_date(day)} END)"
    return Expr(serial, "num", errs(y.err, m.err, d.err,
                                    f"(CASE WHEN NOT {ok} OR {serial} < 0 OR {serial} > {MAX_SERIAL} THEN '#NUM!' END)"))


@register("WEEKDAY")
def weekday(c, args, sheet):
    arity("WEEKDAY", args, 1, 2)
    s = _serial(c, args[0], sheet)
    kind = 1 if len(args) == 1 or args[1] is None else int_literal(args[1])
    sunday0 = f"(CAST({s.sql} AS BIGINT) + 6) % 7"  # serial 1 was a Sunday in Excel's calendar
    sql = {1: f"({sunday0} + 1)", 2: f"(({sunday0} + 6) % 7 + 1)", 3: f"(({sunday0} + 6) % 7)"}.get(kind)
    if sql is None:
        raise Unsupported("WEEKDAY return type must be 1, 2 or 3")
    return Expr(f"(CASE WHEN {_ok(s)} THEN CAST({sql} AS DOUBLE) END)", "num", s.err)


@register("DAYS")
def days(c, args, sheet):
    arity("DAYS", args, 2, 2)
    end, start = _serial(c, args[0], sheet), _serial(c, args[1], sheet)
    return Expr(f"({end.sql} - {start.sql})", "num", errs(end.err, start.err))


def _shift(name: str, last: bool):
    def impl(c, args, sheet):
        arity(name, args, 2, 2)
        s, m = _serial(c, args[0], sheet), num(c.expr(args[1], sheet))
        d = f"CAST({to_date(s.sql)} + to_months(CAST(trunc({m.sql}) AS INTEGER)) AS DATE)"
        if last:
            d = f"last_day({d})"
        ok = f"({_ok(s)} AND ({m.err}) IS NULL AND trunc({m.sql}) BETWEEN -120000 AND 120000)"
        serial = f"(CASE WHEN {ok} THEN {from_date(d)} END)"
        bad = f"(CASE WHEN NOT {ok} OR {serial} NOT BETWEEN 0 AND {MAX_SERIAL} THEN '#NUM!' END)"
        return Expr(serial, "num", errs(s.err, m.err, bad))
    register(name)(impl)


_shift("EDATE", last=False)
_shift("EOMONTH", last=True)


@register("DAYS360")
def days360(c, args, sheet):
    arity("DAYS360", args, 2, 3)
    if len(args) == 3 and not isinstance(args[2], Bool):
        raise Unsupported("DAYS360 method must be TRUE/FALSE written in the formula")
    european = len(args) == 3 and args[2].value
    a, b = _serial(c, args[0], sheet), _serial(c, args[1], sheet)
    d1, d2 = to_date(a.sql), to_date(b.sql)
    y1, m1, y2, m2 = f"year({d1})", f"month({d1})", f"year({d2})", f"month({d2})"
    if european:
        day1, day2, m2 = f"least(day({d1}), 30)", f"least(day({d2}), 30)", m2
    else:  # U.S. (NASD), as Microsoft documents it for DAYS360
        day1 = f"(CASE WHEN {d1} = last_day({d1}) THEN 30 ELSE day({d1}) END)"
        end_last = f"{d2} = last_day({d2})"
        day2 = f"(CASE WHEN {end_last} AND {day1} < 30 THEN 1 WHEN {end_last} THEN 30 ELSE day({d2}) END)"
        m2 = f"({m2} + CASE WHEN {end_last} AND {day1} < 30 THEN 1 ELSE 0 END)"
    sql = f"CAST(({y2} - {y1}) * 360 + ({m2} - {m1}) * 30 + ({day2} - {day1}) AS DOUBLE)"
    return Expr(f"(CASE WHEN {_ok(a)} AND {_ok(b)} THEN {sql} END)", "num", errs(a.err, b.err))


@register("DATEDIF")
def datedif(c, args, sheet):
    arity("DATEDIF", args, 3, 3)
    if not isinstance(args[2], Str) or args[2].value.upper() not in ("Y", "M", "D"):
        raise Unsupported("DATEDIF unit must be \"Y\", \"M\" or \"D\" written in the formula")
    unit = args[2].value.upper()
    a, b = _serial(c, args[0], sheet), _serial(c, args[1], sheet)
    d1, d2 = to_date(a.sql), to_date(b.sql)
    months = (f"((year({d2}) - year({d1})) * 12 + month({d2}) - month({d1}) "
              f"- CASE WHEN day({d2}) < day({d1}) THEN 1 ELSE 0 END)")
    sql = {"D": f"({b.sql} - {a.sql})", "M": months, "Y": f"floor({months} / 12)"}[unit]
    return Expr(f"(CASE WHEN {_ok(a)} AND {_ok(b)} AND {a.sql} <= {b.sql} THEN CAST({sql} AS DOUBLE) END)", "num",
                errs(a.err, b.err, f"(CASE WHEN {a.sql} > {b.sql} THEN '#NUM!' END)"))


def _date_format(pattern: str):
    def sql(x: str) -> str:
        s = f"floor({x})"
        return f"(CASE WHEN {s} BETWEEN 1 AND {MAX_SERIAL} AND {s} <> 60 THEN strftime({to_date(s)}, '{pattern}') END)"
    sql.is_date = True  # TEXT refuses date formats in 1904-system workbooks
    return sql


DATE_FORMATS = {
    "yyyy-mm-dd": "%Y-%m-%d", "mm/dd/yyyy": "%m/%d/%Y", "m/d/yyyy": "%-m/%-d/%Y", "dd/mm/yyyy": "%d/%m/%Y",
    "d/m/yyyy": "%-d/%-m/%Y", "dd-mmm-yyyy": "%d-%b-%Y", "d-mmm-yy": "%-d-%b-%y", "dd-mmm-yy": "%d-%b-%y",
    "mmm": "%b", "mmmm": "%B", "yyyy": "%Y", "yy": "%y", "mmm-yy": "%b-%y", "mmm yyyy": "%b %Y",
    "mmmm yyyy": "%B %Y", "ddd": "%a", "dddd": "%A", "dd": "%d", "d": "%-d", "mm": "%m",
}
FORMATS.update({k: _date_format(v) for k, v in DATE_FORMATS.items()})
```

- [ ] **Step 5: Register it: replace `src/unexcel/fn/__init__.py`**

```python
"""The Excel function library. Importing this package registers every function in unexcel.functions.FUNCS."""
from unexcel.fn import aggregate, criteria, dates, logic, lookup, maths, text  # noqa: F401
```

- [ ] **Step 6: Run tests to verify they pass**

Run: `python -m pytest tests -v`
Expected: all passed.

- [ ] **Step 7: Commit**

```bash
git add src/unexcel/fn src/unexcel/model.py src/unexcel/reader.py src/unexcel/compute.py tests/test_dates.py
git commit -m "feat: date functions on Excel's serial calendar, TEXT date formats" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Review regression gate (after Task 8)

The corrected code above includes the October 12 review fixes. Preserve these checks when building the real branch. They use synthetic expectations; Excel-saved values remain the final referee.

- [ ] Create `tests/test_review_regressions.py` with the following tests. When making further fixes, reproduce the failing test before editing code.

```python
"""Plan 2b review regressions; synthetic expectations, not Excel-saved fixtures."""
import pytest
from xlsxkit import book_of

import unexcel.compiler as compiler
from unexcel.compiler import compile_book
from unexcel.verifier import execute


def result(formula, inputs, expected):
    book = book_of({"S": {**inputs, "D1": (formula, expected), "E1": ("=1+1", 2)}})
    compiled = compile_book(book)
    assert not compiled.unresolved
    got = execute(book, compiled)
    assert got[("S", 1, 5)] == ("n", 2.0, None)
    kind, number, text = got[("S", 1, 4)]
    expected_kind = "n"
    if isinstance(expected, str):
        expected_kind = "e" if expected.startswith("#") else "s"
    assert kind == expected_kind
    return text if kind in ("s", "e") else number


@pytest.mark.parametrize("formula,inputs", [
    ("=INDEX(A1:A2,1E30)", {"A1": 1, "A2": 2}),
    ("=CHOOSE(1E30,1,2)", {}),
    ('=LEFT("abc",1E30)', {}),
    ("=EDATE(46023,1E30)", {}),
    ("=ROUNDUP(1,1000)", {}),
])
def test_supported_numeric_inputs_do_not_abort_batch(formula, inputs):
    # A bad argument must never prevent the companion cell from completing.
    book = book_of({"S": {**inputs, "D1": (formula, 1), "E1": ("=1+1", 2)}})
    got = execute(book, compile_book(book))
    assert got[("S", 1, 5)] == ("n", 2.0, None)


@pytest.mark.parametrize("formula", ["=EDATE(2958465,1)", "=EDATE(1,-1)", "=EOMONTH(2958465,1)"])
def test_shifted_date_stays_in_excel_domain(formula):
    assert result(formula, {}, "#NUM!") == "#NUM!"


def test_rept_respects_excel_character_limit():
    assert result('=LEN(REPT("a",32768))', {}, "#VALUE!") == "#VALUE!"


@pytest.mark.parametrize("mode,direction,expected", [(-1, 1, 10), (1, -1, 20), (0, 1, 10), (0, -1, 20)])
def test_xlookup_duplicate_follows_search_direction(mode, direction, expected):
    inputs = {"A1": 5, "A2": 5, "B1": 10, "B2": 20}
    formula = f'=XLOOKUP(5,A1:A2,B1:B2,"none",{mode},{direction})'
    assert result(formula, inputs, expected) == expected


def test_median_keeps_each_range_argument():
    assert result("=MEDIAN(A1:A2,A1:A1)", {"A1": 0, "A2": 10}, 0) == 0


def test_lookup_preserves_empty_text():
    assert result('=VLOOKUP("x",A1:B1,2,FALSE)', {"A1": "x", "B1": ""}, "") == ""


def test_aggregate_keeps_argument_error_order():
    assert result("=MAX(#DIV/0!,A1:A1)", {"A1": "#N/A"}, "#DIV/0!") == "#DIV/0!"


def test_sql_guard_counts_text_and_kind(monkeypatch):
    monkeypatch.setattr(compiler, "MAX_SQL", 1000)
    book = book_of({"S": {"D1": ('=IF(TRUE,"' + 'x' * 940 + '",0)', 'x' * 940)}})
    compiled = compile_book(book)
    assert compiled.unresolved == {("S", 1, 4): compiler.TOO_BIG}


@pytest.mark.parametrize("formula,expected", [
    ('=INDEX(A1:A2,1E30)', '#REF!'),
    ('=CHOOSE(1E30,1,2)', '#VALUE!'),
    ('=LEFT("abc",1E30)', 'abc'),
    ('=RIGHT("abc",1E30)', 'abc'),
    ('=MID("abc",1E30,1)', ''),
    ('=MID("abc",1,1E30)', 'abc'),
    ('=SUBSTITUTE("aba","a","x",1E30)', 'aba'),
    ('=REPLACE("abc",1E30,1E30,"x")', 'abcx'),
    ('=FIND("a","abc",1E30)', '#VALUE!'),
    ('=SEARCH("a","abc",1E30)', '#VALUE!'),
    ('=EDATE(46023,1E30)', '#NUM!'),
    ('=EOMONTH(46023,1E30)', '#NUM!'),
    ('=WEEKDAY(1E30)', '#NUM!'),
    ('=ROUNDUP(1,1000)', 1),
    ('=ROUNDDOWN(1,1000)', 1),
    ('=TRUNC(1,1000)', 1),
    ('=LEN(REPT("a",32767))', 32767),
    ('=LEN(REPT("ab",16384))', '#VALUE!'),
    ('=REPT("",1E30)', ''),
    ('=REPT("a",-1)', '#VALUE!'),
])
def test_numeric_bounds_and_text_limits(formula, expected):
    assert result(formula, {"A1": 1, "A2": 2}, expected) == expected

@pytest.mark.parametrize("name", ["AVERAGE", "MIN", "MAX", "PRODUCT"])
@pytest.mark.parametrize("first", [True, False])
def test_aggregate_error_order_in_both_directions(name, first):
    arguments = '#DIV/0!,A1:A1' if first else 'A1:A1,#DIV/0!'
    expected = '#DIV/0!' if first else '#N/A'
    assert result(f'={name}({arguments})', {"A1": "#N/A"}, expected) == expected

@pytest.mark.parametrize("formula", [
    '=HLOOKUP("x",A1:A2,2,FALSE)', '=INDEX(A2:A2,1)',
    '=XLOOKUP("x",A1:A1,A2:A2)',
])
def test_shared_lookup_reader_keeps_empty_text(formula):
    assert result(formula, {"A1": "x", "A2": ""}, "") == ""


def test_index_rejects_large_offsets_before_address_arithmetic():
    assert result('=INDEX(A1000000:A1000001,9.2233720368547748E18)',
                  {"A1000000": 1, "A1000001": 2}, "#REF!") == "#REF!"
```

- [ ] Run `python -m pytest tests/test_review_regressions.py -v`, then `python -m pytest tests -q`. The regression file covers extreme numeric arguments, date and text bounds, lookup duplicate direction and empty text, aggregate error order and repeated MEDIAN references, and the complete SQL-size guard.
- [ ] Include `tests/test_review_regressions.py` in the Task 9 commit.

Rounding whose scaled intermediate cannot fit in DOUBLE returns #NUM!; real-Excel boundary fixtures must settle those extreme cases. The example generator includes these open questions. The previous 200k timing predates these fixes and must be measured again on the real branch.

---

### Task 9: Real-Excel ground truth for the library, and the 200k re-check

**Files:**
- Create: `examples/make_functions.py`
- Modify: `tests/test_excel_saved.py` (append), `.gitignore`, `PROGRESS.md`

**Interfaces:**
- Consumes: everything above.
- Produces: `examples/functions.xlsx` (saved by real Excel, committed). It settles the behaviours this plan could not confirm without Excel:
  - Does COUNTIF with a number also count the text "5"? Are empty and error cells counted by `"<>5"`? Is a formula's `""` counted by `"<>"`?
  - How does a number become text for `1E-05` and `1E+15`?
  - What do MATCH and VLOOKUP return on unsorted data?
  - DAYS360 at the end of February; `TEXT(-0.001, "0.00")`; `MOD(1, 0.1)`; CHAR(128) on Mac Excel.

- [ ] **Step 1: Create `examples/make_functions.py`**

```python
"""Writes examples/functions_unsaved.xlsx: the Excel behaviours Plan 2b could not confirm without Excel.

Open it in real Microsoft Excel (Windows or Mac), then File > Save As examples/functions.xlsx.
tests/test_excel_saved.py then checks unexcel's answers against Excel's own.
"""
from pathlib import Path

import openpyxl

DATA = {  # column A: a mixed range; B: numbers; C: a formula that returns ""; D: sorted brackets; E: unsorted
    "A1": 5, "A2": "5", "A3": "apple", "A4": "Banana", "A5": True, "A7": "a*b", "A8": "=1/0",
    "B1": 10, "B2": 20, "B3": 30, "B4": 40, "B5": 50, "B6": 60, "B7": 70, "B8": 80,
    "C1": '=IF(B1>0,"","x")', "D1": 0, "D2": 100, "D3": 500, "E1": 30, "E2": 10, "E3": 20, "H1": 5, "H2": 5, "H3": "#N/A",
}
CHECKS = {  # the open question each one answers
    "G1": '=COUNTIF(A1:A8,5)',            # does a number criterion also count the text "5"?
    "G2": '=COUNTIF(A1:A8,"<>5")',        # are empty and error cells counted by "<>5"?
    "G3": '=COUNTIF(A1:C1,"<>")',         # is a formula's "" counted as not empty?
    "G4": '=COUNTIF(A1:A8,">b")',         # text comparison ignores case
    "G5": '=SUMIF(A1:A8,"<>apple",B1:B8)',  # a matched error row (A8 is #DIV/0!, B8 a number)
    "G6": '=SUMIF(B1:B8,">30",A1:A8)',    # sum range holding an error in a matched row
    "G7": '="x"&0.00001',                 # number -> text for small numbers
    "G8": '="x"&1E+15',
    "G9": '="x"&123456789012345678',
    "G10": '=TRIM("  a   b  ")',
    "G11": '=PROPER("hELLO o\'neil 2nd")',
    "G12": '=MATCH(15,E1:E3)',            # approximate match on unsorted data
    "G13": '=MATCH(150,D1:D3,-1)',
    "G14": '=VLOOKUP("b",A3:B4,2)',       # approximate text match
    "G15": '=DAYS360(DATE(2026,1,30),DATE(2026,2,28))',
    "G16": '=DAYS360(DATE(2026,2,28),DATE(2026,3,31))',
    "G17": '=TEXT(-0.001,"0.00")',
    "G18": '=MOD(1,0.1)',
    "G19": '=ROUNDUP(0.1*3,1)',
    "G20": '=CHAR(128)',
    "G21": '=1<"a"',
    "G22": '="a"<TRUE',
    "G23": '=IFERROR(1/0,"")',
    "G24": '=AVERAGE(A1:A3,"3")',
    "G25": '=INDEX(F1:F3,2)',             # an empty cell found
    "G26": '=XLOOKUP(250,D1:D3,B1:B3,0,-1)',
    "G27": '=SUBTOTAL(9,B1:B3,G28)',
    "G28": '=SUBTOTAL(9,B1:B2)',
    "G29": '=SUM(A8,B1)',                 # error order: the first error wins
    "G30": '=DATE(1900,3,1)',
    "G31": '=WEEKDAY(DATE(2026,1,1))',
    "G32": '=TEXT(DATE(2026,3,6),"dd-mmm-yyyy")',
    "G33": '=DATE(1900,2,29)',
    "G34": '=MAX(1/0,H3)',                 # scalar error before a range error
    "G35": '=MAX(H3,1/0)',
    "G36": '=ROUNDUP(1,1000)',
    "G37": '=ROUNDUP(1,-1000)',
    "G38": '=EDATE(2958465,1)',
    "G39": '=EDATE(1,-1)',
    "G40": '=EOMONTH(2958465,1)',
    "G41": '=LEN(REPT("a",32767))',
    "G42": '=LEN(REPT("a",32768))',
    "G43": '=VLOOKUP(10,B1:C1,2,FALSE)',    # lookup of a formula returning empty text
    "G44": '=MEDIAN(B1:B2,B1:B1)',
    "G45": '=XLOOKUP(5,H1:H2,B1:B2,"none",-1,1)',
    "G46": '=XLOOKUP(5,H1:H2,B1:B2,"none",1,-1)',
}

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Functions"
for ref, value in {**DATA, **CHECKS}.items():
    ws[ref] = value
wb.save(Path(__file__).parent / "functions_unsaved.xlsx")
```

- [ ] **Step 2: Append to `tests/test_excel_saved.py`**

```python
FUNCTIONS_EXAMPLE = Path(__file__).parent.parent / "examples" / "functions.xlsx"


@pytest.mark.skipif(not FUNCTIONS_EXAMPLE.exists(),
                    reason="save examples/functions.xlsx from real Excel (PLAN-2b Task 9)")
def test_real_excel_function_answers_match():
    book = load(FUNCTIONS_EXAMPLE)
    verdicts = verify(book, compile_book(book))
    assert [v for v in verdicts if v.status != "PASS"] == []
    assert len(verdicts) == 48  # A8, C1 and the 46 checks in column G
```

Track `examples/functions_unsaved.xlsx` in Git at the user's request, so it can be downloaded directly. Keep `.superpowers/` ignored.

- [ ] **Step 3: Run the tests**

Run: `python examples/make_functions.py` then `python -m pytest tests -v`
Expected: all passed; `test_real_excel_function_answers_match` is skipped until Step 6.

- [ ] **Step 4: Re-run the 200k perf test**

Run: `UNEXCEL_PERF=1 python -m pytest tests/test_perf.py -s -v`
Expected: passed. Record the printed stage times in the PR (Plan 2a measured 254.9 s; the draft of this plan measured 260.5 s (verify 253.0 s, 801 batches)). If it fails the 300 s budget, record the times and stop: the fix is Plan 2c's table mode, not a tweak here.

- [ ] **Step 5: Update `PROGRESS.md`**

In the "Where things are" table, add a row for PR #3 (`plan-2b`, stacked on #2). Add a section "## Plan 2b — done (PR #3)" listing the 85 new functions by family, the test count and the perf times. Under "Still needs you", add item 3: run `python examples/make_functions.py`, open `examples/functions_unsaved.xlsx` in Excel, Save As `examples/functions.xlsx`, and commit it to `plan-2b`.

- [ ] **Step 6: MANUAL, needs the user's Excel.** Save `examples/functions.xlsx` as in Step 5, then:

Run: `python -m pytest tests/test_excel_saved.py -v`
Expected: passed. Each FAIL names a cell in column G, and the comment beside it in `make_functions.py` says which rule it tests. Change the compiler to match Excel and add the case to the matching `tests/test_<family>.py` with Excel's value. Don't change the expectation.

- [ ] **Step 7: Commit**

```bash
git add examples/make_functions.py tests/test_excel_saved.py .gitignore PROGRESS.md
git commit -m "test: real-Excel ground truth for the function library (Excel save pending)" -m "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## Self-review (done while writing)

- **Spec coverage:** DESIGN §2.2 "~80 most common functions": 85 new, picked from two published corpora (table above); every one of their top functions is covered except INDIRECT (volatile, UNRESOLVED by design) and reference-returning INDEX. "Every unsupported cell is UNRESOLVED": each unsupported form raises `Unsupported` with a reason, and the tests check several by name. §2.3 match rule unchanged. §6 "numbers stored as text": SUM-like functions still report them (`note_text_numbers`), and the criteria engine reads them the way COUNTIF does. §6 dates: 1900 leap-year bug (serial 60) handled; 1904 system detected and UNRESOLVED. §7 real-Excel fixtures: Task 9.
- **Placeholders:** none. Every code block is the tested draft file, copied by script.
- **Name consistency:**
  - `like_pattern` is defined in Task 2 and used in Task 4; `int_literal` is defined in Task 4 and used in Task 8; `FORMATS` is defined in Task 7 and extended in Task 8.
  - The `is_date` marker is set in Task 8 and checked in Task 7's `text()`; `getattr` defaults to False before Task 8.
  - `Compiler.kind_of` returns `"a"` for any-typed cells (Task 1), which `lookup._type_of` treats as mixed.
  - `fn/__init__.py` lists exactly the modules created so far in each task.
- **Verified against real libraries:** DuckDB 1.5.6 behaviour listed under "Measured while planning", and the full draft suite (97 passed) including a real `dbt build`.
- **Not verified, settled by Task 9:** the column-G questions in `examples/make_functions.py`.
