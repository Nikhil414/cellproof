"""The SQL for a selection of cells: `cellproof sql` prints a runnable chain, it does not run it.

The whole workbook is compiled (regions and all), then everything the selection does not need is dropped: the
batches whose cells nothing in the selection reads, and the input cells nobody reads. What is left is the same SQL
the verifier would run, in the same order, as statements you can paste into DuckDB:

    CREATE OR REPLACE TABLE cells (...);
    INSERT INTO cells VALUES (...);          -- the input cells the chain reads
    INSERT INTO cells <batch>;               -- one statement per batch, in dependency order
    ...
    SELECT ... FROM cells WHERE ...;          -- the cells you asked for

Batches are dropped whole, never partially: a batch is one statement, so including it means including every cell it
computes. A single cell whose batch is kept can read cells of an earlier batch, so the batches are walked backwards
and any batch they touch is kept in turn - what stays is closed under "reads".

A selection is every non-empty cell of the rectangle, inputs included, so the last SELECT returns exactly those -
minus the cells the compiler could not translate, which are not in the table at all. Each of those reasons goes to
stderr, ranked: an unresolved cell is a hole in the answer, and it is better to say so than to hide it.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

from cellproof.compiler import INPUT_COLUMNS, Compiled, batch_rows
from cellproof.expr import lit_str
from cellproof.model import Book, Cell, Key, Unsupported, a1
from cellproof.r1c1 import lexed
from cellproof.refs import MAX_ROW, resolve

MAX_COL = 16_384
SEED_ROWS = 1_000  # input rows per INSERT: one statement per 1,000 cells keeps each one small


class QueryError(Exception):
    """The selection cannot be made. The message is shown to the user as it is."""


@dataclass
class Answer:
    sql: str
    selection: list[Key]  # the cells the last SELECT returns, row-major
    computed: list[Key]  # the formula cells the chain computes
    unresolved: dict[Key, str]  # selected cells that are not in the chain, with the reason
    groups: list[dict]  # the reasons ranked, the same way the report ranks them
    batches: int
    inputs: int


@dataclass(frozen=True)
class Target:
    """What was asked for: the cells, a short label, and the SQL that picks them out of the cells table."""

    keys: list[Key]
    label: str
    where: str  # a predicate over cells.sheet / row_num / col_num; "" is every row
    order: str  # the ORDER BY body


def box(text: str) -> tuple[int, int, int, int]:
    """'D2:D1000' -> (2, 4, 1000, 4). Whole-column and whole-row text is refused: an A1 range says what it means."""
    if "!" in text:
        raise QueryError(f"give the sheet with --sheet, not in --range: {text!r}")
    try:
        c1, r1, c2, r2 = range_boundaries(text.replace("$", "").strip().upper())
    except ValueError:
        raise QueryError(f"not an A1 range: {text!r} (like D2:D1000)") from None
    if c1 is None:
        raise QueryError(f"not an A1 range: {text!r} (like D2:D1000)")
    if r1 is None or r2 is None:  # D:D, a whole column
        raise QueryError(f"{text!r} is a whole column: name its rows "
                         f"(like {get_column_letter(c1)}2:{get_column_letter(c2)}1000)")
    if c1 > MAX_COL or c2 > MAX_COL:
        raise QueryError(f"not an A1 range: {text!r} (a sheet has {get_column_letter(MAX_COL)} columns)")
    return r1, c1, r2, c2


def check(book: Book, sheet: str | None, rng: str | None) -> None:
    """Refuse a selection that cannot be made, before the workbook is compiled (a typo'd range costs one load, not a
    compile). `selection` runs the same checks; this one only reports."""
    if sheet is not None and sheet not in book.sheets:
        raise QueryError(f"no sheet {sheet!r} in this workbook: {', '.join(book.sheets)}")
    if rng is None:
        return
    if sheet is None and len(book.sheets) != 1:
        raise QueryError(f"--range needs --sheet: this workbook holds {', '.join(book.sheets)}")
    box(rng)


def selection(book: Book, sheet: str | None, rng: str | None) -> Target:
    """The cells to print. --range needs a sheet: --sheet, or the workbook's only one."""
    check(book, sheet, rng)
    if rng is None:
        if sheet is None:
            keys = _shown(book, sorted(k for k in book.cells if k[0] not in book.arrays))
            where = ("sheet NOT IN (" + ", ".join(lit_str(s).sql for s in sorted(book.arrays)) + ")"
                     if book.arrays else "")
            return Target(keys, "the whole workbook", where, "sheet, row_num, col_num")
        keys = _shown(book, book.present(sheet, 1, 1, MAX_ROW, MAX_COL))
        return Target(keys, sheet, f"sheet = {lit_str(sheet).sql}", "row_num, col_num")
    if sheet is None:
        sheet = book.sheets[0]
    r1, c1, r2, c2 = box(rng)
    where = " AND ".join([f"sheet = {lit_str(sheet).sql}",
                          f"row_num BETWEEN {r1} AND {r2}" if r1 != r2 else f"row_num = {r1}",
                          f"col_num BETWEEN {c1} AND {c2}" if c1 != c2 else f"col_num = {c1}"])
    label = f"{sheet}!{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"
    return Target(_shown(book, book.present(sheet, r1, c1, r2, c2)), label, where, "row_num, col_num")


def _shown(book: Book, keys: list[Key]) -> list[Key]:
    """The cells that hold something to show: a formula, or an input with a value. A styled empty cell has no row
    in the cells table in either mode, so it is not part of the answer."""
    return [k for k in keys if k in book.cells
            and (book.cells[k].formula is not None or book.cells[k].kind != "blank")]


def formula_reads(book: Book, key: Key) -> list[Key]:
    """The formula cells a formula cell reads. Only used to walk the chain backwards, so it may over-approximate:
    a range is taken whole even where Excel's implicit intersection reads one cell of it."""
    shape, texts = lexed(book).get(key, (None, ()))
    if shape is None or not shape.parses():
        return []  # the compiler gives this cell its real reason, so it is not in the chain at all
    out = []
    for text in texts:
        if text[:6].lower() == "_xlpm.":  # a LET name (fn/logic.py): its value's references were lexed too
            continue
        for k in _cells(book, text, key[0]):
            if book.cells[k].formula is not None:
                out.append(k)
    return out


def input_reads(book: Book, key: Key) -> list[Key]:
    """The input cells a formula cell reads, the seed rows its statement needs."""
    shape, texts = lexed(book).get(key, (None, ()))
    if shape is None or not shape.parses():
        return []
    out = []
    for text in texts:
        if text[:6].lower() != "_xlpm.":
            out.extend(k for k in _cells(book, text, key[0]) if book.cells[k].formula is None)
    return out


def _cells(book: Book, text: str, sheet: str) -> list[Key]:
    """The non-empty cells a reference names; none when it cannot be resolved (the compiler says why for the cell)."""
    try:
        return resolve(text, sheet, book).present(book)
    except (Unsupported, RecursionError):
        return []


def array_cells(book: Book) -> set[Key]:
    """The constants behind {1,2,3} live on hidden sheets, which no reference text names, so the seed always carries
    them whole. They are literals written in a formula: there are few of them."""
    return {k for k, c in book.cells.items() if k[0] in book.arrays and c.formula is None and c.kind != "blank"}


def _chain(book: Book, compiled: Compiled, chosen: list[Key]) -> tuple[list, list[Key]]:
    """(the batches to print, in dependency order, the formula cells they compute) for the chosen cells."""
    steps = {s.key for s in compiled.steps}
    closure = {k for k in chosen if k in steps}
    if len(closure) < len(steps):  # a batch's cells read only earlier batches: the walk can stop early
        stack = list(closure)
        while stack:  # every formula cell those cells read, directly or through other formula cells
            for key in formula_reads(book, stack.pop()):
                if key not in closure:
                    closure.add(key)
                    stack.append(key)
    kept: list = []
    for batch in reversed(compiled.batches):
        if any(s.key in closure for s in batch):
            kept.append(batch)
            closure.update(s.key for s in batch)  # its statement computes all of them
    kept.reverse()
    return kept, sorted(k for k in closure if k in steps)


def reasons(compiled: Compiled, keys: list[Key]) -> list[dict]:
    """The selected cells that could not be translated, grouped by the reason to fix and ranked the way the report
    ranks them: a cell that stopped because it reads an unfinished cell counts towards the first unfinished cell its
    chain reaches (impact.roots_of), so the count says what a fix would let compile again."""
    from cellproof.impact import roots_of

    roots = roots_of(compiled.chains)
    counts: Counter = Counter()
    examples: dict[str, list[str]] = defaultdict(list)
    for key in keys:
        root = roots.get(key, key)
        reason = compiled.unresolved.get(root) or compiled.unresolved[key]
        counts[reason] += 1
        if root == key and len(examples[reason]) < 3:
            examples[reason].append(a1(key))
    ranked = sorted(counts, key=lambda r: (-counts[r], r))
    return [{"reason": r, "cells": counts[r], "examples": examples[r]} for r in ranked]


def _row(key: Key, cell: Cell) -> str:
    sheet, row, col = key
    num, text = _value(cell)
    return f"  ({lit_str(sheet).sql}, {row}, {col}, '{cell.kind}', {num}, {text})"


def _value(cell: Cell) -> tuple[str, str]:
    """(num, txt) SQL, exactly as compiler.input_rows lays them out."""
    if cell.kind == "b":
        return ("1.0" if cell.value else "0.0"), "NULL"
    if cell.kind == "n":
        return _double(float(cell.value)), "NULL"
    if cell.kind in ("e", "s"):
        return "NULL", lit_str(str(cell.value)).sql
    return "NULL", "NULL"  # blank: a cell the file holds without a value


def _double(x: float) -> str:
    """A DOUBLE literal. Python writes infinities as 'inf' and NaN as 'nan', which DuckDB reads as column names."""
    if x != x:
        return "CAST('nan' AS DOUBLE)"
    if x in (float("inf"), float("-inf")):
        return f"CAST('{'inf' if x > 0 else '-inf'}' AS DOUBLE)"
    return repr(x)


def build(book: Book, compiled: Compiled, target: Target) -> Answer:
    """The SQL for `target`: the seed rows and batches it needs, then one SELECT for the cells themselves."""
    kept, computed = _chain(book, compiled, target.keys)
    steps = {s.key for s in compiled.steps}
    seed = array_cells(book)
    for key in computed:
        seed.update(input_reads(book, key))
    for key in target.keys:  # an input cell in the selection: the last SELECT reads it from the table
        if book.cells[key].formula is None and book.cells[key].kind != "blank":
            seed.add(key)
    rows = [_row(k, book.cells[k]) for k in sorted(seed)]
    missing = [k for k in target.keys if book.cells[k].formula is not None and k not in steps]
    unresolved = {k: compiled.unresolved.get(k, "not compiled") for k in missing}
    groups = reasons(compiled, missing)

    out = [f"-- cellproof sql: {target.label}",
           f"-- {len(computed):,} formula cells computed from {len(rows):,} input cells, {len(kept):,} statement(s)"]
    if unresolved:
        out.append(f"-- {len(unresolved):,} of the cells asked for could not be translated: they are not in this "
                   "answer (the reasons are on stderr)")
    out += ["", f"CREATE OR REPLACE TABLE cells ({INPUT_COLUMNS});", ""]
    for i in range(0, len(rows), SEED_ROWS):
        out.append("INSERT INTO cells (sheet, row_num, col_num, kind, num, txt) VALUES")
        out.append(",\n".join(rows[i:i + SEED_ROWS]) + ";")
        out.append("")
    for batch in kept:
        more = len(batch) - 1
        out.append(f"-- {a1(batch[0].key)}" + (f" and {more:,} cells after it" if more else ""))
        out.append(f"INSERT INTO cells\n{batch_rows(batch, 'cells')};")
        out.append("")
    out.append("-- the cells asked for")
    out.append("SELECT sheet, row_num, col_num, kind, num, txt")
    out.append("FROM cells")
    if target.where:
        out.append(f"WHERE {target.where}")
    out.append(f"ORDER BY {target.order};")
    return Answer("\n".join(out) + "\n", target.keys, computed, unresolved, groups, len(kept), len(rows))
