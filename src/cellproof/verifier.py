"""Run the compiled batches on DuckDB and compare every cell with Excel's saved value."""
from __future__ import annotations

import math
import tempfile
from pathlib import Path

import duckdb

from cellproof.compiler import READ_SEED, Compiled, batch_rows, write_seed
from cellproof.expr import SETTINGS
from cellproof.model import Book, Cell, Key, Verdict, a1


class SqlFailed(Exception):
    """DuckDB rejected SQL that cellproof generated (a bug: every supported formula must run)."""

    def __init__(self, message: str, keys: list):
        super().__init__(message)
        self.keys = keys  # the cells of the batch DuckDB rejected


MAX_RECOMPILES = 50


def verify_book(book: Book, compile_fn) -> tuple:
    """(compiled, verdicts, failures): verify; when DuckDB rejects a batch, its cells become UNRESOLVED with the reason
    and the book is compiled again, so every cell reading them is "depends on unresolved cell" - none can read a
    missing value. Each round removes a batch."""
    failures = []
    for _ in range(MAX_RECOMPILES):
        compiled = compile_fn(book)
        try:
            return compiled, verify(book, compiled), failures
        except SqlFailed as err:
            failures.append(str(err))
            for key in err.keys:
                book.unread[key] = f"DuckDB could not run the SQL cellproof generated ({str(err).split('(', 1)[-1]}"
    raise SqlFailed(f"DuckDB rejected SQL in {MAX_RECOMPILES} rounds; giving up", [])


def execute(book: Book, compiled: Compiled) -> dict[Key, tuple]:
    """The same row SQL as the dbt project: the seed, then each batch appended in place. Returns formula cells only."""
    con = duckdb.connect()
    with tempfile.TemporaryDirectory() as tmp:
        seed = Path(tmp) / "input_cells.csv"
        write_seed(book, seed)
        con.execute(f"CREATE TABLE cells AS SELECT * FROM {READ_SEED}", [str(seed)])
    for batch in compiled.batches:
        try:
            con.execute(f"INSERT INTO cells {batch_rows(batch, 'cells')}")
        except duckdb.Error as err:
            raise SqlFailed(f"DuckDB could not run the SQL for {a1(batch[0].key)} and {len(batch) - 1} cells after it "
                            f"({str(err).splitlines()[0]}). This is a bug in cellproof; please report it.",
                            [s.key for s in batch]) from err
    wanted = {s.key for s in compiled.steps}
    rows = con.execute("SELECT * FROM cells").fetchall()
    return {(s, r, c): (k, n, t) for s, r, c, k, n, t in rows if (s, r, c) in wanted}


def matches(cell: Cell, kind: str, num, txt) -> bool:
    """Excel's saved value vs the SQL result. An Excel error matches only the same error code."""
    return same(kind, num, txt, cell.kind, cell.value)


def same(kind: str, num, txt, other_kind: str, other) -> bool:
    """Is an answer (`kind`/`num`/`txt`, the way a query row holds it) the same as another answer (`other_kind` and
    its plain value `other`)? One answer holds a number in `num`, a boolean as 1.0/0.0 in `num`, and text or an error
    code in `txt`. `matches` wraps this for a cell and `prove --check` calls it on two recorded answers, so a bundle
    cannot pass a check that the verify run would fail - there is one notion of "equal" in cellproof."""
    if kind == "e":
        return other_kind == "e" and other == txt
    if other_kind == "e" or (num is None and txt is None):
        return False
    if kind == "n":
        return other_kind == "n" and math.isclose(other, num, rel_tol=1e-9, abs_tol=1e-12)
    if kind == "b":
        return other_kind == "b" and num == (1.0 if other else 0.0)
    if kind == "s":
        return other_kind == "s" and other == txt
    return False


def verify(book: Book, compiled: Compiled) -> list[Verdict]:
    return verdicts(book, compiled, execute(book, compiled))


def verdicts(book: Book, compiled: Compiled, got: dict, blocked: dict | None = None) -> list[Verdict]:
    """Every cell's verdict from the computed rows; `blocked` cells (not computed, with why) are UNRESOLVED."""
    out = []
    for step in compiled.steps:
        cell = book.cells[step.key]
        if blocked and step.key in blocked:
            out.append(Verdict(step.key, "UNRESOLVED", cell.value, None, blocked[step.key]))
            continue
        kind, n, t = got[step.key]
        if kind == "e" and t in SETTINGS:  # the answer depends on settings the file does not record
            out.append(Verdict(step.key, "UNRESOLVED", cell.value, None, SETTINGS[t]))
            continue
        actual = t if kind in ("s", "e") else (bool(n) if kind == "b" and n is not None else n)
        out.append(Verdict(step.key, "PASS" if matches(cell, kind, n, t) else "FAIL", cell.value, actual))
    for key, reason in compiled.unresolved.items():
        out.append(Verdict(key, "UNRESOLVED", book.cells[key].value, None, reason))
    return sorted(out, key=lambda v: v.key)
