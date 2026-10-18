"""Run the compiled batches on DuckDB and compare every cell with Excel's saved value."""
from __future__ import annotations

import math
import tempfile
from pathlib import Path

import duckdb

from unexcel.compiler import READ_SEED, Compiled, batch_rows, write_seed
from unexcel.model import Book, Cell, Key, Verdict, a1


class SqlFailed(Exception):
    """DuckDB rejected SQL that unexcel generated (a bug: every supported formula must run)."""


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
            # ponytail: the run stops; marking just this batch's cells (and their dependents) UNRESOLVED needs a recompile
            raise SqlFailed(f"DuckDB could not run the SQL for {a1(batch[0].key)} and {len(batch) - 1} cells after it "
                            f"({str(err).splitlines()[0]}). This is a bug in unexcel; please report it.") from err
    wanted = {s.key for s in compiled.steps}
    rows = con.execute("SELECT * FROM cells").fetchall()
    return {(s, r, c): (k, n, t) for s, r, c, k, n, t in rows if (s, r, c) in wanted}


def matches(cell: Cell, kind: str, num, txt) -> bool:
    """Excel's saved value vs the SQL result. An Excel error matches only the same error code."""
    if cell.kind == "e":
        return kind == "e" and txt == cell.value
    if kind == "e" or (num is None and txt is None):
        return False
    if cell.kind == "n":
        return kind == "n" and math.isclose(cell.value, num, rel_tol=1e-9, abs_tol=1e-12)
    if cell.kind == "b":
        return kind == "b" and num == (1.0 if cell.value else 0.0)
    if cell.kind == "s":
        return kind == "s" and txt == cell.value
    return False


def verify(book: Book, compiled: Compiled) -> list[Verdict]:
    got = execute(book, compiled)
    out = []
    for step in compiled.steps:
        cell = book.cells[step.key]
        kind, n, t = got[step.key]
        actual = t if kind in ("s", "e") else (bool(n) if kind == "b" and n is not None else n)
        out.append(Verdict(step.key, "PASS" if matches(cell, kind, n, t) else "FAIL", cell.value, actual))
    for key, reason in compiled.unresolved.items():
        out.append(Verdict(key, "UNRESOLVED", book.cells[key].value, None, reason))
    return sorted(out, key=lambda v: v.key)
