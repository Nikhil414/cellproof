from xlsxkit import book_of

import unexcel.compiler as compiler
from unexcel.compiler import compile_book
from unexcel.verifier import execute, verify


def run(sheets):
    book = book_of(sheets)
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def test_thousands_of_cells_no_longer_crash():  # Plan 1: DuckDB depth error at ~1k cells, segfault at 4k
    n = 3000
    spec = {f"A{r}": float(r) for r in range(1, n + 1)}
    spec.update({f"B{r}": (f"=A{r}*2", 2.0 * r) for r in range(1, n + 1)})
    book = book_of({"S": spec})
    compiled = compile_book(book, regions=False)
    assert len(compiled.batches) >= n // compiler.BATCH
    assert max(len(b) for b in compiled.batches) <= compiler.BATCH
    assert {v.status for v in verify(book, compiled)} == {"PASS"}


def test_batches_close_by_weight(monkeypatch):
    monkeypatch.setattr(compiler, "BATCH_WEIGHT", 1000)
    book = book_of({"S": {f"A{r}": (f"={r}", float(r)) for r in range(1, 21)}})
    sizes = [len(b) for b in compile_book(book).batches]
    assert sum(sizes) == 20 and max(sizes) < 20


def test_a_cell_heavier_than_the_budget_runs_alone(monkeypatch):  # Review Focus 2
    monkeypatch.setattr(compiler, "BATCH_WEIGHT", 1)
    book = book_of({"S": {f"A{r}": (f"={r}+{r}", 2.0 * r) for r in range(1, 6)}})
    compiled = compile_book(book)
    assert [len(b) for b in compiled.batches] == [1, 1, 1, 1, 1]
    assert {v.status for v in verify(book, compiled)} == {"PASS"}


def test_running_total_thousands_of_rows_deep():  # Review Focus 3
    n = 3000
    spec = {f"A{r}": 1 for r in range(1, n + 1)}
    spec["B1"] = ("=A1", 1)
    spec.update({f"B{r}": (f"=B{r - 1}+A{r}", float(r)) for r in range(2, n + 1)})
    book = book_of({"S": spec})
    verdicts = verify(book, compile_book(book))
    assert len(verdicts) == n and {v.status for v in verdicts} == {"PASS"}


def test_batch_size_limit(monkeypatch):
    monkeypatch.setattr(compiler, "BATCH", 3)
    book = book_of({"S": {f"A{r}": (f"={r}", float(r)) for r in range(1, 8)}})
    assert [len(b) for b in compile_book(book).batches] == [3, 3, 1]


def test_range_over_the_open_batch_closes_it_first():  # Review Focus 4
    book = book_of({"S": {"A1": 1, "B1": ("=A1+1", 2), "B2": ("=A1+2", 3),
                          "C1": ("=SUM(B1:B2)", 5), "C2": ("=C1*2", 10)}})
    compiled = compile_book(book)
    assert [[s.key for s in b] for b in compiled.batches] == [
        [("S", 1, 2), ("S", 2, 2)],  # B1, B2
        [("S", 1, 3), ("S", 2, 3)],  # C1 (range over B1:B2 -> new batch), C2 (single ref -> same batch)
    ]
    assert {v.status for v in verify(book, compiled)} == {"PASS"}


def test_formula_results_feed_sumifs_and_vlookup():  # Plan 1 rejected ranges holding formula cells
    got, c = run({"S": {
        "A1": "x", "A2": "y", "A3": "x",
        "B1": 10, "B2": 20, "B3": 30,
        "C1": ("=B1*2", 20), "C2": ("=B2*2", 40), "C3": ("=B3*2", 60),
        "D1": ('=SUMIFS(C1:C3, A1:A3, "x")', 80),
        "D2": ('=VLOOKUP("y", A1:C3, 3, FALSE)', 40),
    }})
    assert c.unresolved == {}
    assert got[("S", 1, 4)] == ("n", 80.0, None)
    assert got[("S", 2, 4)] == ("n", 40.0, None)


def test_empty_text_input_stays_empty_text():  # the seed CSV reads "" back as NULL
    got, _ = run({"S": {"A1": "", "A2": ('=A1&"x"', "x")}})
    assert got[("S", 2, 1)] == ("s", None, "x")


from unexcel.compiler import render_batch


def test_own_only_batch_has_only_the_new_rows():
    book = book_of({"S": {"A1": 2, "A2": ("=A1*2", 4), "A3": ("=A1*3", 6)}})
    steps = compile_book(book).batches[0]
    full, own = render_batch(steps, "prev"), render_batch(steps, "prev", own_only=True)
    assert "SELECT * FROM prev" in full and "SELECT * FROM prev" not in own
    assert own.count("UNION ALL") == len(steps) - 1


def test_workbook_without_formulas_verifies_and_executes():  # Review Focus 1
    book = book_of({"S": {"A1": 1.0, "A2": "x"}})
    compiled = compile_book(book)
    assert compiled.batches == [] and execute(book, compiled) == {}
    assert verify(book, compiled) == []


from unexcel.compiler import batch_rows  # noqa: E402


def test_batch_rows_of_a_cell_batch_is_its_own_rows_sql():
    book = book_of({"S": {"A1": 2, "A2": ("=A1*2", 4)}})
    steps = compile_book(book).batches[0]
    assert batch_rows(steps, "cells") == render_batch(steps, "cells", own_only=True)
