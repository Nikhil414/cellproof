"""200k-formula workbook, timed (DESIGN.md §2.1 and §7). Slow, so opt-in:

    UNEXCEL_PERF=1 python -m pytest tests/test_perf.py -s
"""
import os
import time

import pytest
from xlsxkit import write_xlsx

from unexcel.compiler import compile_book
from unexcel.emitter import write_project
from unexcel.reader import load
from unexcel.verifier import verify

pytestmark = pytest.mark.skipif(not os.environ.get("UNEXCEL_PERF"), reason="slow: set UNEXCEL_PERF=1")
ROWS = 50_000  # four formula columns -> 200,000 formula cells, plus two summary cells


def make(path):
    """A finance-shaped sheet: amount, amount x rate, running total, threshold, difference."""
    data, total, d_sum = {}, 0.0, 0.0
    for r in range(1, ROWS + 1):
        amount = float(r % 997)
        b = amount * 1.1
        total = b if r == 1 else total + b  # same additions, same order as SQL: exact match
        d = b if b > 100 else 0.0
        d_sum += d
        data[f"A{r}"] = amount
        data[f"B{r}"] = (f"=A{r}*1.1", b)
        data[f"C{r}"] = ("=B1" if r == 1 else f"=C{r - 1}+B{r}", total)
        data[f"D{r}"] = (f"=IF(B{r}>100,B{r},0)", d)
        data[f"E{r}"] = (f"=D{r}-A{r}", d - amount)
    summary = {"A1": ("=SUM(Data!D:D)", d_sum), "A2": (f"=Data!C{ROWS}", total)}
    write_xlsx(path, {"Data": data, "Summary": summary})


def test_200k_formulas_under_five_minutes(tmp_path):
    wb = tmp_path / "big.xlsx"
    make(wb)
    marks = [("start", time.perf_counter())]
    book = load(wb)
    marks.append(("read", time.perf_counter()))
    compiled = compile_book(book)
    marks.append(("compile", time.perf_counter()))
    verdicts = verify(book, compiled)
    marks.append(("verify", time.perf_counter()))
    write_project(book, compiled, tmp_path / "out", "big")
    marks.append(("emit", time.perf_counter()))
    for (_, t0), (name, t1) in zip(marks, marks[1:]):
        print(f"{name:8s} {t1 - t0:6.1f} s")
    total_s = marks[-1][1] - marks[0][1]
    print(f"{'total':8s} {total_s:6.1f} s  ({len(compiled.batches)} batches)")
    assert len(verdicts) == 4 * ROWS + 2
    assert {v.status for v in verdicts} == {"PASS"}
    assert total_s < 300  # DESIGN.md §2.1: under 5 minutes on a laptop
