"""A region that cannot be ordered as one block is split around the cell between its columns instead of dissolved."""
import shutil

import pytest
from xlsxkit import book_of, write_xlsx

from unexcel.cli import main
from unexcel.compiler import compile_book
from unexcel.verifier import execute


def same(a, b):
    return (a[0] == b[0] and a[2] == b[2]
            and (a[1] == b[1] or (a[1] is not None and b[1] is not None and abs(a[1] - b[1]) <= 1e-9 * max(1, abs(a[1])))))


def region_cols(compiled):
    return sorted(sorted({s.key[2] for s in b}) for b in compiled.batches if hasattr(b, "wide_sql"))


def check_equal(book, compiled):
    without = compile_book(book, regions=False)
    assert compiled.unresolved == without.unresolved
    a, b = execute(book, compiled), execute(book, without)
    assert a.keys() == b.keys()
    for key in a:
        assert same(a[key], b[key]), (key, a[key], b[key])


def running_total_between():
    spec = {"A1": "amount", "B1": "tax", "C1": "running", "D1": "double", "E1": "plus"}
    for r in range(2, 10):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*1.1", 0.0)
        spec[f"C{r}"] = (f"=MAX($B$2:B{r})", 0.0)  # cell mode: a growing range outside SUM
        spec[f"D{r}"] = (f"=C{r}*2", 0.0)  # reads the cell-mode column
        spec[f"E{r}"] = (f"=D{r}+1", 0.0)
    return book_of({"S": spec})


def test_a_running_total_between_region_columns_splits_the_region():
    book = running_total_between()
    compiled = compile_book(book)
    assert region_cols(compiled) == [[2], [4, 5]]  # B alone, then D and E after the running total
    assert any("split" in n for n in compiled.notes)
    check_equal(book, compiled)


def share_of_total():
    spec = {"A1": "amount", "B1": "tax", "C1": "share"}
    for r in range(2, 10):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*1.1", 0.0)
        spec[f"C{r}"] = (f"=B{r}/$B$10", 0.0)  # reads the total, which reads column B
    spec["B10"] = ("=SUM(B2:B9)", 0.0)
    return book_of({"S": spec})


def test_a_share_of_total_column_splits_the_region():
    book = share_of_total()
    compiled = compile_book(book)
    assert region_cols(compiled) == [[2], [3]]
    check_equal(book, compiled)


def test_a_split_part_that_still_cycles_stays_in_cell_mode_and_the_pure_part_is_kept():
    spec = {}
    for r in range(2, 10):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)  # pure
        spec[f"C{r}"] = (f"=MAX($A$2:A{r})", 0.0)  # cell mode
        spec[f"D{r}"] = (f"=$Z$1+A{r}", 0.0)  # reads Z1
        spec[f"E{r}"] = (f"=C{r}+1", 0.0)  # reads the cell-mode column
    spec["Z1"] = ("=E3+1", 0.0)  # D -> Z1 -> E: a cycle if D and E are one block, none for cells
    book = book_of({"S": spec})
    compiled = compile_book(book)
    assert region_cols(compiled) == [[2]]
    check_equal(book, compiled)


def test_split_regions_build_in_dbt(tmp_path):
    if shutil.which("dbt") is None:
        pytest.skip("needs: pip install -e .[dbt]")  # the model-run harness in test_emitter covers the layout
    wb = tmp_path / "split.xlsx"
    book = running_total_between()
    saved = execute(book, compile_book(book, regions=False))  # the answers, so every cell must PASS
    spec = {}
    for key, cell in book.cells.items():
        ref = f"{chr(64 + key[2])}{key[1]}"
        spec[ref] = ("=" + cell.formula, saved[key][1]) if cell.formula else cell.value
    write_xlsx(wb, {"S": spec})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out"), "--dbt"]) == 0  # no FAIL and dbt build ok
    names = {p.name for p in (tmp_path / "out" / "models" / "unexcel_cells").iterdir()}
    assert sum(n.startswith("region_") for n in names) == 2


def test_a_column_reading_a_cell_that_reads_the_region_splits_off():
    spec = {}
    for r in range(2, 5):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=$Z$1+B{r}", 0.0)
    spec["Z1"] = ("=B3+1", 0.0)
    book = book_of({"S": spec})
    compiled = compile_book(book)
    assert region_cols(compiled) == [[2], [3]]
    check_equal(book, compiled)
