from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import verify


def _running_formula(n):
    spec, total = {}, 0.0
    for r in range(1, n + 1):
        b = (r % 97) * 1.5
        total += b
        spec[f"A{r}"] = float(r % 97)
        spec[f"B{r}"] = (f"=A{r}*1.5", b)
        spec[f"C{r}"] = (f"=SUM($B$1:B{r})", total)
    return book_of({"S": spec})


def test_running_sum_over_a_formula_column_does_not_make_a_batch_per_row():
    n = 300
    book = _running_formula(n)
    compiled = compile_book(book)
    assert len(compiled.batches) < n // 4  # 301 batches before Plan 2c
    assert {v.status for v in verify(book, compiled)} == {"PASS"}


def test_number_stored_as_text_in_a_summed_range_is_reported():  # Review Focus 5
    book = book_of({"S": {"A1": "15,000", "A2": 5.0, "B1": ("=SUM(A1:A2)", 5.0)}})
    compiled = compile_book(book)
    assert any("S!A1" in f and "number stored as text" in f for f in compiled.findings)
    assert {v.status for v in verify(book, compiled)} == {"PASS"}
