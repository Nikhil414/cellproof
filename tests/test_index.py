import time

from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.refs import Area


def test_present_lists_only_filled_cells_row_major():
    book = book_of({"S": {"B2": 1, "A3": 2, "B1": 3, "C9": 4}})
    assert Area("S", 1, 1, 5, 2).present(book) == [("S", 1, 2), ("S", 2, 2), ("S", 3, 1)]
    assert Area("S", 1, 3, 1_048_576, 3).present(book) == [("S", 9, 3)]
    assert Area("Other", 1, 1, 9, 9).present(book) == []


def test_whole_column_refs_on_a_sheet_formatted_to_the_last_row_stay_fast():
    spec = {f"A{r}": r for r in range(1, 201)}
    spec.update({f"B{r}": ("=SUM(A:A)", 20100) for r in range(1, 201)})
    book = book_of({"S": spec})
    book.max_row["S"] = 1_048_576  # someone formatted the whole sheet, so A:A ends at the last row
    start = time.perf_counter()
    compiled = compile_book(book)
    assert compiled.unresolved == {}
    assert time.perf_counter() - start < 5  # Plan 1 walked 1,048,576 rows per reference: minutes
