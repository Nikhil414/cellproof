from xlsxkit import SALES, book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute


def run(sheets):
    book = book_of(sheets)
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def test_sumifs_whole_columns_case_insensitive_and_skips_text():
    got, c = run({"Sales": SALES, "Commission": {
        "A1": "Asha", "B1": ("=SUMIFS(Sales!C:C, Sales!A:A, A1)", 120000),
        "A2": "Vikram", "B2": ('=SUMIFS(Sales!C:C, Sales!A:A, A2, Sales!B:B, "south")', 60000),
    }})
    assert c.unresolved == {}
    assert got[("Commission", 1, 2)] == ("n", 120000.0, None)
    assert got[("Commission", 2, 2)] == ("n", 60000.0, None)
    assert "Sales!C5: number stored as text ('15,000') is skipped by Excel" in c.findings


def test_sumifs_rejects_what_it_cannot_prove():
    _, c = run({"Sales": SALES, "X": {
        "A1": ('=SUMIFS(Sales!C:C, Sales!A:A, ">5")', 0),
        "A2": ('=SUMIFS(Sales!C2:C5, Sales!A2:A3, "Asha")', 0),
        "A3": ("=SUMIFS(Sales!C:C, Sales!A:A)", 0),
    }})
    assert ("X", 1, 1) not in c.unresolved  # comparison criteria are supported since Plan 2b
    assert "same height" in c.unresolved[("X", 2, 1)]
    assert "criteria pairs" in c.unresolved[("X", 3, 1)]


def test_vlookup_exact_with_header_row_and_missing_key():
    rates = {"A1": "Region", "B1": "Rate", "A2": "West", "B2": 0.05, "A3": "South", "B3": 0.02}
    got, c = run({"Rates": rates, "X": {
        "A1": "south", "B1": ("=VLOOKUP(A1, Rates!A:B, 2, FALSE)", 0.02),
        "A2": "North", "B2": ("=VLOOKUP(A2, Rates!A:B, 2, FALSE)", "#N/A"),
        "B3": ("=VLOOKUP(A1, Rates!A:B, 2)", 0.02),
        "B4": ("=VLOOKUP(A1, Rates!A:B, 3, FALSE)", "#REF!"),
    }})
    assert got[("X", 1, 2)] == ("n", 0.02, None)
    assert got[("X", 2, 2)] == ("e", None, "#N/A")
    # Region, West, South is not in order, so Excel's approximate match depends on its binary search: not guessed
    assert "ascending sorted" in c.unresolved[("X", 3, 2)]
    assert c.unresolved[("X", 4, 2)] == "VLOOKUP column index outside the table"
