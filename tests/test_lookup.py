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


def test_lookup_vector_form_largest_value_not_above():
    grid = {"A1": 1, "A2": 3, "A3": 3, "A4": 7, "B1": "one", "B2": "three", "B3": "THREE", "B4": "seven",
            "D1": 10, "E1": 30, "F1": 30, "G1": 70}
    got, c = run({"S": grid, "X": {
        "A1": ("=LOOKUP(5, S!A1:A4, S!B1:B4)", "THREE"),   # last of equal values, as Excel's binary search
        "A2": ("=LOOKUP(0, S!A1:A4, S!B1:B4)", "#N/A"),    # smaller than every value
        "A3": ("=LOOKUP(99, S!A1:A4)", 7),                 # no result vector: the value found
        "A4": ("=LOOKUP(4, S!A1:A4, S!D1:G1)", 30),        # result vector in the other direction
    }})
    assert c.unresolved == {}
    assert got[("X", 1, 1)] == ("s", None, "THREE")
    assert got[("X", 2, 1)] == ("e", None, "#N/A")
    assert got[("X", 3, 1)] == ("n", 7.0, None)
    assert got[("X", 4, 1)] == ("n", 30.0, None)


def test_lookup_array_form_and_refusals():
    wide = {"A1": "a", "B1": "b", "C1": "c", "A2": 1, "B2": 2, "C2": 3,
            "E1": 1, "E2": 5, "E3": 2, "F1": 10}
    got, c = run({"S": wide, "X": {
        "A1": ('=LOOKUP("B", S!A1:C2)', 2),                 # wider than tall: first row searched, last row returned
        "A2": ("=LOOKUP(3, S!E1:E3, S!F1:F3)", 0),          # unsorted: Excel's answer depends on its search
        "A3": ("=LOOKUP(1, S!E1:E2, S!F1:F3)", 10),         # result vector of another size
    }})
    assert got[("X", 1, 1)] == ("n", 2.0, None)
    assert "ascending sorted" in c.unresolved[("X", 2, 1)]
    assert "same size" in c.unresolved[("X", 3, 1)]
