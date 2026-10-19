from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute

# Rates: brackets sorted ascending (A), rate (B), label (C, one blank)
DATA = {"A1": 0, "B1": 0.1, "C1": "low", "A2": 1000, "B2": 0.2, "C2": "mid", "A3": 5000, "B3": 0.3,
        "A4": "x", "B4": "#N/A", "E1": "Jan", "F1": "Feb", "G1": "Mar", "E2": 1, "F2": 2, "G2": 3,
        "I1": 5000, "I2": 1000, "I3": 0}


def run(formulas):
    book = book_of({"R": DATA, "F": formulas})
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def N(v):
    return ("n", float(v), None)


def test_approximate_lookups_find_the_bracket():
    got, c = run({"A1": ("=VLOOKUP(2500, R!A1:B3, 2)", 0.2), "A2": ("=VLOOKUP(2500, R!A1:B3, 2, TRUE)", 0.2),
                  "A3": ("=MATCH(5000, R!A1:A3)", 3), "A4": ("=MATCH(-1, R!A1:A3, 1)", "#N/A"),
                  "A5": ("=MATCH(999, R!A1:A3, -1)", "#N/A"), "A6": ("=MATCH(999, R!I1:I3, -1)", 2)})
    # Real Excel: -1 assumes descending data, so on the ascending A1:A3 it answers #N/A. cellproof will not guess
    # what a binary search does on data out of order (A5); on descending data (I1:I3) it finds the bracket (A6).
    assert list(c.unresolved) == [("F", 5, 1)] and "descending sorted" in c.unresolved[("F", 5, 1)]
    assert [got[("F", r, 1)] for r in (1, 2, 3, 4, 6)] == [N(0.2), N(0.2), N(3), ("e", None, "#N/A"), N(2)]


def test_exact_lookups_wildcards_and_errors():
    got, c = run({"A1": ('=MATCH("m*", R!C1:C3, 0)', 2), "A2": ('=MATCH("Feb", R!E1:G1, 0)', 2),
                  "A3": ('=HLOOKUP("mar", R!E1:G2, 2, FALSE)', 3), "A4": ('=VLOOKUP("x", R!A1:B4, 2, FALSE)', "#N/A"),
                  "A5": ('=XLOOKUP("mid", R!C1:C3, R!A1:A3)', 1000), "A6": ('=XLOOKUP("m*", R!C1:C3, R!A1:A3)', "#N/A"),
                  "A7": ('=XLOOKUP("zz", R!C1:C3, R!A1:A3, "none")', "none"), "A8": ("=XLOOKUP(3000, R!A1:A3, R!B1:B3, 0, -1)", 0.2)})
    assert c.unresolved == {}
    assert [got[("F", r, 1)] for r in range(1, 9)] == [
        N(2), N(2), N(3), ("e", None, "#N/A"), N(1000), ("e", None, "#N/A"), ("s", None, "none"), N(0.2)]


def test_index_choose_and_mixed_results():
    got, c = run({"A1": ("=INDEX(R!A1:B3, 2, 2)", 0.2), "A2": ("=INDEX(R!E1:G1, 3)", "Mar"), "A3": ("=INDEX(R!A1:A3, 9)", "#REF!"),
                  "A4": ("=INDEX(R!C1:C3, 3)", 0), "A5": ('=INDEX(R!A1:A3, MATCH("mid", R!C1:C3, 0))', 1000),
                  "A6": ('=CHOOSE(2, "a", 5, TRUE)', 5), "A7": ('=CHOOSE(4, "a", 5)', "#VALUE!"),
                  "A8": ("=INDEX(R!C1:C3, 1)", "low")})
    assert c.unresolved == {}
    assert [got[("F", r, 1)] for r in range(1, 9)] == [
        N(0.2), ("s", None, "Mar"), ("e", None, "#REF!"), N(0), N(1000), N(5), ("e", None, "#VALUE!"), ("s", None, "low")]
