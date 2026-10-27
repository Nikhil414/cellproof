from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute

# Rates: brackets sorted ascending (A), rate (B), label (C, one blank)
DATA = {"A1": 0, "B1": 0.1, "C1": "low", "A2": 1000, "B2": 0.2, "C2": "mid", "A3": 5000, "B3": 0.3,
        "A4": "x", "B4": "#N/A", "E1": "Jan", "F1": "Feb", "G1": "Mar", "E2": 1, "F2": 2, "G2": 3,
        "I1": 5000, "I2": 1000, "I3": 0,
        "J1": 1000, "K1": 0.1, "J2": 0, "K2": 0.2, "J3": 5000, "K3": 0.3}  # brackets not in order


def run(formulas):
    book = book_of({"R": DATA, "F": formulas})
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def run_with_book(sheets):
    book = book_of(sheets)
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


def test_approximate_lookup_on_unsorted_data_follows_excel_s_search_path():
    # Excel's binary search over (1000, 0, 5000) for 2500: probe 0 (<=2500, go right), probe 5000 (>2500, go left)
    # -> the 0 at row 2, not the true bracket 1000 at row 1. The search path, not the sorted answer.
    got, c = run({"A1": ("=VLOOKUP(2500, R!J1:K3, 2)", 0.2), "A2": ("=MATCH(2500, R!J1:J3, 1)", 2),
                  "A3": ("=LOOKUP(2500, R!J1:J3, R!K1:K3)", 0.2), "A4": ("=MATCH(-1, R!J1:J3, 1)", "#N/A"),
                  "A5": ("=MATCH(9.99E+307, R!J1:J3, 1)", 3)})
    assert c.unresolved == {}
    assert [got[("F", r, 1)] for r in range(1, 6)] == [
        N(0.2), N(2), N(0.2), ("e", None, "#N/A"), N(3)]


def test_approximate_match_type_minus_1_on_unsorted_data_stays_unresolved():
    # MATCH(999, (0,1000,5000), -1) is #N/A in Excel, MATCH(150, (0,100,500), -1) is #N/A too, but the rule behind
    # the two is not derivable from Excel's answers (LibreOffice's order-broken bail-out differs), so no guess.
    _, c = run({"A1": ("=MATCH(2500, R!J1:J3, -1)", "#N/A")})
    assert "descending sorted" in c.unresolved[("F", 1, 1)]


def test_the_walk_needs_the_table_written_as_inputs():
    data = dict(DATA)
    data["J3"] = ("=5000", 5000.0)
    _, c = run_with_book({"R": data, "F": {"A1": ("=VLOOKUP(2500, R!J1:K3, 2)", 0.2)}})
    assert "written as inputs" in c.unresolved[("F", 1, 1)]


def test_the_walk_refuses_empty_cells_in_the_table():
    data = dict(DATA)
    del data["J2"]
    data["J3"] = 0  # (1000, empty, 0): not in order any more, so the search path would have to probe the empty cell
    _, c = run_with_book({"R": data, "F": {"A1": ("=VLOOKUP(2500, R!J1:K3, 2)", 0.2)}})
    assert "empty cell" in c.unresolved[("F", 1, 1)]


def test_the_compiled_search_refuses_changed_table_values():
    from cellproof.verifier import verify
    book = book_of({"R": DATA, "F": {"A1": ("=VLOOKUP(2500, R!J1:K3, 2)", 0.2)}})
    compiled = compile_book(book)
    changed = dict(DATA)
    changed["J2"] = 7.0  # the table the SQL was compiled from no longer matches the data it runs on
    verdicts = verify(book_of({"R": changed, "F": {"A1": ("=VLOOKUP(2500, R!J1:K3, 2)", 0.2)}}), compiled)
    assert verdicts[0].status == "UNRESOLVED" and "changed after" in verdicts[0].reason


def test_walk_too_large_table_stays_unresolved():
    data = {f"J{r}": float(1200 - r) for r in range(1, 1200)}  # descending: not in order for a type-1 search
    _, c = run_with_book({"R": data, "F": {"A1": ("=MATCH(5, R!J1:J1199, 1)", 5)}})
    assert "search is too large" in c.unresolved[("F", 1, 1)]


def test_the_walk_matches_excel_s_bisection_on_random_tables():
    # The compiled decision tree must be exactly Excel's bisection: probe (lo+hi)//2, go right while the value is
    # <= x, answer the last probe it moved right from, #N/A when there is none. Differential against the same walk
    # in Python; Excel's own answers are pinned by the tests above and in test_excel_findings.py.
    import random

    def bisection(vec, x):
        lo, hi, best = 0, len(vec) - 1, None
        while lo <= hi:
            mid = (lo + hi) // 2
            if vec[mid] <= x:
                best, lo = mid, mid + 1
            else:
                hi = mid - 1
        return best

    rng = random.Random(7)
    for _ in range(8):
        n = rng.randrange(2, 40)
        vec = [float(rng.randrange(0, 100)) for _ in range(n)]
        xs = [float(rng.randrange(-5, 105)) for _ in range(6)]
        rows = {f"A{r}": v for r, v in enumerate(vec, 1)}
        forms = {f"A{i}": (f"=MATCH({x!r}, S!A1:A{n}, 1)", 0) for i, x in enumerate(xs, 1)}
        got, c = run_with_book({"S": rows, "F": forms})
        assert c.unresolved == {}
        for i, x in enumerate(xs, 1):
            want = bisection(vec, x)
            kind, num, txt = got[("F", i, 1)]
            assert (kind == "e" and txt == "#N/A") if want is None else num == want + 1
