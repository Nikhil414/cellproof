"""PMT, PV, FV, NPER, NPV; LARGE, SMALL, RANK, STDEV/VAR, PERCENTILE, QUARTILE; NETWORKDAYS, WORKDAY.

Expected values are the examples of Microsoft's function reference (rounded as printed there, so compared to that
precision); exact Excel answers for PV, NPV and PMT come from the public corpus (tests/corpus.py).
Edge cases whose Excel answer was not checked (zero periods, rate <= -100%, payment type other than 0/1, a k that
is not whole) are the UNCHECKED marker, which the verifier reports as UNRESOLVED - never a guess.
"""
import math

from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.expr import UNCHECKED
from cellproof.verifier import execute

DATA = {  # STDEV example column, the RANK example, the PERCENTILE/QUARTILE examples, the holidays of the docs
    **{f"A{i}": v for i, v in enumerate([1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299], 1)},
    **{f"B{i}": v for i, v in enumerate([7, 3.5, 3.5, 1, 2], 1)},
    **{f"C{i}": v for i, v in enumerate([1, 3, 2, 4], 1)},
    **{f"D{i}": v for i, v in enumerate([1, 2, 4, 7, 8, 9, 10, 12], 1)},
    **{f"E{i}": v for i, v in enumerate([3, 5, 3, 5, 4, 4, 2, 4, 6, 7], 1)},
    "F1": 41235, "F2": 41247, "F3": 41295, "G1": 39778, "G2": 39786, "G3": 39834,
    "H1": -10000, "H2": 3000, "H3": 4200, "H4": 6800, "H5": "note", "H7": "#N/A",
}


def run(cases):
    book = book_of({"D": DATA, "F": {f"A{i}": (f, 0.0) for i, (f, _) in enumerate(cases, 1)}})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert compiled.unresolved == {}, compiled.unresolved
    for i, (formula, want) in enumerate(cases, 1):
        kind, n, t = got[("F", i, 1)]
        if isinstance(want, str):
            assert (kind, t) == ("e", want), (formula, kind, n, t)
        else:
            value, places = want if isinstance(want, tuple) else (want, 9)
            assert kind == "n" and abs(n - value) <= 0.5 * 10 ** -places + 1e-12 * abs(value), (formula, n)


def test_finance_functions_match_microsofts_examples():
    run([("=PMT(0.08/12,10,10000)", (-1037.03, 2)), ("=PMT(0.08/12,10,10000,0,1)", (-1030.16, 2)),
         ("=PMT(0.06/12,18*12,0,50000)", (-129.08, 2)), ("=PMT(0,10,1000)", -100),
         ("=FV(0.06/12,10,-200,-500,1)", (2581.40, 2)), ("=FV(0.12/12,12,-1000)", (12682.50, 2)),
         ("=FV(0.11/12,35,-2000,,1)", (82846.25, 2)), ("=FV(0,10,-100,-50)", 1050),
         ("=PV(0.08/12,12*20,500,,0)", (-59777.15, 2)), ("=PV(0,10,-100)", 1000),
         ("=NPER(0.12/12,-100,-1000,10000,1)", (59.6738657, 7)), ("=NPER(0.12/12,-100,-1000)", (-9.57859404, 8)),
         ("=NPER(0.12/12,-100,-1000,10000)", (60.0821229, 7)), ("=NPER(0,-100,1000)", 10),
         ("=NPER(0.1,-50,1000)", "#NUM!"),
         ("=NPV(0.1,-10000,3000,4200,6800)", (1188.44, 2)), ("=NPV(0.1,D!H1:H4)", (1188.44, 2)),
         ("=NPV(0.1,D!H1:H2,D!H3:H6)", (1188.44, 2)),  # text and empty cells in a reference are skipped
         ("=NPV(0.1,D!H1:H7)", "#N/A")])


def test_finance_edge_cases_excel_was_not_asked_about_are_unchecked():
    run([("=PMT(0.1,0,1000)", UNCHECKED), ("=PMT(-1,10,1000)", UNCHECKED), ("=PMT(0.1,10,1000,0,2)", UNCHECKED),
         ("=FV(-2,10,100)", UNCHECKED), ("=NPER(0,0,1000)", UNCHECKED), ("=NPV(-1,1,2)", UNCHECKED),
         ("=PMT(1/0,10,1000)", "#DIV/0!")])


def test_ranking_functions_match_microsofts_examples():
    run([("=LARGE(D!E1:E10,3)", 5), ("=LARGE(D!E1:E10,7)", 4), ("=SMALL(D!E1:E10,4)", 4),
         ("=SMALL(D!B1:B5,2)", 2), ("=LARGE(D!H1:H5,1)", 6800),  # text skipped
         ("=LARGE(D!E1:E10,0)", "#NUM!"), ("=SMALL(D!E1:E10,11)", "#NUM!"), ("=LARGE(D!H1:H7,1)", "#N/A"),
         ("=LARGE(D!E1:E10,1.5)", UNCHECKED), ("=SMALL(D!H1:H7,1/0)", "#DIV/0!"),
         ("=RANK(D!B3,D!B1:B5,1)", 3), ("=RANK(D!B1,D!B1:B5)", 1), ("=_xlfn.RANK.EQ(2,D!B1:B5,0)", 4),
         ("=RANK(3.5,D!B1:B5)", 2), ("=RANK(5,D!B1:B5)", "#N/A"), ("=RANK(1,D!H1:H7)", UNCHECKED)])


def test_spread_functions_match_microsofts_examples():
    run([("=STDEV(D!A1:A10)", (27.46392, 5)), ("=_xlfn.STDEV.S(D!A1:A10)", (27.46392, 5)),
         ("=_xlfn.STDEV.P(D!A1:A10)", (26.05456, 5)), ("=STDEVP(D!A1:A10)", (26.05456, 5)),
         ("=VAR(D!A1:A10)", (754.2667, 4)), ("=_xlfn.VAR.S(D!A1:A10)", (754.2667, 4)),
         ("=_xlfn.VAR.P(D!A1:A10)", (678.84, 2)), ("=VARP(D!A1:A10)", (678.84, 2)),
         ("=STDEV(2,4,4,4,5,5,7,9)", math.sqrt(32 / 7)), ("=_xlfn.STDEV.P(2,4,4,4,5,5,7,9)", 2),
         ("=STDEV(D!H1:H5)", math.sqrt(sum((x - 1000) ** 2 for x in (-10000, 3000, 4200, 6800)) / 3)),
         ("=STDEV(D!A1:A10*2)", (2 * 27.46392, 4)), ("=STDEV(5)", "#DIV/0!"), ("=_xlfn.VAR.P(D!H5:H6)", "#DIV/0!"),
         ("=STDEV(D!H1:H7)", "#N/A")])


def test_percentiles_match_microsofts_examples():
    run([("=PERCENTILE(D!C1:C4,0.3)", 1.9), ("=_xlfn.PERCENTILE.INC(D!C1:C4,0.3)", 1.9),
         ("=PERCENTILE(D!C1:C4,1)", 4), ("=PERCENTILE(D!C1:C4,0)", 1),
         ("=QUARTILE(D!D1:D8,1)", 3.5), ("=_xlfn.QUARTILE.INC(D!D1:D8,3)", 9.25), ("=QUARTILE(D!D1:D8,2.9)", 7.5),
         ("=QUARTILE(D!D1:D8,4)", 12), ("=QUARTILE(D!D1:D8,5)", "#NUM!"), ("=PERCENTILE(D!C1:C4,1.1)", "#NUM!"),
         ("=PERCENTILE(D!H5:H6,0.5)", "#NUM!"), ("=PERCENTILE(D!H1:H7,0.5)", "#N/A"),
         # formulas' test.xlsx (Excel-saved): a negative quart is #NUM! before truncating; k's error comes first
         ("=QUARTILE(D!D1:D8,-0.1)", "#NUM!"), ("=PERCENTILE(D!H1:H7,1/0)", "#DIV/0!")])


def test_working_days_match_microsofts_examples():
    run([("=NETWORKDAYS(41183,41334)", 110), ("=NETWORKDAYS(41183,41334,D!F1)", 109),
         ("=NETWORKDAYS(41183,41334,D!F1:F3)", 107), ("=NETWORKDAYS(41334,41183)", -110),
         ("=NETWORKDAYS(41183,41183)", 1), ("=NETWORKDAYS(41188,41189)", 0),  # a Saturday and a Sunday
         ("=NETWORKDAYS(-1,41183)", "#NUM!"),
         ("=WORKDAY(39722,151)", 39933), ("=WORKDAY(39722,151,D!G1:G3)", 39938),
         ("=WORKDAY(39933,-151)", 39722), ("=WORKDAY(39938,-151,D!G1:G3)", 39722),
         ("=WORKDAY(41188,0)", 41188), ("=WORKDAY(41188,1)", 41190), ("=WORKDAY(41183,1.9)", 41184),
         ("=WORKDAY(2958465,5)", "#NUM!")])


def test_working_days_against_a_day_by_day_count():
    # every start weekday x every span: the closed form must equal counting day by day (serial mod 7: 0 Sat, 1 Sun)
    cases = []
    for start in range(41183, 41190):
        for span in (0, 1, 4, 5, 6, 7, 13, 30):
            want = sum(1 for d in range(start, start + span + 1) if d % 7 not in (0, 1))
            cases.append((f"=NETWORKDAYS({start},{start + span})", want))
    run(cases)


def test_pasted_down_columns_equal_cell_mode():
    s = {"A1": "rate", "B1": "days", "C1": "start", "G1": "hol"}
    for r in range(2, 14):
        s[f"A{r}"], s[f"B{r}"], s[f"C{r}"] = 0.01 * (r - 7), float(r * 9 - 50), float(41180 + r)
        for col, f in zip("DEFHIJKLMN", ["=PMT(A{r},12,1000)", "=WORKDAY(C{r},B{r},$G$2:$G$4)",
                                          "=NETWORKDAYS(C{r},C{r}+B{r},$G$2:$G$4)", "=LARGE($B$2:$B$13,r-1)",
                                          "=RANK(B{r},$B$2:$B$13)", "=STDEV(A{r}:C{r})", "=NPV(A{r},B{r},C{r})",
                                          "=PERCENTILE($B$2:$B$13,A{r}+0.06)", "=NPER(A{r},-100,1000)",
                                          "=FV(A{r},B{r},-10)"]):
            s[f"{col}{r}"] = (f.replace("{r}", str(r)).replace("r-1", str(r - 1)), 0.0)
    s["G2"], s["G3"], s["G4"] = 41190.0, 41201.0, 41250.0
    book = book_of({"S": s})
    regions, cells = compile_book(book), compile_book(book, regions=False)
    assert regions.unresolved == cells.unresolved
    assert execute(book, regions) == execute(book, cells)
    assert len(cells.unresolved) < 12, cells.unresolved  # only the UNCHECKED-free formulas compile everywhere
