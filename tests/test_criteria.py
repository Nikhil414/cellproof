from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute

# A1:A9 - 5, "5", apple, Banana, TRUE, (blank), "", a*b, #N/A ; B: amounts ; C: regions
DATA = {
    "A1": 5, "A2": "5", "A3": "apple", "A4": "Banana", "A5": True, "A7": "", "A8": "a*b", "A9": "#N/A",
    "B1": 10, "B2": 20, "B3": 30, "B4": 40, "B5": 50, "B6": 60, "B7": 70, "B8": 80, "B9": 90,
    "C1": "W", "C2": "E", "C3": "W", "C4": "E", "C5": "W", "C6": "E", "C7": "W", "C8": "E", "C9": "W",
    "D1": ">", "D2": 4,
}


def run(formulas):
    book = book_of({"S": DATA, "F": formulas})
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def N(v):
    return ("n", float(v), None)


def test_countif_criteria_forms():
    # Real Excel (examples/functions.xlsx): "<>5" skips only the number 5, so the text "5" is counted (9 - 1 = 8);
    # "<>" counts every existing cell, so the empty text in A7 is counted (8, not 7).
    cases = {'"5"': 2, '">4"': 1, '"<>5"': 8, '"a*"': 2, '"a~*b"': 1, '">b"': 1, '"TRUE"': 1, '""': 2, '"="': 2,
             '"<>"': 8, '"?pple"': 1, '"%"': 0, 'S!D1&S!D2': 1, '"<=5"': 1, "5": 2, "TRUE": 1}
    formulas = {f"A{i}": (f"=COUNTIF(S!A1:A9, {crit})", n) for i, (crit, n) in enumerate(cases.items(), 1)}
    got, c = run(formulas)
    assert c.unresolved == {}
    assert [got[("F", i, 1)] for i in range(1, len(cases) + 1)] == [N(n) for n in cases.values()]


def test_countifs_counts_blank_rows_too():
    got, c = run({"A1": ('=COUNTIFS(S!A1:A9, "", S!C1:C9, "E")', 1),  # A6 blank & E
                  "A2": ('=COUNTIFS(S!A1:A9, "<>5", S!C1:C9, "W")', 4)})
    assert c.unresolved == {}
    assert got[("F", 1, 1)] == N(1) and got[("F", 2, 1)] == N(4)


def test_sumif_averageif_and_the_ifs_family():
    got, c = run({
        "A1": ('=SUMIF(S!C1:C9, "W", S!B1:B9)', 250), "A2": ('=SUMIF(S!C1:C9, "W", S!B1)', 250),  # resized
        "A3": ('=SUMIFS(S!B1:B9, S!C1:C9, "E", S!B1:B9, ">"&S!D2*10)', 140),
        "A4": ('=AVERAGEIF(S!C1:C9, "E", S!B1:B9)', 50), "A5": ('=AVERAGEIFS(S!B1:B9, S!C1:C9, "X")', "#DIV/0!"),
        "A6": ('=MINIFS(S!B1:B9, S!C1:C9, "W")', 10), "A7": ('=MAXIFS(S!B1:B9, S!C1:C9, "W")', 90),
        "A8": ('=SUMIF(S!B1:B9, ">=50")', 350),
    })
    assert c.unresolved == {}
    assert [got[("F", r, 1)] for r in range(1, 9)] == [
        N(250), N(250), N(140), N(50), ("e", None, "#DIV/0!"), N(10), N(90), N(350)]


def test_a_matched_error_makes_the_sum_an_error():
    got, c = run({"A1": ('=SUMIF(S!C1:C2, "W", S!A1:A2)', 5), "A2": ('=SUMIF(S!C8:C9, "W", S!A8:A9)', "#N/A")})
    assert c.unresolved == {}
    assert got[("F", 1, 1)] == N(5)  # A2 is the text "5": skipped, like Excel
    assert got[("F", 2, 1)] == ("e", None, "#N/A")


def test_criteria_from_an_empty_cell_is_unresolved():
    _, c = run({"A1": ("=COUNTIF(S!A1:A9, S!Z99)", 0)})
    assert c.unresolved[("F", 1, 1)] == "criteria from an empty cell not supported yet"
