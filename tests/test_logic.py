from xlsxkit import book_of

from unexcel.compiler import compile_book
from unexcel.verifier import execute


def run(cells):
    book = book_of({"S": cells})
    compiled = compile_book(book)
    return execute(book, compiled), compiled


def B(v):
    return ("b", 1.0 if v else 0.0, None)


def test_ifs_switch_not_xor():
    got, c = run({"A1": 75, "A2": 0,
                  "B1": ('=IFS(A1>=90, "A", A1>=70, "B", TRUE, "C")', "B"), "B2": ('=IFS(A1>100, "X")', "#N/A"),
                  "B3": ('=IFS(1/A2>1, "X", TRUE, "Y")', "#DIV/0!"), "B4": ('=SWITCH(A1, 50, "half", 75, "most", "?")', "most"),
                  "B5": ('=SWITCH(A1, 1, "one")', "#N/A"), "B6": ('=SWITCH("b", "A", 1, "B", 2)', 2),
                  "B7": ("=NOT(A1>5)", False), "B8": ("=XOR(TRUE, A1>5, FALSE)", False), "B9": ("=XOR(TRUE, A2>5)", True)})
    assert c.unresolved == {}
    assert [got[("S", r, 2)] for r in range(1, 10)] == [
        ("s", None, "B"), ("e", None, "#N/A"), ("e", None, "#DIV/0!"), ("s", None, "most"), ("e", None, "#N/A"),
        ("n", 2.0, None), B(False), B(False), B(True)]


def test_ifna_and_the_is_functions():
    got, c = run({"A1": 0, "A2": "5", "A3": ("=1/A1", "#DIV/0!"), "A4": ("=NA()", "#N/A"), "A5": ('=IF(A1=0, "", 1)', ""),
                  "B1": ('=IFNA(A4, "gone")', "gone"), "B2": ('=IFNA(A3, "gone")', "#DIV/0!"),
                  "B3": ("=ISNA(A4)", True), "B4": ("=ISNA(A3)", False), "B5": ("=ISERROR(A3)", True),
                  "B6": ("=ISERR(A4)", False), "B7": ("=ISERR(A3)", True), "B8": ("=ISNUMBER(A2)", False),
                  "B9": ("=ISNUMBER(A1)", True), "B10": ("=ISTEXT(A5)", True), "B11": ("=ISBLANK(Z9)", True),
                  "B12": ("=ISBLANK(A5)", False), "B13": ("=ISLOGICAL(TRUE())", True), "B14": ("=ISNUMBER(A3)", False)})
    assert c.unresolved == {}
    assert [got[("S", r, 2)] for r in range(1, 15)] == [
        ("s", None, "gone"), ("e", None, "#DIV/0!"), B(True), B(False), B(True), B(False), B(True), B(False),
        B(True), B(True), B(True), B(False), B(True), B(False)]
