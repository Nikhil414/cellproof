from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute


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


def test_let_binds_names_to_values_ranges_and_earlier_names():
    from cellproof.compiler import compile_book
    from cellproof.verifier import execute
    from xlsxkit import book_of
    s = {"A1": 2.0, "A2": 3.0, "A3": 4.0, "B1": "x",
         # as a file stores them: _xlfn.LET, names prefixed _xlpm.
         "C1": ("=_xlfn.LET(_xlpm.x,A1*10,_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y)", 420.0),
         "C2": ("=_xlfn.LET(_xlpm.r,A1:A3,SUM(_xlpm.r)/COUNT(_xlpm.r))", 3.0),
         "C3": ('=_xlfn.LET(_xlpm.n,B1,IF(_xlpm.n="x","yes","no"))', "yes"),
         "C4": ("=_xlfn.LET(_xlpm.x,1,_xlfn.LET(_xlpm.x,_xlpm.x+5,_xlpm.x*2)+_xlpm.x)", 13.0),  # inner shadows
         "C5": ("=_xlfn.LET(_xlpm.bad,1/0,_xlpm.ok,7,_xlpm.ok)", 7.0),  # an unused name's error never shows
         "C6": ("=_xlfn.LET(_xlpm.d,C1,_xlpm.d+1)", 421.0),  # waits for C1
         "C7": ("=_xlfn.LET(_xlpm.x,1)", 0.0), "C8": ("=_xlfn.LET(A1,1,A1)", 0.0)}
    book = book_of({"S": s})
    compiled = compile_book(book)
    got = execute(book, compiled)
    assert [got[("S", r, 3)] for r in range(1, 7)] == [
        ("n", 420.0, None), ("n", 3.0, None), ("s", None, "yes"), ("n", 13.0, None), ("n", 7.0, None),
        ("n", 421.0, None)]
    assert "LET needs" in compiled.unresolved[("S", 7, 3)] and "LET name" in compiled.unresolved[("S", 8, 3)]


def test_let_pasted_down_equals_cell_mode():
    from cellproof.compiler import compile_book
    from cellproof.verifier import execute
    from xlsxkit import book_of
    s = {}
    for r in range(1, 30):
        s[f"A{r}"], s[f"B{r}"] = float(r), float(r % 7)
        s[f"C{r}"] = (f"=_xlfn.LET(_xlpm.t,A{r}*B{r},IF(_xlpm.t>20,_xlpm.t,-_xlpm.t))", 0.0)
    book = book_of({"S": s})
    regions, cells = compile_book(book), compile_book(book, regions=False)
    assert regions.unresolved == cells.unresolved == {}
    assert execute(book, regions) == execute(book, cells)
