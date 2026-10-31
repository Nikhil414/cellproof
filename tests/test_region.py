import tempfile
from pathlib import Path

import duckdb
import pytest
from xlsxkit import book_of

import cellproof.compiler as compiler
from cellproof.compiler import READ_SEED, Compiler, compile_book, write_seed
from cellproof.model import Unsupported
from cellproof.region import compile_region, names_for, region_note
from cellproof.regions import find_regions
from cellproof.verifier import execute


def small_book():
    spec = {"A1": "amount", "B1": "tax", "C1": "big"}
    for r in range(2, 5):
        spec[f"A{r}"] = float(r * 40)
        spec[f"B{r}"] = (f"=A{r}*1.1", 0.0)
        spec[f"C{r}"] = (f"=IF(B{r}>100,B{r},0)", 0.0)
    return book_of({"S": spec})


GOLDEN_WIDE = """WITH
src AS (
  SELECT row_num,
    MAX(num) FILTER (WHERE col_num = 1) AS amount
  FROM __cellproof_cells__ WHERE sheet = 'S' AND row_num BETWEEN 2 AND 4 AND col_num IN (1) GROUP BY row_num),
tax AS (
  SELECT *, (_.amount * CAST(1.1 AS DOUBLE)) AS tax
  FROM src AS _),
big AS (
  SELECT *, (CASE WHEN (_.tax > CAST(100.0 AS DOUBLE)) THEN _.tax ELSE CAST(0.0 AS DOUBLE) END) AS big
  FROM tax AS _)
SELECT row_num, tax, big
FROM big"""


def test_the_wide_model_reads_as_plain_sql():  # the readability contract, reviewed by the user
    book = small_book()
    batch, rejected = compile_region(Compiler(book), find_regions(book).regions[0])
    assert rejected == {} and batch.wide_sql == GOLDEN_WIDE
    assert batch.model == "region_s_2_4"
    assert batch.unpivot("wide") == (
        "SELECT 'S', row_num, 2, 'n', CAST(tax AS DOUBLE), NULL FROM wide\n"
        "UNION ALL\n"
        "SELECT 'S', row_num, 3, 'n', CAST(big AS DOUBLE), NULL FROM wide")
    assert len(batch) == 6 and batch[0].key == ("S", 2, 2)


def run_region(book, region):
    """The region's rows next to what cell mode computes for the same cells."""
    comp = Compiler(book)
    batch, rejected = compile_region(comp, region)
    con = duckdb.connect()
    with tempfile.TemporaryDirectory() as tmp:
        seed = Path(tmp) / "s.csv"
        write_seed(book, seed)
        con.execute(f"CREATE TABLE cells AS SELECT * FROM {READ_SEED}", [str(seed)])
    if batch is not None:
        con.execute("INSERT INTO cells " + batch.render_rows("cells"))
    got = {(s, r, c): (k, n, t) for s, r, c, k, n, t in con.execute("SELECT * FROM cells").fetchall()}
    return got, execute(book, compile_book(book, regions=False)), batch, rejected


def same(a, b):
    return (a is not None and b is not None and a[0] == b[0] and a[2] == b[2]
            and (a[1] == b[1] or (a[1] is not None and b[1] is not None and abs(a[1] - b[1]) <= 1e-9 * max(1, abs(a[1])))))


def test_a_mixed_region_computes_what_cell_mode_computes():
    spec = {"A1": "amount", "B1": "tax", "C1": "net", "D1": "ratio", "E1": "divisor", "F1": "label", "G1": "flag",
            "H1": "mixed", "I1": "scaled", "J1": "rate", "K1": 0.5}
    for r in range(2, 14):
        spec[f"A{r}"] = float(r * 17 % 53)
        spec[f"E{r}"] = float(r % 3)  # zeros: #DIV/0!
        spec[f"B{r}"] = (f"=A{r}*1.1", 0.0)
        spec[f"C{r}"] = (f"=IF(B{r}>30,B{r},0)", 0.0)
        spec[f"D{r}"] = (f"=A{r}/E{r}", 0.0)
        spec[f"F{r}"] = (f'=IF(A{r}>20,"hi","lo")&"!"', "x")
        spec[f"G{r}"] = (f"=B{r}>30", True)
        spec[f"H{r}"] = (f'=IF(A{r}>25,A{r},"n/a")', 0.0)
        spec[f"I{r}"] = (f"=A{r}*$K$1+ROUND(B{r},1)", 0.0)
    book = book_of({"S": spec})
    region = find_regions(book).regions[0]
    got, cell_mode, batch, rejected = run_region(book, region)
    assert rejected == {} and len(batch) == len(region.keys())
    for key in region.keys():
        assert same(got[key], cell_mode[key]), (key, got[key], cell_mode[key])


def test_an_absolute_reference_into_the_regions_own_rows_is_one_scalar():  # Review Focus 2
    spec = {"A1": "a"}
    for r in range(2, 7):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*$A$3+A$2", 0.0)  # $A$3 and A$2 sit inside rows 2..6 but are not row-wise
    book = book_of({"S": spec})
    got, cell_mode, batch, rejected = run_region(book, find_regions(book).regions[0])
    assert rejected == {}
    assert all(same(got[("S", r, 2)], cell_mode[("S", r, 2)]) for r in range(2, 7))
    assert got[("S", 4, 2)][1] == 4.0 * 3.0 + 2.0


def test_an_input_column_with_a_blank_or_an_error_keeps_the_reading_column_in_cell_mode():  # Review Focus 1
    spec = {}
    for r in range(2, 7):
        spec[f"B{r}"] = float(r)
        spec[f"C{r}"] = (f"=B{r}*2", 0.0)
        spec[f"D{r}"] = (f"=A{r}*2", 0.0)
        spec[f"E{r}"] = (f"=F{r}+1", 0.0)
        spec[f"F{r}"] = float(r)
    spec.update({f"A{r}": float(r) for r in (2, 3, 5, 6)})  # A4 is empty
    spec["F5"] = "#N/A"  # an error cell in column F
    book = book_of({"S": spec})
    region = find_regions(book).regions[0]
    got, cell_mode, batch, rejected = run_region(book, region)
    assert set(rejected) == {4, 5}  # D reads A (blank row), E reads F (an error)
    assert "empty cell" in rejected[4] and "holds e" in rejected[5]
    assert all(same(got[k], cell_mode[k]) for k in batch_keys(batch))  # column C is region-computed and equal


def batch_keys(batch):
    return [s.key for s in batch]


def test_a_relative_cross_sheet_reference_is_not_a_region_but_an_absolute_one_is():  # Review Focus 4
    spec, other = {}, {"A1": 3.0}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*Other!$A$1", 0.0)
        spec[f"C{r}"] = (f"=A{r}*Other!A{r}", 0.0)
    book = book_of({"S": spec, "Other": other})
    scan = find_regions(book)
    assert [(r.cols) for r in scan.regions] == [[2]]
    got, cell_mode, batch, rejected = run_region(book, scan.regions[0])
    assert all(same(got[k], cell_mode[k]) for k in batch_keys(batch))


def test_header_names_fall_back_to_the_column_letter():  # Review Focus 5
    spec = {"A1": "x y", "B1": "x y", "C1": 5.0, "D1": "1st", "E1": "", "F1": "ok"}
    for r in range(2, 6):
        for c in "ABCDEF":
            spec[f"{c}{r}"] = float(r)
    book = book_of({"S": spec})
    from cellproof.regions import Region
    fake = Region("S", 2, 5, [2], {}, {})  # only the row above (row 1) and the sheet matter for naming
    assert names_for(book, fake, [1, 2, 3, 4, 5, 6]) == {
        1: "col_a", 2: "col_b",  # duplicate header text
        3: "col_c",  # a number, not text
        4: "col_d",  # starts with a digit
        5: "col_e",  # empty
        6: "ok"}


def test_a_column_that_cannot_compile_is_rejected_with_a_reason_and_its_dependants_follow():
    spec = {}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=MAX(A{r},1)", 0.0)  # a function that reads a reference itself: not row-wise in slice 1
        spec[f"C{r}"] = (f"=B{r}+1", 0.0)
        spec[f"D{r}"] = (f"=A{r}*2", 0.0)
    book = book_of({"S": spec})
    region = find_regions(book).regions[0]
    batch, rejected = compile_region(Compiler(book), region)
    assert set(rejected) == {2, 3}
    assert "stays in cell mode" in rejected[3]
    assert [s.key[2] for s in batch] == [4] * 4
    notes = region_note(region, batch, rejected)
    assert any("region_s_2_5" in n for n in notes) and any("column B" in n for n in notes)


def test_a_row_dependent_function_never_yields_a_region_value():
    spec = {"A1": "a"}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = ("=ROW()", 0.0)
        spec[f"C{r}"] = (f"=A{r}*2", 0.0)
    book = book_of({"S": spec})
    scan = find_regions(book)
    batch, rejected = compile_region(Compiler(book), scan.regions[0])
    assert rejected == {2: "ROW() in a block of rows not supported yet (stays row by row)"}
    assert all(s.key[2] != 2 for s in batch)


def _kw_book(head_a, head_b):
    spec = {"A1": head_a, "B1": head_b}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
    return book_of({"S": spec})


def _assert_region_equals_cell_mode(book):
    region = find_regions(book).regions[0]
    got, cell_mode, batch, rejected = run_region(book, region)
    assert rejected == {} and batch is not None
    for key in region.keys():
        assert same(got[key], cell_mode[key]), (key, got[key], cell_mode[key])


def test_a_header_that_looks_like_a_fallback_name_cannot_shadow_another_column():
    _assert_region_equals_cell_mode(_kw_book("col_b", ""))
    _assert_region_equals_cell_mode(_kw_book("col_b", "col_a"))


KEYWORDS = ["order", "group", "select", "from", "end", "case", "table", "desc", "limit", "all", "true", "false",
            "null", "user", "date", "year", "row", "index", "key", "value"]


@pytest.mark.parametrize("word", KEYWORDS)
def test_keyword_headers_on_formula_and_source_columns_run_like_cell_mode(word):
    _assert_region_equals_cell_mode(_kw_book("a", word.capitalize()))  # on the formula column
    _assert_region_equals_cell_mode(_kw_book(word, "b"))  # on the source column


def test_headers_differing_only_in_case_both_fall_back():
    spec = {"A1": "Tax", "B1": "tax"}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = float(r)
        spec[f"C{r}"] = (f"=A{r}+B{r}", 0.0)
    book = book_of({"S": spec})
    from cellproof.regions import Region
    fake = Region("S", 2, 5, [3], {}, {})
    assert names_for(book, fake, [1, 2]) == {1: "col_a", 2: "col_b"}


def test_a_boolean_input_column_pivots_correctly():
    spec = {"A1": "flag"}
    for r in range(2, 6):
        spec[f"A{r}"] = bool(r % 2)
        spec[f"B{r}"] = (f"=IF(A{r},10,20)", 0.0)
    _assert_region_equals_cell_mode(book_of({"S": spec}))


ALIAS_FORMULAS = ['=COUNTIF($D$2:$D$6,A{r})', '=A{r}&"x"', '=SUMIFS($D$2:$D$6,$D$2:$D$6,A{r})',
                  '=VLOOKUP(A{r},$D$2:$E$6,2,FALSE)', '=IF(A{r}>2,A{r},0)', '=ROUND(A{r},1)']


def _alias_book(header, formulas=ALIAS_FORMULAS):
    spec = {"A1": header}
    for i, r in enumerate(range(2, 7)):
        spec[f"A{r}"] = float([1, 2, 2, 3, 1][i])
        spec[f"D{r}"] = float([1, 2, 2, 3, 3][i])
        spec[f"E{r}"] = float(10 * (i + 1))
        for j, f in enumerate(formulas):
            spec[f"{'FGHIJK'[j]}{r}"] = (f.format(r=r), 0.0)
    return book_of({"S": spec})


def _check_alias(header, formulas=ALIAS_FORMULAS):
    book = _alias_book(header, formulas)
    region = find_regions(book).regions[0]
    got, cell_mode, batch, rejected = run_region(book, region)
    assert batch is not None
    for k in batch_keys(batch):
        assert same(got[k], cell_mode[k]), (header, k, got[k], cell_mode[k])
    return rejected


@pytest.mark.parametrize("header,formula", [("one", ALIAS_FORMULAS[0]), ("x", ALIAS_FORMULAS[0]),
                                            ("q1", ALIAS_FORMULAS[0]), ("v", ALIAS_FORMULAS[1])])
def test_a_header_equal_to_a_library_alias_cannot_shadow_a_column(header, formula):
    assert _check_alias(header, [formula]) == {}


HEADERS = ["a", "b", "c", "d", "e", "k", "o", "q", "q1", "q2", "r", "s", "t", "u", "v", "w", "x", "y", "z", "one",
           "a0", "a1", "l", "m", "n", "p", "i", "j"]


@pytest.mark.parametrize("header", HEADERS)
def test_no_header_changes_what_the_function_library_computes(header):
    rejected = _check_alias(header)
    assert 10 not in rejected and 11 not in rejected  # IF (J) and ROUND (K) are always accepted


def test_a_dissolved_region_is_reported_and_values_are_unchanged():
    # No pure column (both read a formula cell outside the region), so the region cannot be split: it dissolves.
    # (B pure + C reading a total of B is split instead: tests/test_regions_split.py.)
    spec = {"Y1": ("=1+1", 0.0)}
    for r in range(2, 5):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=$Z$1+A{r}", 0.0)
        spec[f"C{r}"] = (f"=$Y$1+A{r}", 0.0)
    spec["Z1"] = ("=C3+1", 0.0)  # B -> Z1 -> C3: a cycle for the block, none for cells
    book = book_of({"S": spec})
    on = compile_book(book)
    assert any("stay in cell mode" in n and "they could not be ordered as one block" in n for n in on.notes)
    assert not any(hasattr(b, "wide_sql") for b in on.batches)
    assert execute(book, on) == execute(book, compile_book(book, regions=False))


# ---- final-review wave: the SQL of a region is validated by DuckDB in both run forms --------------------------------

import shutil  # noqa: E402

from cellproof.cli import main  # noqa: E402
try:
    from cellproof.region import WIDE, _validate  # noqa: E402
except ImportError:
    _validate = None
from xlsxkit import write_xlsx  # noqa: E402


def _whole_run_equal(book):
    on = execute(book, compile_book(book))
    off = execute(book, compile_book(book, regions=False))
    assert set(on) == set(off)
    for k in off:
        assert same(on[k], off[k]), (k, on[k], off[k])


def _lookup_sheet(formulas):
    """Column A holds keys; E1:F4 is a lookup table, P1:S2 a horizontal one, T1:T4 text; formulas go to columns B...."""
    spec = {"A1": "key"}
    for i, r in enumerate(range(2, 7)):
        spec[f"A{r}"] = float([0, 1, 2, 3, 7][i])
    for i in range(4):
        spec[f"E{i + 1}"] = float(i)
        spec[f"F{i + 1}"] = float(10 * (i + 1))
        spec[f"T{i + 1}"] = "x" if i % 2 else "y"
        spec[f"{'PQRS'[i]}1"] = float(i)
        spec[f"{'PQRS'[i]}2"] = float(100 + i)
    for col, f in formulas.items():
        for r in range(2, 7):
            spec[f"{col}{r}"] = (f.format(r=r), 0.0)
    return spec


LOOKUPS = {
    "B": "=INDEX($F$1:$F$4,MATCH(A{r},$E$1:$E$4,0))",
    "C": "=A{r}*2",
    "D": "=INDEX($F$1:$F$4,A{r}+1)",
    "G": "=IFERROR(INDEX($F$1:$F$4,MATCH(A{r},$E$1:$E$4,0)),0)",
    "H": "=VLOOKUP(A{r},$E$1:$F$4,2,FALSE)",
    "I": "=HLOOKUP(A{r},$P$1:$S$2,2,FALSE)",
    "J": '=COUNTIFS($E$1:$E$4,">"&A{r},$T$1:$T$4,"x")',
    "K": "=XLOOKUP(A{r},$E$1:$E$4,$F$1:$F$4)",
    "L": "=MATCH(A{r},$E$1:$E$4,0)",
    "M": "=IFNA(VLOOKUP(A{r},$E$1:$F$4,2,FALSE),-1)",
}


@pytest.mark.parametrize("col", sorted(LOOKUPS))
def test_a_lookup_column_in_a_region_never_crashes_and_equals_cell_mode(col):
    _whole_run_equal(book_of({"S": _lookup_sheet({col: LOOKUPS[col], "C": LOOKUPS["C"]})}))


def test_a_column_duckdb_cannot_plan_is_rejected_with_its_reason_and_its_sibling_is_kept(monkeypatch):
    import cellproof.region as region_mod
    real = region_mod._validate  # no known formula fails to plan any more: make column J's SQL fail
    monkeypatch.setattr(region_mod, "_validate", lambda wide, rows: "Binder Error: boom" if " AS col_j" in wide
                        else real(wide, rows))
    book = book_of({"S": _lookup_sheet({"J": LOOKUPS["J"], "C": LOOKUPS["C"]})})
    batch, rejected = compile_region(Compiler(book), find_regions(book).regions[0])
    assert 10 in rejected and "DuckDB" in rejected[10]
    assert batch is not None and {s.key[2] for s in batch} == {3}
    _whole_run_equal(book)


def _with_saved_values(sheets):
    book = book_of(sheets)
    res = execute(book, compile_book(book, regions=False))
    out = {}
    for name, spec in sheets.items():
        out[name] = {}
        for ref, v in spec.items():
            if isinstance(v, tuple):
                from openpyxl.utils.cell import coordinate_to_tuple
                r, c = coordinate_to_tuple(ref)
                k, n, t = res[(name, r, c)]
                v = (v[0], {"n": n, "s": t, "e": t, "b": bool(n) if n is not None else None}[k])
            out[name][ref] = v
    return out


@pytest.mark.skipif(shutil.which("dbt") is None, reason="needs: pip install -e .[dbt]")
def test_dbt_build_passes_for_a_region_full_of_lookups(tmp_path):
    wb = tmp_path / "l.xlsx"
    write_xlsx(wb, _with_saved_values({"S": _lookup_sheet(LOOKUPS)}))
    assert main(["convert", str(wb), "--out", str(tmp_path / "out"), "--dbt"]) == 0


def test_validate_reports_the_first_line_of_a_duckdb_error_and_accepts_good_sql():
    row = "SELECT 'S', row_num, 2, 'n', CAST(row_num AS DOUBLE), NULL FROM " + WIDE
    good = "WITH\nsrc AS (SELECT row_num FROM range(2, 4) AS t(row_num))\nSELECT row_num\nFROM src"
    assert _validate(good, [row]) is None
    bad = good.replace("SELECT row_num\nFROM src", "SELECT nope\nFROM src")
    msg = _validate(bad, [row])
    assert msg and "\n" not in msg


@pytest.mark.parametrize("header", ["Cells", "All Cells", "all_cells", "Wide"])
def test_headers_that_name_internal_relations_still_run_and_build(header, tmp_path):
    spec = {"A1": "x", "B1": header, "C1": "other"}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=A{r}*3", 0.0)
    book = book_of({"S": spec})
    _whole_run_equal(book)
    if shutil.which("dbt") is not None:
        wb = tmp_path / "h.xlsx"
        write_xlsx(wb, _with_saved_values({"S": spec}))
        assert main(["convert", str(wb), "--out", str(tmp_path / "out"), "--dbt"]) == 0


def test_the_error_typed_column_reference_is_qualified():
    spec = {"A1": "x"}
    for r in range(2, 5):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = ("=#N/A", "#N/A")
        spec[f"C{r}"] = (f"=B{r}+A{r}", "#N/A")
    book = book_of({"S": spec})
    batch, _ = compile_region(Compiler(book), find_regions(book).regions[0])
    assert "CAST(_.col_b__err AS VARCHAR)" in batch.wide_sql
    _whole_run_equal(book)


def test_a_literal_that_looks_like_the_dbt_cte_does_not_reject_the_column():
    spec = {"A1": "a", "B1": "label"}
    for r in range(2, 6):
        spec[f"A{r}"] = "x"
        spec[f"B{r}"] = (f'=A{r}&" all_cells AS "', "x")
    book = book_of({"S": spec})
    region = find_regions(book).regions[0]
    got, cell_mode, batch, rejected = run_region(book, region)
    assert rejected == {} and batch is not None
    assert all(same(got[k], cell_mode[k]) for k in region.keys())


ROW_WISE = sorted(set(LOOKUPS) - {"C"})


@pytest.mark.parametrize("col", ROW_WISE)
def test_a_lookup_column_compiles_row_wise(col):
    book = book_of({"S": _lookup_sheet({col: LOOKUPS[col], "C": LOOKUPS["C"]})})
    batch, rejected = compile_region(Compiler(book), find_regions(book).regions[0])
    assert rejected == {}
    assert batch is not None and {s.key[2] for s in batch} == {3, ord(col) - 64}
    _whole_run_equal(book)


TEXT_LOOKUPS = {
    "B": "=VLOOKUP(A{r},$E$1:$F$6,2,FALSE)",  # first of duplicates, any case, wildcards, empty and error targets
    "C": "=XLOOKUP(A{r},$E$1:$E$6,$F$1:$F$6,\"none\",0,-1)",  # last match
    "D": "=MATCH(A{r},$E$1:$E$6,0)",
    "G": "=XLOOKUP(A{r},$E$1:$E$6,$F$1:$F$6,\"none\",2)",  # wildcard mode
    "H": "=INDEX($F$1:$F$6,MATCH(A{r},$E$1:$E$6,0))",
    "I": "=VLOOKUP(N{r},$K$1:$L$4,2)",  # approximate on sorted numbers
    "J": "=MATCH(N{r},$K$1:$K$4,1)",
    "M": "=XLOOKUP(N{r},$K$1:$K$4,$L$1:$L$4,-1,1)",  # next larger
    "O": "=LOOKUP(N{r},$K$1:$K$4,$L$1:$L$4)",
    "P": "=LOOKUP(N{r},$K$1:$L$4)",  # array form: last column
    "Q": '=MATCH(A{r},{{"apple","pear","kiwi"}},0)',  # array constants ({{ }} for str.format)
    "R": '=LOOKUP(N{r},{{0,10,20}},{{"low","mid","high"}})',
    "S": "=N{r}*SUM({{1,2;3,4}})",
}


def _text_lookup_sheet(formulas):
    spec = {"A1": "key", "N1": "amount"}
    keys = ["apple", "APPLE", "b*", "pear", "zzz", "x?z", "nope", "kiwi"]
    amounts = [5.0, 10.0, 15.0, 0.0, 99.0, 30.0, 7.5, 25.0]
    for i, r in enumerate(range(2, 10)):
        spec[f"A{r}"] = keys[i]
        spec[f"N{r}"] = amounts[i]
    table = [("Apple", 1.0), ("banana", "yellow"), ("apple", 3.0), ("pear", None), ("xyz", ("=1/0", "#DIV/0!")),
             ("kiwi", True)]
    for i, (k, v) in enumerate(table, 1):
        spec[f"E{i}"] = k
        if v is not None:
            spec[f"F{i}"] = v
    for i, (k, v) in enumerate([(0.0, "low"), (10.0, "mid"), (20.0, "high"), (30.0, "top")], 1):
        spec[f"K{i}"], spec[f"L{i}"] = k, v
    for col, f in formulas.items():
        for r in range(2, 10):
            spec[f"{col}{r}"] = (f.format(r=r), 0.0)
    return spec


@pytest.mark.parametrize("col", sorted(TEXT_LOOKUPS))
def test_text_and_approximate_lookups_compile_row_wise_and_equal_cell_mode(col):
    book = book_of({"S": _text_lookup_sheet({col: TEXT_LOOKUPS[col]})})
    notes = compile_book(book).notes
    assert any("-> model" in n for n in notes) and not any("cell mode" in n for n in notes), notes
    _whole_run_equal(book)


CRITERIA = {
    "B": "=COUNTIF($E$1:$E$8,A{r})",
    "C": '=COUNTIF($E$1:$E$8,">"&A{r})',
    "D": '=COUNTIF($E$1:$E$8,"<>"&A{r})',
    "G": "=SUMIF($E$1:$E$8,A{r},$F$1:$F$8)",
    "H": '=SUMIFS($F$1:$F$8,$E$1:$E$8,"<="&A{r},$T$1:$T$8,"x*")',
    "I": '=COUNTIFS($E$1:$E$8,">="&A{r},$T$1:$T$8,"<>y")',
    "J": "=AVERAGEIF($E$1:$E$8,A{r},$F$1:$F$8)",
    "K": '=MAXIFS($F$1:$F$8,$E$1:$E$8,"<"&A{r})',
    "L": '=MINIFS($F$1:$F$8,$T$1:$T$8,A{r}&"*")',
    "M": "=AVERAGEIFS($F$1:$F$8,$E$1:$E$8,A{r},$T$1:$T$8,\"<>\")",
}


CRITERIA_VALUES = {"num": [2.0, 0.0, 5.0, 1.0, -1.0, 2.0, 99.0, 0.5],
                   "txt": ["2", "a", "x", "b?", "A", "x*", "TRUE", "~*"]}


def _criteria_sheet(formulas, kind):
    """Criteria from the row (numbers, or text with wildcards), over numbers, text numbers, TRUE, blanks, an error."""
    spec = {"A1": "crit"}
    for i, v in enumerate(CRITERIA_VALUES[kind], 2):
        spec[f"A{i}"] = v
    column_e = [2.0, "2", "a", None, True, 5.0, "b1", 0.0]
    column_f = [10.0, 20.0, ("=1/0", "#DIV/0!"), 40.0, "t", 60.0, None, 80.0]
    column_t = ["x1", "y", None, "xx", "x", "y", "z", ""]
    for i in range(8):
        for col, values in (("E", column_e), ("F", column_f), ("T", column_t)):
            if values[i] is not None:
                spec[f"{col}{i + 1}"] = values[i]
    for col, f in formulas.items():
        for r in range(2, 10):
            spec[f"{col}{r}"] = (f.format(r=r), 0.0)
    return spec


@pytest.mark.parametrize("kind", sorted(CRITERIA_VALUES))
@pytest.mark.parametrize("col", sorted(CRITERIA))
def test_criteria_from_the_row_compile_row_wise_and_equal_cell_mode(col, kind):
    book = book_of({"S": _criteria_sheet({col: CRITERIA[col]}, kind)})
    notes = compile_book(book).notes
    assert any("-> model" in n for n in notes) and not any("cell mode" in n for n in notes), notes
    _whole_run_equal(book)
