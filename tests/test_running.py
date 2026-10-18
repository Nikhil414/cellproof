"""Running totals become window functions inside a region; values and verdicts must equal cell mode."""
import shutil

import pytest
from xlsxkit import book_of, write_xlsx

from unexcel.cli import main
from unexcel.compiler import compile_book
from unexcel.verifier import execute

N = 8  # rows 2..N+1
PLAIN = lambda r: f"=A{r}*1.5"  # noqa: E731
# text, #N/A, a number stored as text, TRUE and #DIV/0! partway down
MIXED = lambda r: f'=IF(A{r}=4,"abc",IF(A{r}=6,NA(),IF(A{r}=7,"5",IF(A{r}=8,TRUE,IF(A{r}=9,1/0,A{r})))))'  # noqa: E731


def same(a, b):
    return (a[0] == b[0] and a[2] == b[2]
            and (a[1] == b[1] or (a[1] is not None and b[1] is not None and abs(a[1] - b[1]) <= 1e-9 * max(1, abs(a[1])))))


def region_cols(compiled):
    return sorted(sorted({s.key[2] for s in b}) for b in compiled.batches if hasattr(b, "wide_sql"))


def in_region(compiled, col):
    return any(col in cols for cols in region_cols(compiled))


def check_equal(book):
    compiled = compile_book(book)
    without = compile_book(book, regions=False)
    assert compiled.unresolved == without.unresolved
    a, b = execute(book, compiled), execute(book, without)
    assert a.keys() == b.keys()
    for key in a:
        assert same(a[key], b[key]), (key, a[key], b[key])
    return compiled


def sheet(c_formula, b_formula=PLAIN, seed=None, extra=None):
    """A = row number, B = b_formula, C = c_formula from row 2; C1 = seed (None = empty)."""
    spec = {}
    if seed is not None:
        spec["C1"] = seed
    for r in range(2, N + 2):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (b_formula(r), 0.0)
        spec[f"C{r}"] = (c_formula(r), 0.0)
    spec.update(extra or {})
    return book_of({"S": spec})


def test_recurrence_becomes_a_region_column():
    compiled = check_equal(sheet(lambda r: f"=C{r - 1}+B{r}", seed=0.0))
    assert in_region(compiled, 3)


def test_recurrence_with_formula_seed_like_the_benchmark():
    spec = {"A1": 1.0, "B1": ("=A1*1.1", 0.0), "C1": ("=B1", 0.0)}
    for r in range(2, N + 2):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*1.1", 0.0)
        spec[f"C{r}"] = (f"=C{r - 1}+B{r}", 0.0)
        spec[f"D{r}"] = (f"=IF(B{r}>3,B{r},0)", 0.0)
    compiled = check_equal(book_of({"S": spec}))
    assert in_region(compiled, 3)


@pytest.mark.parametrize("seed", [None, "total", "7", True, 2.5, "#N/A", ("=1/0", "#DIV/0!")])
def test_recurrence_seed_kinds(seed):
    compiled = check_equal(sheet(lambda r: f"=C{r - 1}+B{r}", seed=seed))
    assert in_region(compiled, 3)


@pytest.mark.parametrize("formula", [lambda r: f"=C{r - 1}+B{r}", lambda r: f"=C{r - 1}-B{r}",
                                     lambda r: f"=B{r}+C{r - 1}", lambda r: f"=C{r - 1}+B{r}*2-A{r}"])
def test_recurrence_with_errors_text_and_booleans_partway_down(formula):
    compiled = check_equal(sheet(formula, MIXED, seed=1.0))
    assert in_region(compiled, 3)


def test_latest_error_wins_when_the_previous_row_is_on_the_right():
    book = sheet(lambda r: f"=B{r}+C{r - 1}", lambda r: f"=IF(A{r}=3,NA(),IF(A{r}=5,1/0,A{r}))", seed=0.0)
    compiled = check_equal(book)
    assert in_region(compiled, 3)
    assert execute(book, compiled)[("S", 9, 3)][2] == "#DIV/0!"  # B9 + C8: C8 carries B5's #DIV/0!, not B3's #N/A


@pytest.mark.parametrize("formula", [lambda r: f"=C{r - 1}*1.05+B{r}", lambda r: f"=B{r}-C{r - 1}",
                                     lambda r: f"=C{r - 1}+C{r - 1}", lambda r: f"=MAX(C{r - 1},B{r})"])
def test_other_recurrences_stay_in_cell_mode(formula):
    compiled = check_equal(sheet(formula, seed=1.0))
    assert not in_region(compiled, 3)


@pytest.mark.parametrize("start", [2, 1])
def test_growing_sum_becomes_a_region_column(start):
    extra = {"B1": 10.0} if start == 1 else {}
    compiled = check_equal(sheet(lambda r: f"=SUM($B${start}:B{r})", extra=extra))
    assert in_region(compiled, 3)


def test_growing_sum_skips_text_and_booleans_and_keeps_the_first_error():
    compiled = check_equal(sheet(lambda r: f"=SUM(B$2:B{r})", MIXED))
    assert in_region(compiled, 3)


@pytest.mark.parametrize("b1", ["header", ("=1/0", "#DIV/0!"), ("=2*3", 6.0), True])
def test_growing_sum_with_a_row_above_the_region(b1):
    check_equal(sheet(lambda r: f"=SUM($B$1:B{r})", extra={"B1": b1}))


def test_growing_sum_inside_a_larger_formula():
    compiled = check_equal(sheet(lambda r: f"=SUM($B$2:B{r})/A{r}+1"))
    assert in_region(compiled, 3)


def test_growing_sum_over_an_input_column():
    spec = {}
    for r in range(2, N + 2):
        spec[f"B{r}"] = float(r * r)
        spec[f"C{r}"] = (f"=SUM($B$2:B{r})", 0.0)
    compiled = check_equal(book_of({"S": spec}))
    assert in_region(compiled, 3)


@pytest.mark.parametrize("formula", [lambda r: f"=SUM($C$2:C{r})", lambda r: f"=SUM($B$2:B{r},1)",
                                     lambda r: f"=SUM($A$2:B{r})", lambda r: f"=COUNT($B$2:B{r})",
                                     lambda r: f"=SUM(B2:B{r})"])
def test_other_growing_ranges_stay_in_cell_mode(formula):
    compiled = check_equal(sheet(formula))
    assert not in_region(compiled, 3)


def with_saved_values(spec):
    """The same workbook with every formula's saved value set to cell mode's answer (Excel's left-to-right order)."""
    book = book_of({"S": spec})
    got = execute(book, compile_book(book, regions=False))
    out = dict(spec)
    for (_, r, c), (kind, n, t) in got.items():
        ref = f"{chr(64 + c)}{r}"
        out[ref] = (spec[ref][0], {"n": n, "b": bool(n), "s": t, "e": t}[kind])
    return book_of({"S": out})


def ledger(rows):
    """Amounts with cents; every 50th row books back the 49 before it, so the balance keeps returning to ~0."""
    spec, run = {"C1": 0.0}, 0.0
    for r in range(2, rows + 2):
        if (r - 1) % 50 == 0:
            spec[f"B{r}"] = -run
            run = 0.0
        else:
            v = ((r * 7919) % 1000001) / 100.0 - 5000.0
            spec[f"B{r}"] = v
            run += v
        spec[f"C{r}"] = (f"=C{r - 1}+B{r}", 0.0)
        spec[f"D{r}"] = (f"=SUM($B$2:B{r})", 0.0)
    return with_saved_values(spec)


def test_a_balance_that_cancels_to_zero_stays_in_excels_order():
    from unexcel.verifier import verify
    book = ledger(1000)
    compiled = compile_book(book)
    assert not in_region(compiled, 3) and not in_region(compiled, 4)
    assert any("cancel out" in n for n in compiled.notes)
    assert {v.status for v in verify(book, compiled)} == {"PASS"}


def test_a_one_sign_running_total_with_saved_values_stays_a_window():
    from unexcel.verifier import verify
    spec = {"C1": 0.0}
    for r in range(2, 2002):
        spec[f"B{r}"] = ((r * 7919) % 100001) / 100.0
        spec[f"C{r}"] = (f"=C{r - 1}+B{r}", 0.0)
        spec[f"D{r}"] = (f"=SUM($B$2:B{r})", 0.0)
    book = with_saved_values(spec)
    compiled = compile_book(book)
    assert in_region(compiled, 3) and in_region(compiled, 4)
    assert {v.status for v in verify(book, compiled)} == {"PASS"}


@pytest.mark.parametrize("formula", [lambda r: f"=C{r - 1}+SUM($B$2:B{r})", lambda r: f"=SUM($B$2:B{r})+C{r - 1}"])
def test_a_running_total_of_a_growing_sum(formula):
    compiled = check_equal(sheet(formula, seed=1.0))
    assert in_region(compiled, 3)


@pytest.mark.parametrize("name", ["__ROWREF_2__", "__RUN_2_2__"])
def test_a_name_spelled_like_a_placeholder_is_not_read_as_a_column(name):
    spec = {"Z1": 100.0}
    for r in range(2, N + 2):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*1.5", 0.0)
        spec[f"C{r}"] = (f"={name}+A{r}", 0.0)
    book = book_of({"S": spec}, names={name: ("S", "Z1")})
    compiled = check_equal(book)
    assert not in_region(compiled, 3)
    assert execute(book, compiled)[("S", 2, 3)][1] == 102.0


def test_two_running_columns_and_a_column_reading_them():
    spec = {"C1": 0.0}
    for r in range(2, N + 2):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*1.5", 0.0)
        spec[f"C{r}"] = (f"=C{r - 1}+B{r}", 0.0)
        spec[f"D{r}"] = (f"=SUM($C$2:C{r})", 0.0)
        spec[f"E{r}"] = (f"=D{r}-C{r}", 0.0)
    compiled = check_equal(book_of({"S": spec}))
    assert in_region(compiled, 3) and in_region(compiled, 4) and in_region(compiled, 5)


def test_running_columns_build_in_dbt(tmp_path):
    if shutil.which("dbt") is None:
        pytest.skip("needs: pip install -e .[dbt]")
    spec = {"C1": 0.0}
    for r in range(2, N + 2):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (MIXED(r), 0.0)
        spec[f"C{r}"] = (f"=C{r - 1}+B{r}", 0.0)
        spec[f"D{r}"] = (f"=SUM($B$2:B{r})", 0.0)
    book = book_of({"S": spec})
    saved = execute(book, compile_book(book, regions=False))
    out = {}
    for key, cell in book.cells.items():
        ref = f"{chr(64 + key[2])}{key[1]}"
        if cell.formula:
            kind, n, t = saved[key]
            out[ref] = ("=" + cell.formula, {"n": n, "b": bool(n), "s": t, "e": t}[kind])
        else:
            out[ref] = cell.value
    wb = tmp_path / "running.xlsx"
    write_xlsx(wb, {"S": out})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out"), "--dbt"]) == 0
