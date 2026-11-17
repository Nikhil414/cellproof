"""`cellproof sql`: the printed chain must run, and must answer exactly what the verifier answers."""
import duckdb
import pytest
from xlsxkit import book_of

from cellproof import query
from cellproof.compiler import compile_book
from cellproof.verifier import execute


def _region_book(n=12):
    """A region (B = A*2 down the rows), a cross-sheet VLOOKUP, an array constant, and a chain on top."""
    sheet = {"A1": "amount", "B1": "tax"}
    for r in range(2, n + 2):
        sheet[f"A{r}"] = float(r)
        sheet[f"B{r}"] = (f"=A{r}*2", float(r * 2))
    sheet.update({"D1": "Rate", "D2": 0.05,
                  "F1": ("=VLOOKUP(10, A2:B13, 2, FALSE)", 20.0),
                  "G1": ("=SUM({1,2,3})*D2", 0.3),
                  "H1": ("=F1+G1", 20.3),
                  "J1": ("=Z1+1", 1.0)})  # reads nothing: a batch the chain must drop
    return book_of({"S": sheet})


def _run(sql: str) -> list[tuple]:
    return duckdb.connect().execute(sql).fetchall()


def _answer(book, sheet, rng):
    compiled = compile_book(book)
    return compiled, query.build(book, compiled, query.selection(book, sheet, rng))


def test_chain_answers_exactly_what_the_verifier_answers():
    book = _region_book()
    compiled, answer = _answer(book, "S", "F1:H1")
    got = execute(book, compiled)
    rows = _run(answer.sql)
    assert [(r[1], r[2]) for r in rows] == [(1, 6), (1, 7), (1, 8)]  # F1, G1, H1 in column order
    for sheet, row, col, kind, num, txt in rows:
        assert got[(sheet, row, col)] == (kind, num, txt)


def test_a_batch_nothing_reads_is_dropped(monkeypatch):
    import cellproof.compiler as compiler
    monkeypatch.setattr(compiler, "BATCH", 1)  # one cell per batch, so J1 can be dropped alone
    book = _region_book()
    compiled, answer = _answer(book, "S", "F1:H1")
    assert ("S", 1, 10) in {s.key for s in compiled.steps}  # J1 was compiled...
    assert ("S", 1, 10) not in answer.computed  # ...but nothing in the chain reads it
    assert "J1" not in answer.sql


def test_a_batch_is_kept_whole_even_when_one_of_its_cells_is_not_needed():
    """A batch is one statement: J1 shares F1/G1/H1's batch here, so it is computed whether or not it is wanted."""
    book = _region_book()
    _, answer = _answer(book, "S", "F1:H1")
    assert ("S", 1, 10) in answer.computed
    assert [k for k in answer.selection] == [("S", 1, 6), ("S", 1, 7), ("S", 1, 8)]  # but not returned


def test_a_cell_reading_a_region_carries_the_whole_region_batch():
    book = _region_book()
    compiled, answer = _answer(book, "S", "F1")
    # the VLOOKUP reads A2:B13, so every cell of the region's one statement comes along
    assert {("S", r, 2) for r in range(2, 14)} <= set(answer.computed)
    assert ("S", 2, 2) in answer.computed and ("S", 13, 2) in answer.computed


def test_inputs_the_chain_reads_are_in_the_seed_and_no_others():
    book = _region_book()
    _, answer = _answer(book, "S", "G1")  # SUM({1,2,3})*D2: the array constant and D2, nothing else
    assert answer.inputs == 4  # D1 ("Rate"), D2 (0.05) and the three array constants... less D1
    rows = _run(answer.sql)
    assert [(r[1], r[2]) for r in rows] == [(1, 7)]
    assert rows[0][3] == "n" and abs(rows[0][4] - 0.3) < 1e-12


def test_an_array_constant_survives_the_pruning():
    book = book_of({"S": {"A1": 0.5, "B1": ("=SUM({1,2,3})*A1", 3.0)}})
    _, answer = _answer(book, None, None)
    rows = _run(answer.sql)
    assert [(r[0], r[1], r[2]) for r in rows] == [("S", 1, 1), ("S", 1, 2)]  # no #array- sheets
    assert rows[1][4] == 3.0


def test_the_whole_workbook_hides_the_array_constant_sheets():
    book = book_of({"S": {"A1": 3, "B1": ("=SUM({1,2,3})", 6)}})
    compiled, answer = _answer(book, None, None)
    assert book.arrays  # the constants are cells of a hidden sheet...
    assert {k[0] for k in answer.selection} == {"S"}  # ...that the answer does not list
    assert [(r[0], r[1], r[2]) for r in _run(answer.sql)] == [("S", 1, 1), ("S", 1, 2)]


def test_a_sheet_selection_covers_the_whole_sheet():
    book = book_of({"S": {"A1": 1, "C9": ("=A1*3", 3)}, "T": {"A1": 5}})
    _, answer = _answer(book, "S", None)
    assert [(r[1], r[2]) for r in _run(answer.sql)] == [(1, 1), (9, 3)]


def test_the_whole_workbook_lists_every_sheet():
    book = book_of({"S": {"A1": 1}, "T": {"B2": ("=1+1", 2)}})
    _, answer = _answer(book, None, None)
    assert [(r[0], r[1], r[2]) for r in _run(answer.sql)] == [("S", 1, 1), ("T", 2, 2)]


def test_unresolved_cells_are_reported_not_silently_missing():
    book = book_of({"S": {"A1": 1, "B1": ("=BESSELJ(1,1)", 0.0), "C1": ("=A1+1", 2)}})
    _, answer = _answer(book, "S", "A1:C1")
    assert set(answer.unresolved) == {("S", 1, 2)}
    assert "BESSELJ" in answer.unresolved[("S", 1, 2)]
    assert [(r[1], r[2]) for r in _run(answer.sql)] == [(1, 1), (1, 3)]  # B1 has no row to return


def test_a_reader_of_an_unresolved_cell_says_so_too():
    book = book_of({"S": {"A1": ("=BESSELJ(1,1)", 0.0), "B1": ("=A1+1", 1)}})
    _, answer = _answer(book, "S", "B1")
    assert "unresolved" in answer.unresolved[("S", 1, 2)]


def test_a_blank_cell_in_the_rectangle_is_not_in_the_answer():
    book = book_of({"S": {"A1": 1, "A3": 3}})
    _, answer = _answer(book, "S", "A1:A3")
    assert [k for k in answer.selection] == [("S", 1, 1), ("S", 3, 1)]
    assert [(r[1], r[2]) for r in _run(answer.sql)] == [(1, 1), (3, 1)]


def test_text_and_errors_round_trip():
    book = book_of({"S": {"A1": "it's; a \\ test", "A2": ("=1/0", "#DIV/0!"), "A3": True, "A4": 1e300}})
    _, answer = _answer(book, "S", "A1:A4")
    rows = _run(answer.sql)
    assert rows == [("S", 1, 1, "s", None, "it's; a \\ test"), ("S", 2, 1, "e", None, "#DIV/0!"),
                    ("S", 3, 1, "b", 1.0, None), ("S", 4, 1, "n", 1e300, None)]


# --- what the command refuses -------------------------------------------------------------------

def test_range_without_sheet_needs_a_one_sheet_workbook():
    book = book_of({"S": {"A1": 1}, "T": {"A1": 2}})
    with pytest.raises(query.QueryError, match="--range needs --sheet"):
        query.selection(book, None, "A1:A2")


def test_range_without_sheet_uses_the_only_sheet():
    book = book_of({"Only": {"A1": 1, "A2": 2}})
    assert query.selection(book, None, "A1:A2").label == "Only!A1:A2"


def test_an_unknown_sheet_is_refused():
    book = book_of({"S": {"A1": 1}})
    with pytest.raises(query.QueryError, match="no sheet 'Nope'"):
        query.selection(book, "Nope", None)


@pytest.mark.parametrize("text", ["D:D", "A:C"])
def test_a_whole_column_is_refused(text):
    with pytest.raises(query.QueryError, match="whole column"):
        query.box(text)


def test_whole_row_is_refused_as_not_an_a1_range():
    with pytest.raises(query.QueryError, match="not an A1 range"):
        query.box("1:1")


@pytest.mark.parametrize("text", ["hello", "A1:", "A1:B2:C3", ""])
def test_text_that_is_not_a_range_is_refused(text):
    with pytest.raises(query.QueryError, match="not an A1 range"):
        query.box(text)


def test_a_sheet_inside_the_range_text_is_refused():
    with pytest.raises(query.QueryError, match="give the sheet with --sheet"):
        query.box("S!A1:A9")


def test_box_ignores_case_and_dollars():
    assert query.box("$d$2:$d$9") == (2, 4, 9, 4)
    assert query.box(" a1 ") == (1, 1, 1, 1)


def test_a_range_past_the_last_column_is_refused():
    with pytest.raises(query.QueryError, match="columns"):
        query.box("A1:XFE9")


# --- the printed SQL ----------------------------------------------------------------------------

def test_only_sql_goes_to_stdout_and_it_is_ascii():
    book = _region_book()
    _, answer = _answer(book, "S", "F1:H1")
    assert answer.sql.isascii()
    assert answer.sql.startswith("-- cellproof sql:")
    assert answer.sql.rstrip().endswith("ORDER BY row_num, col_num;")


def test_the_chain_needs_nothing_but_the_cells_table():
    """No UDFs, no macros: what the compiler emits is all a fresh connection needs."""
    book = _region_book()
    _, answer = _answer(book, "S", "F1:H1")
    con = duckdb.connect()
    assert con.execute(answer.sql).fetchall()


def test_stats_are_in_the_header_comment():
    book = _region_book()
    _, answer = _answer(book, "S", "F1:H1")
    assert f"{len(answer.computed):,} formula cells computed" in answer.sql
    assert f"{answer.inputs:,} input cells" in answer.sql


def test_check_refuses_what_selection_refuses():
    """`check` runs before the compile (a typo'd range must not cost one) and must not be a weaker gate."""
    book = book_of({"S": {"A1": 1}, "T": {"A1": 2}})
    assert query.check(book, None, None) is None
    assert query.check(book, "S", "A1:A2") is None
    for sheet, rng in [(None, "A1:A2"), ("Nope", None), ("S", "A:A"), ("S", "S!A1:A2")]:
        with pytest.raises(query.QueryError):
            query.check(book, sheet, rng)


def test_unresolved_reasons_are_grouped_and_ranked():
    """A cell that reads an untranslatable one is counted with it, so the number says what one fix would free."""
    book = book_of({"S": {"A1": ("=BESSELJ(1,1)", 0.0), "A2": ("=A1+1", 1), "A3": ("=A2+1", 2),
                          "B1": ("=TODAY()", 1), "B2": ("=B1+1", 2)}})
    _, answer = _answer(book, "S", "A1:B3")
    assert [g["cells"] for g in answer.groups] == [3, 2]
    assert "BESSELJ" in answer.groups[0]["reason"]
    assert "TODAY" in answer.groups[1]["reason"]
    assert answer.groups[0]["examples"] == ["S!A1"]  # only the root cell is named, not the ones it blocks
