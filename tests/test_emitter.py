import re
import shutil
from pathlib import Path

import duckdb
import pytest
from xlsxkit import SALES, book_of, write_xlsx

import cellproof.compiler as compiler
from cellproof.cli import main
from cellproof.compiler import READ_SEED, compile_book, render_batch
from cellproof.emitter import CHECKPOINT, checkpoint_name, dbt_sql, model_name, project_name, run_dbt, sources_for, write_project
from cellproof.verifier import execute


def test_writes_dbt_project(tmp_path):
    book = book_of({"S": {"A1": 2, "A2": ("=A1*2", 4)}})
    write_project(book, compile_book(book), tmp_path, "My Book 2026")
    assert "name: my_book_2026" in (tmp_path / "dbt_project.yml").read_text(encoding="utf-8")
    assert "send_anonymous_usage_stats: false" in (tmp_path / "dbt_project.yml").read_text(encoding="utf-8")
    assert (tmp_path / "profiles.yml").exists() and (tmp_path / "models" / "schema.yml").exists()
    seed = (tmp_path / "seeds" / "input_cells.csv").read_text(encoding="utf-8").splitlines()
    assert seed == ["sheet,row_num,col_num,kind,num,txt", "S,1,1,n,2.0,"]
    batch = (tmp_path / "models" / "cellproof_cells" / "cells_0001.sql").read_text(encoding="utf-8")
    assert "{{ ref('input_cells') }}" in batch
    assert "{{ ref('cp_0001') }}" in (tmp_path / "models" / "workbook_cells.sql").read_text(encoding="utf-8")


def test_rerun_removes_stale_batches(tmp_path):
    two = book_of({"S": {"A1": 1, "A2": ("=A1", 1), "A3": ("=SUM(A2:A2)", 1)}})  # the range closes batch 1
    write_project(two, compile_book(two), tmp_path, "b")
    one = book_of({"S": {"A1": 1, "A2": ("=A1", 1)}})
    write_project(one, compile_book(one), tmp_path, "b")
    assert sorted(p.name for p in (tmp_path / "models" / "cellproof_cells").iterdir()) == ["cells_0001.sql", "cp_0001.sql"]


def test_project_name():
    assert project_name("My Book 2026") == "my_book_2026"
    assert project_name("2026 plan!") == "wb_2026_plan"
    assert project_name("***") == "workbook"


def test_dbt_sql_survives_jinja():
    jinja2 = pytest.importorskip("jinja2")
    book = book_of({"S": {"A1": "a", "A2": ('=A1&"{{ oops }} {% x %} {#"', "a{{ oops }} {% x %} {#")}})
    steps = compile_book(book).batches[0]
    rendered = jinja2.Template(dbt_sql(steps, ["input_cells"])).render(ref=lambda name: name)
    prefix = "WITH\nall_cells AS (\nSELECT * FROM input_cells\n),\n"
    assert rendered == render_batch(steps, "all_cells", own_only=True).replace("WITH\n", prefix, 1)


def test_sources_for_reads_the_latest_checkpoint_plus_later_batches():  # Review Focus 3
    assert sources_for(1) == ["input_cells"]
    assert sources_for(64) == ["input_cells"] + [model_name(i) for i in range(1, 64)]
    assert sources_for(65) == [checkpoint_name(64)]
    assert sources_for(66) == [checkpoint_name(64), model_name(65)]
    assert sources_for(129) == [checkpoint_name(128)]
    assert CHECKPOINT == 64


def _run_models(out: Path):
    """Run the generated models in DuckDB without dbt: ref('x') becomes the table x, built in batch order."""
    con = duckdb.connect()
    con.execute(f"CREATE TABLE input_cells AS SELECT * FROM {READ_SEED}", [str(out / "seeds" / "input_cells.csv")])
    models = out / "models"
    n = len(list((models / "cellproof_cells").glob("cells_*.sql")))
    names = []
    for k in range(1, n + 1):
        names.append(model_name(k))
        if (models / "cellproof_cells" / f"{checkpoint_name(k)}.sql").exists():
            names.append(checkpoint_name(k))
    unref = lambda sql: re.sub(r"\{\{ ref\('(\w+)'\) \}\}", r"\1", sql)
    made = set()
    for name in names:
        sql = (models / "cellproof_cells" / f"{name}.sql").read_text(encoding="utf-8")
        for region in re.findall(r"ref\('(region_\w+)'\)", sql):  # a region's wide model comes before its long twin
            wide = (models / "cellproof_cells" / f"{region}.sql").read_text(encoding="utf-8")
            if region not in made:
                made.add(region)
                con.execute(f"CREATE TABLE {region} AS " + unref(wide))
        con.execute(f"CREATE TABLE {name} AS " + unref(sql))
    final = (models / "workbook_cells.sql").read_text(encoding="utf-8")
    return con.execute(re.sub(r"\{\{ ref\('(\w+)'\) \}\}", r"\1", final)).fetchall()


def test_generated_models_equal_the_verifier_across_checkpoints(tmp_path, monkeypatch):  # Review Focus 3
    monkeypatch.setattr(compiler, "BATCH", 1)  # one cell per batch: 130 batches, so cp_0064, cp_0128 and a final cp_0130
    spec = {"A1": 1.0}
    spec.update({f"B{r}": (f"=A1+{r}", 1.0 + r) for r in range(1, 131)})
    spec["C1"] = ("=SUM(B1:B130)", sum(1.0 + r for r in range(1, 131)))
    book = book_of({"S": spec})
    compiled = compile_book(book)
    assert len(compiled.batches) > 129
    write_project(book, compiled, tmp_path, "many")
    names = {p.name for p in (tmp_path / "models" / "cellproof_cells").iterdir()}
    assert {"cp_0064.sql", "cp_0128.sql", f"cp_{len(compiled.batches):04d}.sql"} <= names
    got = {(s, r, c): (k, n, t) for s, r, c, k, n, t in _run_models(tmp_path)}
    assert got == execute(book, compiled)


def test_project_without_formulas_runs(tmp_path):  # Review Focus 1
    book = book_of({"S": {"A1": 1.0}})
    write_project(book, compile_book(book), tmp_path, "plain")
    assert _run_models(tmp_path) == []


def test_each_batch_model_selects_only_its_own_rows(tmp_path):
    book = book_of({"S": {"A1": 2, "A2": ("=A1*2", 4)}})
    write_project(book, compile_book(book), tmp_path, "own")
    sql = (tmp_path / "models" / "cellproof_cells" / "cells_0001.sql").read_text(encoding="utf-8")
    assert "all_cells AS" in sql and "{{ ref('input_cells') }}" in sql and "SELECT * FROM all_cells" not in sql


def test_run_dbt_without_dbt_installed(tmp_path, monkeypatch):
    monkeypatch.setattr("cellproof.emitter.shutil.which", lambda name: None)
    assert run_dbt(tmp_path) == (False, "dbt is not installed - run: pip install cellproof[dbt]")


@pytest.mark.skipif(shutil.which("dbt") is None, reason="needs: pip install -e .[dbt]")
def test_dbt_build_passes_on_generated_project(tmp_path):
    wb = tmp_path / "commission.xlsx"
    write_xlsx(wb, {"Sales": SALES, "C": {"A1": "Asha", "B1": ("=SUMIFS(Sales!C:C, Sales!A:A, A1)", 120000)}})
    out = tmp_path / "out"
    assert main(["convert", str(wb), "--out", str(out), "--dbt"]) == 0
    con = duckdb.connect(str(out / "cellproof.duckdb"), read_only=True)
    assert con.execute("SELECT sheet, row_num, col_num, kind, num FROM workbook_cells").fetchall() == [
        ("C", 1, 2, "n", 120000.0)]


def region_book(n=8):
    spec = {"A1": "amount", "B1": "tax"}
    for r in range(2, n + 2):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*1.1", 0.0)
        spec[f"C{r}"] = (f"=SUM($A$2:A{r})", 0.0)  # growing range: a cell-mode batch after the region
    return book_of({"S": spec})


def test_a_region_writes_a_wide_model_and_a_long_model(tmp_path):
    book = region_book()
    compiled = compile_book(book)
    assert any(hasattr(b, "wide_sql") for b in compiled.batches)
    write_project(book, compiled, tmp_path, "r")
    names = {p.name for p in (tmp_path / "models" / "cellproof_cells").iterdir()}
    assert "region_s_2_9.sql" in names and "cells_0001.sql" in names
    wide = (tmp_path / "models" / "cellproof_cells" / "region_s_2_9.sql").read_text(encoding="utf-8")
    assert wide.startswith("WITH\nall_cells AS (") and "AS tax" in wide and "{{ ref('input_cells') }}" in wide
    long = (tmp_path / "models" / "cellproof_cells" / "cells_0001.sql").read_text(encoding="utf-8")
    assert "{{ ref('region_s_2_9') }}" in long
    schema = (tmp_path / "models" / "schema.yml").read_text(encoding="utf-8")
    assert "region_s_2_9" in schema


def test_region_models_equal_the_verifier(tmp_path):
    book = region_book()
    compiled = compile_book(book)
    write_project(book, compiled, tmp_path, "r")
    got = {(s, r, c): (k, n, t) for s, r, c, k, n, t in _run_models(tmp_path)}
    expected = execute(book, compiled)
    assert got.keys() == expected.keys()
    for key, value in expected.items():
        assert got[key][0] == value[0] and got[key][2] == value[2]
        assert (got[key][1] is None and value[1] is None) or abs(got[key][1] - value[1]) < 1e-9


def test_region_text_with_braces_survives_jinja(tmp_path):
    jinja2 = pytest.importorskip("jinja2")
    spec = {"A1": "a"}
    for r in range(2, 6):
        spec[f"A{r}"] = "x"
        spec[f"B{r}"] = (f'=A{r}&"{{{{ oops }}}} {{% x %}}"', "x")
    book = book_of({"S": spec})
    compiled = compile_book(book)
    batch = next(b for b in compiled.batches if hasattr(b, "wide_sql"))
    from cellproof.emitter import dbt_region_wide
    rendered = jinja2.Template(dbt_region_wide(batch, ["input_cells"])).render(ref=lambda name: name)
    assert "{{ oops }} {% x %}" in rendered


def _sheet(n=8):
    spec = {"A1": "amount", "B1": "tax"}
    for r in range(2, n + 2):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*1.1", r * 1.1)
    return spec


def _assert_models_equal_verifier(book, tmp_path):
    compiled = compile_book(book)
    write_project(book, compiled, tmp_path, "r")
    got = {(s, r, c): (k, n, t) for s, r, c, k, n, t in _run_models(tmp_path)}
    expected = execute(book, compiled)
    assert got.keys() == expected.keys()
    for key, value in expected.items():
        assert got[key][0] == value[0] and got[key][2] == value[2]
        assert (got[key][1] is None and value[1] is None) or abs(got[key][1] - value[1]) < 1e-9
    return compiled


def test_region_models_of_colliding_sheet_names_are_unique(tmp_path):
    book = book_of({"My Sheet": _sheet(), "My_Sheet": _sheet()})
    _assert_models_equal_verifier(book, tmp_path)
    cells = tmp_path / "models" / "cellproof_cells"
    wides = sorted(p.name for p in cells.glob("region_*.sql"))
    assert len(wides) == 2 and wides[0] != wides[1]
    schema = (tmp_path / "models" / "schema.yml").read_text(encoding="utf-8")
    names = re.findall(r"- name: (region_\w+)", schema)
    assert len(names) == 2 and len(set(names)) == 2
    assert "cells_0001" in schema and "cells_0002" in schema


def test_three_way_collision_gets_numbered_suffixes(tmp_path):
    book = book_of({"A-B": _sheet(), "A_B": _sheet(), "A B": _sheet()})
    _assert_models_equal_verifier(book, tmp_path)
    names = {p.stem for p in (tmp_path / "models" / "cellproof_cells").glob("region_*.sql")}
    base = min(names, key=len)
    assert names == {base, base + "_2", base + "_3"}


def test_sheet_named_like_the_unpivot_marker(tmp_path):
    book = book_of({"__cellproof_wide__": _sheet()})
    _assert_models_equal_verifier(book, tmp_path)


@pytest.mark.skipif(shutil.which("dbt") is None, reason="needs: pip install -e .[dbt]")
def test_dbt_build_with_colliding_region_names(tmp_path):
    wb = tmp_path / "c.xlsx"
    write_xlsx(wb, {"My Sheet": _sheet(), "My_Sheet": _sheet()})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out"), "--dbt"]) == 0


def test_array_constants_reach_the_dbt_models(tmp_path):
    spec = {"A1": ('=MATCH("b",{"a","b"},0)', 2.0)}
    spec.update({f"B{r}": (f"=LOOKUP({r},{{0,3,6}},{{10,20,30}})", 0.0) for r in range(1, 9)})
    book = book_of({"S": spec})
    compiled = compile_book(book)
    assert compiled.unresolved == {}
    write_project(book, compiled, tmp_path, "arrays")
    got = {(s, r, c): (k, n, t) for s, r, c, k, n, t in _run_models(tmp_path)}
    assert got == execute(book, compiled)
