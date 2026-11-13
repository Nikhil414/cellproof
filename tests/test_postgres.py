"""The Postgres target (cellproof.postgres): every cell verified again on a real Postgres."""
import json
import os

import pytest

pytest.importorskip("sqlglot")
pytest.importorskip("psycopg")
pgserver = pytest.importorskip("pgserver")

from xlsxkit import book_of, write_xlsx  # noqa: E402

from cellproof import postgres  # noqa: E402
from cellproof.cli import main  # noqa: E402
from cellproof.compiler import compile_book  # noqa: E402


@pytest.fixture(scope="module")
def con():
    # CI's Postgres service, else a Postgres of cellproof's own (pgserver)
    with postgres.connect(os.environ.get("CELLPROOF_TEST_PG_URL")) as c:
        yield c


def test_a_workbook_verifies_on_postgres(con):
    book = book_of({"S": {"A1": 2.0, "A2": 3.0, "B1": ("=A1*A2+1", 7.0), "B2": ('=IF(B1>5,"big","small")', "big"),
                          "B3": ("=SUM(A1:A2)", 5.0), "B4": ("=B1/0", "#DIV/0!")}})
    got = {v.key: v.status for v in postgres.verify(book, compile_book(book), con)}
    assert set(got.values()) == {"PASS"}


def test_what_postgres_cannot_run_is_unresolved_with_every_cell_reading_it(con, monkeypatch):
    book = book_of({"S": {"A1": 2.0, "B1": ("=A1*3", 6.0), "C1": ("=B1+1", 7.0), "D1": ("=A1+10", 12.0)}})
    import cellproof.compiler as compiler
    monkeypatch.setattr(compiler, "BATCH", 1)
    compiled = compile_book(book)
    real = postgres.translate
    monkeypatch.setattr(postgres, "translate",
                        lambda sql: real(sql).replace("3.0", "no_such_column") if "3.0" in sql else real(sql))
    got = {v.key: v for v in postgres.verify(book, compiled, con)}
    assert got[("S", 1, 2)].status == "UNRESOLVED" and "Postgres refused" in got[("S", 1, 2)].reason
    assert got[("S", 1, 3)].status == "UNRESOLVED" and "depends on S!B1" in got[("S", 1, 3)].reason
    assert got[("S", 1, 4)].status == "PASS"


def test_cli_target_postgres(tmp_path):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2.0, "B1": ("=A1*3", 6.0)}})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out"), "--target", "postgres"]) == 0
    report = json.loads((tmp_path / "out" / "report.json").read_text(encoding="utf-8"))
    assert any("verified on Postgres" in f for f in report["findings"])


def test_list_indexes_survive_translation(con):
    # sqlglot's index offsets once turned xs[LEAST(i + 1, n)] into xs[LEAST(...) + -1]: PERCENTILE took the wrong
    # element on Postgres (formulas' test.xlsx, caught by the second verification)
    book = book_of({"S": {"A1": 1.0, "A2": 3.0, "A3": 2.0, "A4": 4.0, "B1": ("=PERCENTILE(A1:A4,0.3)", 1.9),
                          "B2": ("=QUARTILE(A1:A4,3)", 3.25), "B3": ("=LARGE(A1:A4,2)", 3.0),
                          "B4": ('=SUBSTITUTE("a-b-c","-","+",2)', "a-b+c")}})
    got = {v.key: (v.status, v.actual) for v in postgres.verify(book, compile_book(book), con)}
    assert {s for s, _ in got.values()} == {"PASS"}, got


def test_duckdb_substr_semantics_on_postgres(con):
    import duckdb
    con.execute(postgres.PRELUDE)
    for a, n in [(3, -1), (3, 0), (0, 2), (-1, 2), (-3, 2), (1, -5), (2, 10), (6, 2), (0, -1), (-10, 3), (2, None)]:
        q = f"SELECT substr('hello', {a}, {n})" if n is not None else f"SELECT substr('hello', {a})"
        assert con.execute(postgres.translate(q)).fetchone()[0] == duckdb.sql(q).fetchall()[0][0], (a, n)


def test_rounding_ties_go_away_from_zero_on_postgres(con):
    # Postgres's round(double) rounds 100.5 to 100 (ties to even); Excel, DuckDB and these cells go away from zero
    book = book_of({"S": {"A1": ("=ROUND(2.5,0)", 3.0), "A2": ("=ROUND(-2.5,0)", -3.0), "A3": ("=ROUND(1.005,2)", 1.01),
                          "A4": ("=ROUND(2.675,2)", 2.68), "A5": ("=MROUND(5,2)", 6.0), "A6": ("=ROUND(0.5,0)", 1.0)}})
    got = {v.key: (v.status, v.actual) for v in postgres.verify(book, compile_book(book), con)}
    assert {s for s, _ in got.values()} == {"PASS"}, got


@pytest.mark.parametrize("sql", [
    "SELECT list_sum(range(1, 5)) AS v",                    # range as a list
    "SELECT list_sum(range(10, 1, -3)) AS v",                # negative step stops before the end
    "SELECT sum(x) AS v FROM range(1, 5) AS g(x)",           # range in FROM
    "SELECT regexp_extract('hello', '(?s)^(.*?)(l)', 1) AS v",
    "SELECT regexp_extract('hello', 'z', 1) AS v",           # no match: '' as DuckDB, not NULL
    "SELECT array_to_string(list_transform(list_zip(string_split('ab cD-éx', ''), list_concat([' '], "
    "string_split('ab cD-éx', ''))), lambda x: CASE WHEN regexp_matches(x[2], '\\p{L}') THEN lower(x[1]) "
    "ELSE upper(x[1]) END), '') AS v",                       # PROPER
    "SELECT len(string_split('', '')) AS v",                 # '' splits into ['']
])
def test_rewrites_match_duckdb(con, sql):
    import duckdb
    assert con.execute(postgres.translate(sql)).fetchall()[0][0] == duckdb.sql(sql).fetchall()[0][0]


def test_proper_verifies_on_postgres(con):
    book = book_of({"S": {"A1": "hello wORLD o'neil", "B1": ("=PROPER(A1)", "Hello World O'Neil"),
                          "B2": ('=FIND("l",A1)', 3.0), "B3": ('=SEARCH("w*d",A1)', 7.0)}})
    got = {v.key: (v.status, v.actual) for v in postgres.verify(book, compile_book(book), con)}
    assert {s for s, _ in got.values()} == {"PASS"}, got



def test_boolean_index_and_error_spill_on_postgres(con, tmp_path):
    # formulas' test.xlsx: INDEX of a TRUE/FALSE cell read max(bool) (no such function on Postgres); a spill that is
    # one error typed its num column as text (CASE text/double mismatch)
    from cellproof.reader import load
    path = tmp_path / "b.xlsx"
    write_xlsx(path, {"S": {"A1": True, "B1": False, "C1": ("=INDEX(A1:B1,1,1)", True),
                            "D1": ("=INDEX(A1:B1,1,2)", False),
                            "E1": ("{=_xlfn._xlws.SORT(A1:B1,-1,1)}", "#VALUE!", "E1", "dynamic")}})
    book = load(path)
    got = {v.key: (v.status, v.actual, v.reason) for v in postgres.verify(book, compile_book(book), con)}
    assert {s for s, _, _ in got.values()} == {"PASS"}, got


# --- 1.0: the dbt project on Postgres ---------------------------------------------------------------------------

# A workbook that needs the Postgres helpers (MOD -> cellproof_fmod, TEXT and & -> cellproof_printf), a region (one
# formula pasted down), workbook text with braces (Jinja must not read them) and an error.
DBT_BOOK = {"S": {
    "A1": 10.0, "A2": 20.0, "A3": 30.0,
    "B1": ("=A1*1.5", 15.0), "B2": ("=A2*1.5", 30.0), "B3": ("=A3*1.5", 45.0),
    "C1": ("=MOD(A1,3)", 1.0), "C2": ('=TEXT(A2/3,"0.00")', "6.67"), "C3": ('="{"&A3&"}"', "{30}"),
    "E1": ("=SUM(B1:B3)/0", "#DIV/0!"),
}}


def test_verify_book_drops_what_postgres_cannot_run_and_every_reader(con, monkeypatch):
    book = book_of({"S": {"A1": 2.0, "B1": ("=A1*3", 6.0), "C1": ("=B1+1", 7.0), "D1": ("=A1+10", 12.0)}})
    import cellproof.compiler as compiler
    monkeypatch.setattr(compiler, "BATCH", 1)
    real = postgres.translate
    monkeypatch.setattr(postgres, "translate",
                        lambda sql: real(sql).replace("3.0", "no_such_column") if "3.0" in sql else real(sql))
    compiled, verdicts = postgres.verify_book(book, compile_book, con)
    got = {v.key: v for v in verdicts}
    assert got[("S", 1, 2)].status == "UNRESOLVED" and "Postgres refused" in got[("S", 1, 2)].reason
    assert got[("S", 1, 3)].status == "UNRESOLVED"
    assert got[("S", 1, 4)].status == "PASS"
    in_batches = {s.key for b in compiled.batches for s in b}
    assert ("S", 1, 2) not in in_batches and ("S", 1, 3) not in in_batches  # the dbt project holds only what ran


def test_the_postgres_dbt_project(tmp_path):
    from cellproof.emitter import write_project
    book = book_of(DBT_BOOK)
    write_project(book, compile_book(book), tmp_path, "book", target="postgres")
    assert "type: postgres" in (tmp_path / "profiles.yml").read_text()
    project = (tmp_path / "dbt_project.yml").read_text()
    assert "cellproof_helpers()" in project and "double precision" in project
    helpers = (tmp_path / "macros" / "cellproof_helpers.sql").read_text()
    assert "pg_temp" not in helpers and "{{ target.schema }}.{% raw %}cellproof_fmod" in helpers
    models = "\n".join(p.read_text() for p in (tmp_path / "models").rglob("*.sql"))
    assert "pg_temp" not in models.lower() and "{{ target.schema }}.cellproof_fmod" in models.lower()
    assert "{{ ref('input_cells') }}" in models and "'{{ '{' }}'" in models  # the formula's "{" stays text  # braces stay literal


def _dbt_postgres():
    import shutil
    import subprocess
    exe = shutil.which("dbt")
    return exe is not None and "postgres" in subprocess.run([exe, "--version"], capture_output=True, text=True).stdout


@pytest.mark.skipif(not _dbt_postgres(), reason="dbt-postgres not installed")
def test_dbt_build_on_postgres_recomputes_the_workbook(con, tmp_path):
    from cellproof.emitter import run_dbt, write_project
    from cellproof.verifier import matches
    book = book_of(DBT_BOOK)
    compiled = compile_book(book)
    write_project(book, compiled, tmp_path, "book", target="postgres")
    ok, log = run_dbt(tmp_path, env=postgres.dbt_env(con, schema="cellproof_test"))
    assert ok, log
    rows = con.execute("SELECT sheet, row_num, col_num, kind, num, txt FROM cellproof_test.workbook_cells").fetchall()
    got = {(s, r, c): (k, n, t) for s, r, c, k, n, t in rows}
    assert set(got) == {s.key for s in compiled.steps}
    for key, (k, n, t) in got.items():
        assert matches(book.cells[key], k, n, t), (key, k, n, t)
    con.execute("DROP SCHEMA cellproof_test CASCADE")


@pytest.mark.skipif(not _dbt_postgres(), reason="dbt-postgres not installed")
def test_cli_target_postgres_with_dbt(tmp_path, capsys):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, DBT_BOOK)
    assert main(["convert", str(wb), "--out", str(tmp_path / "out"), "--target", "postgres", "--dbt"]) == 0
    printed = capsys.readouterr().out
    assert "7 pass, 0 fail, 0 unresolved" in printed and "dbt build: ok" in printed
    assert "type: postgres" in (tmp_path / "out" / "profiles.yml").read_text()


def test_cli_postgres_url_that_does_not_connect(tmp_path, capsys):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2.0, "B1": ("=A1*3", 6.0)}})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out"), "--target", "postgres",
                 "--pg-url", "postgresql://nobody@127.0.0.1:1/none?connect_timeout=2"]) == 2
    assert "could not connect to Postgres" in capsys.readouterr().err
