"""`cellproof prove`: a bundle whose claim can be checked without cellproof, and `--check`, which trusts none of it."""
import hashlib
import json

import duckdb
from xlsxkit import book_of, write_xlsx

from cellproof import proof
from cellproof.cli import main
from cellproof.compiler import compile_book
from cellproof.verifier import execute

SALES = {"A1": "Rep", "B1": "Amount",
         "A2": "Asha", "B2": 80000,
         "A3": "Vikram", "B3": 60000}
COMMISSION = {"A1": "Rep", "B1": "Total", "C1": "Bonus",
              "A2": "Asha", "B2": ("=SUMIFS(Sales!B:B, Sales!A:A, A2)", 80000),
              "A3": "Vikram", "B3": ("=SUMIFS(Sales!B:B, Sales!A:A, A3)", 60000),
              "C2": ("=IF(B2>100000, B2*5%, B2*2%)", 1600)}


def _book(tmp_path, name="b.xlsx", calc=None, app=""):
    wb = tmp_path / name
    write_xlsx(wb, {"Sales": SALES, "Commission": COMMISSION}, calc=calc, app=app)
    return wb


def _bundle(tmp_path, capsys, wb, *extra):
    assert main(["prove", str(wb), *extra]) == 0
    capsys.readouterr()
    if "--out" in extra:
        return tmp_path / extra[extra.index("--out") + 1]
    return tmp_path / (wb.stem + "_proof")


def _manifest(out):
    return json.loads((out / "proof.json").read_text(encoding="utf-8"))


def _region_book():
    """A region (B2:B6 = A*2 down the rows) and one cell on top of it."""
    sheet = {"A1": "amount", "B1": "tax", "D1": 0.05}
    for r in range(2, 7):
        sheet[f"A{r}"] = float(r)
        sheet[f"B{r}"] = (f"=A{r}*2", float(r * 2))
    sheet["E1"] = ("=SUM(B2:B6)*D1", 2.0)
    return book_of({"S": sheet})


# --- cells.sql is the proof ----------------------------------------------------------------------

def test_cells_sql_answers_exactly_what_the_verifier_answers():
    """The script must need nothing but DuckDB, and must answer what the verifier answered - not its own idea."""
    book = _region_book()
    compiled = compile_book(book)
    got = execute(book, compiled)
    rows = duckdb.connect().execute(proof.cells_sql(book, compiled)).fetchall()
    mine = {(s, r, c): (k, n, t) for s, r, c, k, n, t in rows if book.cells[(s, r, c)].formula is not None}
    assert set(mine) == set(got)
    assert all(mine[key] == got[key] for key in got)


def test_cells_sql_seeds_the_input_cells_the_workbook_holds():
    book = _region_book()
    rows = duckdb.connect().execute(proof.cells_sql(book, compile_book(book))).fetchall()
    got = {(s, r, c) for s, r, c, _, _, _ in rows}
    assert {("S", 1, 1), ("S", 1, 4)} <= got  # the labels and D1: input cells, not formula cells


# --- the bundle ----------------------------------------------------------------------------------

def test_proof_writes_the_four_files_and_a_sealed_manifest(tmp_path, capsys):
    wb = _book(tmp_path)
    out = _bundle(tmp_path, capsys, wb)
    assert sorted(p.name for p in out.iterdir()) == ["REPORT.md", "cells.sql", "proof.json", "report.json"]
    manifest = _manifest(out)
    assert manifest["format"] == "cellproof-proof/1"
    assert manifest["workbook"] == {"file": "b.xlsx", "bytes": wb.stat().st_size,
                                    "sha256": hashlib.sha256(wb.read_bytes()).hexdigest()}
    assert manifest["engine"] == {"name": "duckdb", "version": duckdb.__version__}
    assert set(manifest["files"]) == {"cells.sql", "REPORT.md", "report.json"}
    assert manifest["counts"] == {"PASS": 3, "FAIL": 0, "UNRESOLVED": 0}
    assert manifest["computed_answers"] == 0
    sealed = "".join(f"{name}:{manifest['files'][name]}\n" for name in sorted(manifest["files"]))
    assert manifest["bundle"] == hashlib.sha256(sealed.encode("utf-8")).hexdigest()


def test_the_report_says_the_seal_is_not_a_signature(tmp_path, capsys):
    out = _bundle(tmp_path, capsys, _book(tmp_path))
    assert "not signed" in (out / "REPORT.md").read_text(encoding="utf-8")


def test_two_runs_differ_only_in_created(tmp_path, capsys):
    wb = _book(tmp_path)
    first = _bundle(tmp_path, capsys, wb, "--out", str(tmp_path / "one"))
    second = _bundle(tmp_path, capsys, wb, "--out", str(tmp_path / "two"))
    a, b = _manifest(first), _manifest(second)
    assert a.pop("created") and b.pop("created")
    assert a == b
    assert (first / "cells.sql").read_bytes() == (second / "cells.sql").read_bytes()


def test_the_manifest_records_what_the_file_says_about_its_own_values(tmp_path, capsys):
    wb = _book(tmp_path, calc={"calcMode": "manual", "fullCalcOnLoad": True, "calcId": 191029},
               app="Microsoft Excel")
    assert _manifest(_bundle(tmp_path, capsys, wb))["saved_values"] == {
        "calc_mode": "manual", "full_calc_on_load": True, "calc_id": 191029,
        "last_saved_by": "Microsoft Excel"}


def test_a_file_that_says_nothing_is_null_not_guessed(tmp_path, capsys):
    assert _manifest(_bundle(tmp_path, capsys, _book(tmp_path)))["saved_values"] == {
        "calc_mode": None, "full_calc_on_load": None, "calc_id": None, "last_saved_by": ""}


def test_an_untranslated_cell_is_a_hole_in_the_script_not_a_short_answer(tmp_path, capsys):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2, "A2": ("=BESSELJ(1,1)", 0.5)}})
    assert main(["prove", str(wb)]) == 0  # an unresolved cell is not a FAIL
    out = tmp_path / "b_proof"
    assert "could not be translated" in (out / "cells.sql").read_text(encoding="utf-8")
    assert _manifest(out)["counts"] == {"PASS": 0, "FAIL": 0, "UNRESOLVED": 1}
    capsys.readouterr()
    assert main(["prove", "--check", str(out)]) == 0
    assert "1 cell(s) could not be translated" in capsys.readouterr().out


def test_a_fail_does_not_stop_the_bundle(tmp_path, capsys):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2, "A2": ("=A1*2", 5)}})
    assert main(["prove", str(wb)]) == 1
    out = tmp_path / "b_proof"
    assert _manifest(out)["counts"]["FAIL"] == 1
    capsys.readouterr()
    assert main(["prove", "--check", str(out)]) == 0  # a FAIL is a verdict, not a broken bundle


def test_an_apostrophe_sheet_and_non_ascii_text_survive_the_bundle(tmp_path, capsys):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"Bob's": {"A1": "naive - ok", "A2": ('=A1&"!"', "naive - ok!"),
                             "B1": ("=LEN(" + '"caf\u00e9"' + ")", 4)}})
    assert main(["prove", str(wb)]) == 0
    out = tmp_path / "b_proof"
    assert "Bob''s" in (out / "cells.sql").read_text(encoding="utf-8")  # the SQL quotes it, and quotes it safely
    capsys.readouterr()
    assert main(["prove", "--check", str(out)]) == 0
    assert capsys.readouterr().out.isascii()


# --- --check -------------------------------------------------------------------------------------

def test_check_passes_on_a_fresh_bundle(tmp_path, capsys):
    out = _bundle(tmp_path, capsys, _book(tmp_path))
    assert main(["prove", "--check", str(out)]) == 0
    text = capsys.readouterr().out
    assert "3 files intact" in text
    assert "cells.sql reproduces 3 of 3 answers" in text
    assert "the bundle checks out" in text
    assert text.isascii()


def test_check_also_matches_the_workbook_when_given_one(tmp_path, capsys):
    wb = _book(tmp_path)
    out = _bundle(tmp_path, capsys, wb)
    assert main(["prove", "--check", str(out), "--workbook", str(wb)]) == 0
    text = capsys.readouterr().out
    assert "workbook hash matches" in text
    assert "3 saved values match the report" in text


def test_check_refuses_another_workbook(tmp_path, capsys):
    wb = _book(tmp_path)
    out = _bundle(tmp_path, capsys, wb)
    other = tmp_path / "other.xlsx"
    write_xlsx(other, {"Sales": SALES})
    assert main(["prove", "--check", str(out), "--workbook", str(other)]) == 2
    assert "not the workbook this bundle was made from" in capsys.readouterr().err


def test_check_refuses_a_bundle_whose_file_changed(tmp_path, capsys):
    out = _bundle(tmp_path, capsys, _book(tmp_path))
    (out / "REPORT.md").write_text("edited\n", encoding="utf-8")
    assert main(["prove", "--check", str(out)]) == 2
    assert "REPORT.md was changed after the bundle was written" in capsys.readouterr().err


def test_check_refuses_a_bundle_with_a_missing_file(tmp_path, capsys):
    out = _bundle(tmp_path, capsys, _book(tmp_path))
    (out / "cells.sql").unlink()
    assert main(["prove", "--check", str(out)]) == 2
    assert "cells.sql is missing" in capsys.readouterr().err


def test_check_refuses_an_unknown_format(tmp_path, capsys):
    out = _bundle(tmp_path, capsys, _book(tmp_path))
    manifest = _manifest(out)
    manifest["format"] = "cellproof-proof/9"
    (out / "proof.json").write_text(json.dumps(manifest), encoding="utf-8")
    assert main(["prove", "--check", str(out)]) == 2
    assert "is not a cellproof-proof/1 bundle" in capsys.readouterr().err


def test_check_catches_an_answer_that_moved_even_when_resealed(tmp_path, capsys):
    """The seal is not a signature: a forger can reseal. What it still cannot hide is cells.sql's own answer."""
    wb = _book(tmp_path)
    out = _bundle(tmp_path, capsys, wb)
    data = json.loads((out / "report.json").read_text(encoding="utf-8"))
    data["cells"][0]["actual"] = 999.0
    (out / "report.json").write_text(json.dumps(data), encoding="utf-8")
    proof.seal(out, wb)  # exactly what someone with write access to the folder can do
    assert main(["prove", "--check", str(out)]) == 1
    captured = capsys.readouterr()
    assert data["cells"][0]["cell"] in captured.err
    assert "cells.sql reproduces 2 of 3 answers" in captured.out


def test_check_refuses_a_bundle_that_is_not_there(tmp_path, capsys):
    assert main(["prove", "--check", str(tmp_path / "nope")]) == 2
    assert "cannot read" in capsys.readouterr().err


def test_prove_needs_a_workbook_without_check(tmp_path, capsys):
    assert main(["prove"]) == 2
    assert "prove needs a workbook" in capsys.readouterr().err


def test_prove_refuses_a_workbook_it_cannot_read(tmp_path, capsys):
    bad = tmp_path / "x.xls"
    bad.write_bytes(b"x")
    assert main(["prove", str(bad)]) == 2
    assert "save the workbook as .xlsx" in capsys.readouterr().err
