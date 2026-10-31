import io
import json
import sys

from xlsxkit import SALES, write_xlsx

from cellproof.cli import main

COMMISSION = {
    "A1": "Rep", "B1": "Total", "C1": "Bonus",
    "A2": "Asha", "B2": ("=SUMIFS(Sales!C:C, Sales!A:A, A2)", 120000), "C2": ("=IF(B2>100000, B2*5%, B2*2%)", 6000),
    "A3": "Vikram", "B3": ("=SUMIFS(Sales!C:C, Sales!A:A, A3)", 60000), "C3": 1500,
    "B5": ("=SUM(B2:B3)", 180000), "C5": ("=TODAY()", 46000),
}


def test_convert_commission_end_to_end(tmp_path, capsys):
    wb = tmp_path / "commission.xlsx"
    write_xlsx(wb, {"Sales": SALES, "Commission": COMMISSION})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out")]) == 0
    assert "4 pass, 0 fail, 1 unresolved" in capsys.readouterr().out
    md = (tmp_path / "out" / "REPORT.md").read_text(encoding="utf-8")
    assert "Sales!C5: number stored as text ('15,000')" in md
    assert "| Commission!C5 | volatile function TODAY |" in md
    assert (tmp_path / "out" / "models" / "workbook_cells.sql").exists()


def test_failures_exit_1(tmp_path):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2, "A2": ("=A1*2", 5)}})
    assert main(["convert", str(wb)]) == 1
    assert (tmp_path / "b_dbt" / "REPORT.md").exists()


def test_unreadable_file_exits_2(tmp_path, capsys):
    bad = tmp_path / "x.xls"
    bad.write_bytes(b"x")
    assert main(["convert", str(bad)]) == 2
    assert "save the workbook as .xlsx" in capsys.readouterr().err


def test_console_that_cannot_print_unicode(tmp_path, monkeypatch):
    wb = tmp_path / "Ümsatz✓.xlsx"
    write_xlsx(wb, {"S": {"A1": 1, "A2": ("=A1+1", 2)}})
    monkeypatch.setattr(sys, "stdout", io.TextIOWrapper(io.BytesIO(), encoding="cp1252"))
    assert main(["convert", str(wb)]) == 0


def _region_sheet(n=6, extra=None):
    spec = {"A1": "amount", "B1": "tax"}
    for r in range(2, n + 2):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", float(r * 2))
        if extra:
            spec[f"C{r}"] = (extra.format(r=r), 0.0)
    return spec


def test_report_lists_regions(tmp_path):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": _region_sheet()})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out")]) == 0
    md = (tmp_path / "out" / "REPORT.md").read_text(encoding="utf-8")
    assert "## Regions" in md and "model region_s_2_7" in md


def test_report_says_why_a_column_stays_in_cell_mode(tmp_path):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": _region_sheet(extra="=MAX(A{r},1)")})
    main(["convert", str(wb), "--out", str(tmp_path / "out")])
    md = (tmp_path / "out" / "REPORT.md").read_text(encoding="utf-8")
    assert "column C rows 2-7 stays in cell mode: reads a cell reference directly" in md
    assert "__ROWREF" not in md


def test_duplicate_model_names_are_unique_in_the_report_and_on_disk(tmp_path):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"My Sheet": _region_sheet(), "My_Sheet": _region_sheet()})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out")]) == 0
    md = (tmp_path / "out" / "REPORT.md").read_text(encoding="utf-8")
    models = {p.stem for p in (tmp_path / "out" / "models" / "cellproof_cells").glob("region_*.sql")}
    assert len(models) == 2
    for m in models:
        assert f"model {m}" in md


def test_sql_that_duckdb_rejects_exits_3_with_the_cell(tmp_path, capsys, monkeypatch):
    import cellproof.verifier as verifier
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 1, "B1": ("=A1+1", 2)}})
    monkeypatch.setattr(verifier, "batch_rows", lambda batch, prev: "SELECT no_such_column")
    assert main(["convert", str(wb), "--out", str(tmp_path / "out")]) == 3
    assert "S!B1" in capsys.readouterr().err


def test_sql_duckdb_rejects_makes_its_cells_and_their_readers_unresolved_not_the_run(tmp_path, monkeypatch, capsys):
    # a bug in generated SQL (like 0.3.1's SUBSTITUTE on DuckDB 2.0) used to stop the whole run with no report
    import cellproof.compiler as compiler
    import cellproof.verifier as verifier
    from xlsxkit import write_xlsx
    monkeypatch.setattr(compiler, "BATCH", 1)  # one cell per batch: B1's batch fails alone
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2.0, "B1": ("=A1*3", 6.0), "C1": ("=B1+1", 7.0), "D1": ("=A1+10", 12.0)}})
    real = verifier.batch_rows

    def broken(batch, prev):
        sql = real(batch, prev)
        return sql.replace("* CAST(3.0 AS DOUBLE)", "* no_such_column")  # only the batch holding B1 breaks
    monkeypatch.setattr(verifier, "batch_rows", broken)
    out = tmp_path / "out"
    assert main(["convert", str(wb), "--out", str(out)]) == 3  # still says: this is a cellproof bug
    report = json.loads((out / "report.json").read_text(encoding="utf-8"))
    status = {c["cell"]: (c["status"], c["reason"]) for c in report["cells"]}
    assert status["S!B1"][0] == "UNRESOLVED" and "DuckDB could not run" in status["S!B1"][1]
    assert status["S!C1"][0] == "UNRESOLVED" and status["S!D1"] == ("PASS", "")
    assert "please report" in capsys.readouterr().err


def test_a_rejected_region_or_spill_is_not_rebuilt_forever(tmp_path, monkeypatch):
    import cellproof.verifier as verifier
    from test_dynamic import DATA, spill
    from xlsxkit import write_xlsx
    s = {f"A{r}": float(r) for r in range(1, 30)}
    s.update({f"B{r}": (f"=A{r}*7", 7.0 * r) for r in range(1, 30)})  # a region
    s.update(spill("D1", "_xlfn._xlws.SORT(A1:A3)", [[1.0], [2.0], [3.0]]))
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": s})
    real = verifier.batch_rows
    monkeypatch.setattr(verifier, "batch_rows", lambda b, p: real(b, p).replace("CAST(7.0 AS DOUBLE)", "nope")
                        .replace("ORDER BY", "ORDER BY nope,", 1))
    out = tmp_path / "out"
    assert main(["convert", str(wb), "--out", str(out)]) == 3
    report = json.loads((out / "report.json").read_text(encoding="utf-8"))
    assert {c["status"] for c in report["cells"]} == {"UNRESOLVED"}
