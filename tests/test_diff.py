"""`cellproof diff`: what changed between two versions of a workbook, and which of it broke the answers."""
from __future__ import annotations

import json

from xlsxkit import write_xlsx

from cellproof.cli import main
from cellproof.diff import compare


def _book(spec, names=None):
    """A book, as the compiler sees it, without going through a file (book_of has no calc/app fields).

    The reader is what fills Book.calc and Book.app, and `diff` reports the cache warning from them, so the
    end-to-end tests write a real .xlsx; this is only for the comparison itself."""
    from xlsxkit import book_of
    return book_of(spec, names)


# --- cells added, removed, changed ---------------------------------------------------------------

def test_added_removed_and_changed_inputs():
    old = _book({"S": {"A1": 2, "A2": 3, "B1": ("=SUM(A1:A2)", 5)}})
    new = _book({"S": {"A1": 2, "A3": 4, "B1": ("=SUM(A1:A2)", 5)}})
    d = compare(old, new)
    assert d["added"] == ["S!A3"]
    assert d["removed"] == ["S!A2"]
    # B1's text did not change, but A2 (which its range covered) is gone: the range's outside now differs
    assert "S!A1" not in d["inputs"]
    assert d["counts"] == {"added": 1, "removed": 1, "inputs": 0, "text": 0, "sql": 0, "verdicts": 0}


def test_an_input_whose_saved_value_changed_is_named_with_both_values():
    old = _book({"S": {"A1": 2, "B1": ("=A1*2", 4)}})
    new = _book({"S": {"A1": 5, "B1": ("=A1*2", 10)}})
    d = compare(old, new)
    assert d["inputs"] == [{"cell": "S!A1", "was": 2.0, "now": 5.0}]
    assert d["text"] == [] and d["sql"] == []


def test_a_saved_value_that_moves_kind_counts_as_changed():
    old = _book({"S": {"A1": 2}})
    new = _book({"S": {"A1": "2"}})
    d = compare(old, new)
    assert [(i["cell"], i["was"], i["now"]) for i in d["inputs"]] == [("S!A1", 2.0, "2")]


def test_a_formula_that_became_an_input_cell_is_both_removed_from_formulas_and_an_input_change():
    old = _book({"S": {"A1": ("=1+1", 2), "B1": ("=A1+1", 3)}})
    new = _book({"S": {"A1": 2, "B1": ("=A1+1", 3)}})
    d = compare(old, new)
    assert d["added"] == [] and d["removed"] == []
    assert d["text"] == [{"cell": "S!A1", "was": "=1+1", "now": None}]  # it lost its formula

# --- formula text vs SQL -------------------------------------------------------------------------

def test_a_formula_whose_text_changed_is_reported_with_both_texts():
    old = _book({"S": {"A1": 2, "B1": ("=A1+1", 3)}})
    new = _book({"S": {"A1": 2, "B1": ("=A1*2", 4)}})
    d = compare(old, new)
    assert d["text"] == [{"cell": "S!B1", "was": "=A1+1", "now": "=A1*2"}]
    assert d["sql"] == []  # its text changed, so it is not a "same text, moved inputs" case


def test_a_formula_whose_sql_changed_but_text_did_not_is_the_surprising_case():
    """A named range that grows to cover one more row: the text is the same, what it reads is not."""
    old = _book({"S": {"A1": 1, "A2": 2, "B1": ("=SUM(Rate)", 3)}}, {"Rate": ("S", "A1:A2")})
    new = _book({"S": {"A1": 1, "A2": 2, "A3": 5, "B1": ("=SUM(Rate)", 8)}}, {"Rate": ("S", "A1:A3")})
    d = compare(old, new)
    assert d["text"] == []
    assert [s["cell"] for s in d["sql"]] == ["S!B1"]


def test_an_unrelated_formula_elsewhere_is_not_reported_as_a_sql_change():
    """Batch CTE names renumber when the dependency order moves: that must not read as cells changing."""
    base = {"A1": 2, "A2": 3, "B1": ("=A1+1", 3), "B2": ("=B1*2", 6), "B3": ("=B2+A2", 9),
            "C1": ("=B3&\"!\"", "9!"), "D1": ("=IF(C1=\"9!\", 1, 0)", 1)}
    old = _book({"S": dict(base)})
    new = _book({"S": dict(base, **{"Q7": ("=1+1", 2)})})
    d = compare(old, new)
    assert d["sql"] == []
    assert d["text"] == []
    assert d["added"] == ["S!Q7"]


def test_a_batch_boundary_moving_does_not_shift_the_canonical_sql():
    """A long chain crosses batches; adding weight elsewhere changes where they split, not what a cell reads."""
    spec = {"A1": 5}
    for i in range(2, 51):
        spec[f"A{i}"] = (f"=A{i-1}+1", 4 + i)
    spec["B1"] = ("=SUM(A1:A50)", 1275)
    old = _book({"S": dict(spec)})
    new = _book({"S": dict(spec, **{"Z1": ('="' + "x" * 6000 + '"', "x" * 6000)})})
    d = compare(old, new)
    assert d["sql"] == []


def test_a_region_column_becoming_cell_mode_does_not_shift_sql_either():
    """Regions change how cells are compiled, not what they read: the SQL comparison uses cell mode on both."""
    spec = {"A1": 1, "A2": 2, "A3": 3, "B1": ("=A1*2", 2), "B2": ("=A2*2", 4), "B3": ("=A3*2", 6)}
    old = _book({"S": dict(spec)})
    new = _book({"S": dict(spec, **{"B4": ("=A4*2", 8), "A4": 4})})
    d = compare(old, new)
    assert d["sql"] == []  # B1-B3 still read the same cells


def test_nothing_changed_says_so_plainly():
    spec = {"A1": 2, "B1": ("=A1+1", 3)}
    d = compare(_book({"S": dict(spec)}), _book({"S": dict(spec)}))
    assert d["added"] == [] and d["removed"] == [] and d["inputs"] == []
    assert d["text"] == [] and d["sql"] == [] and d["verdicts"] == []
    assert d["summary"] == "the two workbooks answer the same way"

# --- verdicts -------------------------------------------------------------------------------------

def test_a_new_fail_is_the_headline_and_sets_exit_1(tmp_path):
    old = tmp_path / "old.xlsx"
    new = tmp_path / "new.xlsx"
    write_xlsx(old, {"S": {"A1": 2, "B1": ("=A1*2", 4)}})
    write_xlsx(new, {"S": {"A1": 2, "B1": ("=A1*3", 4)}})  # text changed, saved value not: a new FAIL
    out = tmp_path / "out"
    assert main(["diff", str(old), str(new), "--out", str(out)]) == 1
    d = json.loads((out / "diff.json").read_text(encoding="utf-8"))
    assert [v["cell"] for v in d["verdicts"]] == ["S!B1"]
    assert (d["verdicts"][0]["was"], d["verdicts"][0]["now"]) == ("PASS", "FAIL")
    md = (out / "DIFF.md").read_text(encoding="utf-8")
    assert md.index("S!B1") < md.index("## ")


def test_a_difference_alone_is_not_a_failure(tmp_path):
    old = tmp_path / "old.xlsx"
    new = tmp_path / "new.xlsx"
    write_xlsx(old, {"S": {"A1": 2, "B1": ("=A1*2", 4)}})
    write_xlsx(new, {"S": {"A1": 2, "B1": ("=A1*2", 4), "C1": 7}})  # a new input cell, nothing broken
    out = tmp_path / "out"
    assert main(["diff", str(old), str(new), "--out", str(out)]) == 0
    d = json.loads((out / "diff.json").read_text(encoding="utf-8"))
    assert d["added"] == ["S!C1"] and d["verdicts"] == []


def test_a_fixed_fail_is_reported_but_does_not_set_exit_1(tmp_path):
    old = tmp_path / "old.xlsx"
    new = tmp_path / "new.xlsx"
    write_xlsx(old, {"S": {"A1": 2, "B1": ("=A1*3", 4)}})
    write_xlsx(new, {"S": {"A1": 2, "B1": ("=A1*2", 4)}})
    out = tmp_path / "out"
    assert main(["diff", str(old), str(new), "--out", str(out)]) == 0
    d = json.loads((out / "diff.json").read_text(encoding="utf-8"))
    assert [(v["cell"], v["was"], v["now"]) for v in d["verdicts"]] == [("S!B1", "FAIL", "PASS")]


def test_a_cell_that_stopped_being_translatable_says_why(tmp_path):
    old = tmp_path / "old.xlsx"
    new = tmp_path / "new.xlsx"
    write_xlsx(old, {"S": {"A1": 2, "B1": ("=A1+1", 3)}})
    write_xlsx(new, {"S": {"A1": 2, "B1": ("=TODAY()", 46000)}})
    out = tmp_path / "out"
    assert main(["diff", str(old), str(new), "--out", str(out)]) == 0
    d = json.loads((out / "diff.json").read_text(encoding="utf-8"))
    assert d["verdicts"][0]["now"] == "UNRESOLVED"
    assert d["verdicts"][0]["reason"]

# --- sheets ---------------------------------------------------------------------------------------

def test_a_renamed_sheet_reports_everything_removed_and_added(tmp_path, capsys):
    old = tmp_path / "old.xlsx"
    new = tmp_path / "new.xlsx"
    write_xlsx(old, {"Sales": {"A1": 2, "B1": ("=A1*2", 4)}})
    write_xlsx(new, {"Revenue": {"A1": 2, "B1": ("=A1*2", 4)}})
    out = tmp_path / "out"
    assert main(["diff", str(old), str(new), "--out", str(out)]) == 0
    d = json.loads((out / "diff.json").read_text(encoding="utf-8"))
    assert d["removed"] == ["Sales!A1", "Sales!B1"]
    assert d["added"] == ["Revenue!A1", "Revenue!B1"]
    assert d["verdicts"] == []

# --- the write-out --------------------------------------------------------------------------------

def test_the_summary_goes_to_the_console_and_a_short_diff_writes_no_markdown(tmp_path, capsys):
    old = tmp_path / "old.xlsx"
    new = tmp_path / "new.xlsx"
    write_xlsx(old, {"S": {"A1": 2, "B1": ("=A1+1", 3)}})
    write_xlsx(new, {"S": {"A1": 2, "B1": ("=A1+1", 3), "C1": 9}})
    out = tmp_path / "out"
    main(["diff", str(old), str(new), "--out", str(out)])
    said = capsys.readouterr().out
    assert "1 added" in said or "added 1" in said
    assert (out / "diff.json").exists()
    assert not (out / "DIFF.md").exists()


def test_a_bigger_diff_writes_markdown_behind_a_new_fail(tmp_path):
    old = tmp_path / "old.xlsx"
    new = tmp_path / "new.xlsx"
    write_xlsx(old, {"S": {"A1": 2, "B1": ("=A1*2", 4), "C1": ("=B1+1", 5)}})
    write_xlsx(new, {"S": {"A1": 2, "B1": ("=A1*3", 4), "C1": ("=B1+1", 5)}})
    out = tmp_path / "out"
    main(["diff", str(old), str(new), "--out", str(out)])
    md = (out / "DIFF.md").read_text(encoding="utf-8")
    assert md.index("S!B1") < md.index("S!C1")  # the new FAIL comes before the cell that reads it
    assert "Excel" in md and "SQL" in md


def test_the_diff_never_calls_a_cell_stale(tmp_path):
    old = tmp_path / "old.xlsx"
    new = tmp_path / "new.xlsx"
    write_xlsx(old, {"S": {"A1": 2, "B1": ("=A1*2", 4)}}, calc={"calcMode": "manual", "calcId": 191029})
    write_xlsx(new, {"S": {"A1": 3, "B1": ("=A1*2", 8)}}, calc={"calcMode": "manual", "calcId": 191029})
    out = tmp_path / "out"
    main(["diff", str(old), str(new), "--out", str(out)])
    for text in [(out / "diff.json").read_text(encoding="utf-8"),
                 (out / "DIFF.md").read_text(encoding="utf-8") if (out / "DIFF.md").exists() else ""]:
        assert "stale" not in text.lower()
