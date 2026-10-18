import json
import sys

import pytest
from xlsxkit import SALES, write_xlsx

from unexcel.cli import main
from unexcel.compute import fill_missing
from unexcel.reader import load


def test_without_the_formulas_library_cells_stay_unresolved_with_a_hint(tmp_path, monkeypatch):
    monkeypatch.setitem(sys.modules, "formulas", None)  # makes `import formulas` raise ImportError
    wb = tmp_path / "script.xlsx"
    write_xlsx(wb, {"S": {"A1": 2, "A2": ("=A1*2", None)}})
    assert main(["convert", str(wb)]) == 0
    md = (tmp_path / "script_dbt" / "REPORT.md").read_text(encoding="utf-8")
    assert ("| S!A2 | no saved value in the file - open and save it in Excel "
            "(or install unexcel[compute] to compute it) |") in md


def test_sheet_names_with_apostrophes_are_not_computed(tmp_path):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"Bob's": {"A1": 2, "A2": ("=A1*2", None)}})
    _, computed, why = fill_missing(load(wb), wb)
    assert computed == set() and "apostrophe" in why


def test_saved_values_are_never_replaced(tmp_path):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": 2, "A2": ("=A1*2", 5)}})
    book = load(wb)
    assert fill_missing(book, wb) == (book, set(), "")


def test_computes_missing_answers_and_labels_them(tmp_path):
    pytest.importorskip("formulas")
    wb = tmp_path / "script.xlsx"
    write_xlsx(wb, {"Sales": SALES, "C": {
        "A1": "Asha",
        "B1": ("=SUMIFS(Sales!C:C, Sales!A:A, A1)", None),
        "B2": ("=1/0", None),
        "B3": ('=A1&"!"', None),
        "B4": ("=B1>5", None),
    }})
    assert main(["convert", str(wb)]) == 0
    out = tmp_path / "script_dbt"
    assert "4 formula cells had no saved value" in (out / "REPORT.md").read_text(encoding="utf-8")
    cells = json.loads((out / "report.json").read_text(encoding="utf-8"))["cells"]
    assert {c["cell"]: (c["status"], c["answer"]) for c in cells} == {
        f"C!B{r}": ("PASS", "computed") for r in range(1, 5)}


def test_formulas_error_and_empty_tokens_are_not_text():  # both subclass str in formulas 1.3.4
    pytest.importorskip("formulas")
    from formulas.tokens.operand import XlError
    from schedula.utils.gen import Token

    from unexcel.compute import _answer
    assert _answer(XlError("#DIV/0!")) == ("#DIV/0!", "e")
    assert _answer(Token("empty")) == (None, "blank")
    assert _answer("text") == ("text", "s")
