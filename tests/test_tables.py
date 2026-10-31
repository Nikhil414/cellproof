"""Structured references to Excel tables, rewritten to A1 text by the reader (cellproof.tables)."""
import pytest
from xlsxkit import write_xlsx

from cellproof.compiler import compile_book
from cellproof.model import Unsupported
from cellproof.reader import load
from cellproof.tables import Table, to_a1
from cellproof.verifier import verify

# A1:D6: a header row, four data rows, a totals row
SALES = Table("Sales", "S", 1, 1, 6, 4, 1, 1, ("Rep", "Region", "Amount", "Unit Price [USD]"))
T = {"SALES": SALES}


@pytest.mark.parametrize("text, at, want", [
    ("Sales[Amount]", ("X", 9, 9), "'S'!$C$2:$C$5"),
    ("sales[amount]", ("X", 9, 9), "'S'!$C$2:$C$5"),  # names ignore case
    ("Sales[]", ("X", 9, 9), "'S'!$A$2:$D$5"),
    ("Sales[#Data]", ("X", 9, 9), "'S'!$A$2:$D$5"),
    ("Sales[#All]", ("X", 9, 9), "'S'!$A$1:$D$6"),
    ("Sales[#Headers]", ("X", 9, 9), "'S'!$A$1:$D$1"),
    ("Sales[#Totals]", ("X", 9, 9), "'S'!$A$6:$D$6"),
    ("Sales[[#Headers],[Amount]]", ("X", 9, 9), "'S'!$C$1"),
    ("Sales[[#Totals],[Amount]]", ("X", 9, 9), "'S'!$C$6"),
    ("Sales[[#Headers],[#Data],[Amount]]", ("X", 9, 9), "'S'!$C$1:$C$5"),
    ("Sales[[#Data],[#Totals],[Amount]]", ("X", 9, 9), "'S'!$C$2:$C$6"),
    ("Sales[[Rep]:[Amount]]", ("X", 9, 9), "'S'!$A$2:$C$5"),
    ("Sales[[#This Row],[Amount]]", ("S", 3, 4), "$C3"),
    ("Sales[[#This Row], [Rep]:[Region]]", ("X", 4, 9), "'S'!$A4:$B4"),
    ("Sales[#This Row]", ("X", 2, 9), "'S'!$A2:$D2"),
    ("Sales[@Amount]", ("S", 5, 4), "$C5"),
    ("[@Amount]", ("S", 5, 4), "$C5"),  # written inside its own table
    ("Sales[@[Unit Price '[USD']]]", ("S", 2, 1), "$D2"),  # ' escapes a bracket in a column name
    ("[Amount]", ("S", 2, 1), "$C$2:$C$5"),
    ("Sales[[#This Row],[Amount]]", ("X", 8, 1), "#VALUE!"),  # outside the data rows, as in Excel
])
def test_structured_references_become_a1(text, at, want):
    assert to_a1(text, T, at) == want


@pytest.mark.parametrize("text, at, reason", [
    ("Sales[Nope]", ("X", 9, 9), "no column 'Nope'"),
    ("Other[Amount]", ("X", 9, 9), "unknown table 'Other'"),
    ("[Amount]", ("X", 9, 9), "outside any table"),
    ("Sales[[Rep],[Amount]]", ("X", 9, 9), "several column items"),
])
def test_structured_references_that_cannot_be_rewritten_say_why(text, at, reason):
    with pytest.raises(Unsupported, match=reason.replace("[", r"\[")):
        to_a1(text, T, at)


def test_a_workbook_with_a_table_verifies(tmp_path):
    s = {"A1": "Rep", "B1": "Region", "C1": "Amount", "D1": "Double"}
    for r, (rep, region, amount) in enumerate([("Asha", "West", 80), ("Ravi", "East", 40), ("Asha", "East", 60),
                                                ("Vikram", "West", 15)], 2):
        s[f"A{r}"], s[f"B{r}"], s[f"C{r}"] = rep, region, float(amount)
        s[f"D{r}"] = ("=Sales[[#This Row],[Amount]]*2", 2.0 * amount)  # the long form a file stores
    s["A6"], s["C6"] = "Total", ("=SUBTOTAL(9,Sales[Amount])", 195.0)
    s["F1"] = ("=SUM(Sales[Amount])", 195.0)
    s["F2"] = ('=COUNTIF(Sales[Region],"West")', 2.0)
    s["F3"] = ('=SUMIFS(Sales[Amount],Sales[Rep],"Asha")', 140.0)
    s["F4"] = ('=VLOOKUP("Vikram",Sales[[Rep]:[Amount]],3,FALSE)', 15.0)
    s["F5"] = ("=Sales[[#Headers],[Amount]]", "Amount")
    s["F6"] = ("=Sales[[#Totals],[Amount]]", 195.0)
    s["F8"] = ("=Sales[[#This Row],[Amount]]", "#VALUE!")  # row 8 is not a data row
    s["F9"] = ("=SUM(Sales[Nope])", 0.0)
    path = tmp_path / "t.xlsx"
    write_xlsx(path, {"S": s}, tables={"S": [("Sales", "A1:D6", ["Rep", "Region", "Amount", "Double"], 1)]})
    book = load(path)
    compiled = compile_book(book)
    verdicts = {v.key: v for v in verify(book, compiled)}
    assert {k: v.status for k, v in verdicts.items() if v.status != "PASS"} == {("S", 9, 6): "UNRESOLVED"}
    assert "no column 'Nope'" in verdicts[("S", 9, 6)].reason
    assert any("column" in n and "D" in n and "region" in n for n in compiled.notes), compiled.notes  # pasted down
