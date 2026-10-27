"""Impact-ranked unresolveds: which reason, fixed first, reaches the most cells.

A cell that stops because it reads an unfinished cell carries 'depends on unresolved cell X'; the compiler records
X, so every chained cell is attributed to the first blocker its chain reaches and the report can rank the reasons.
"""
import json

from xlsxkit import book_of, write_xlsx

from cellproof.cli import main
from cellproof.compiler import compile_book
from cellproof.impact import impact, roots_of
from cellproof.model import Verdict
from cellproof.report import write_report
from cellproof.verifier import verify


def test_roots_of_walks_to_the_first_blocker():
    chains = {("S", 3, 1): ("S", 2, 1), ("S", 2, 1): ("S", 1, 1)}
    assert roots_of(chains) == {("S", 3, 1): ("S", 1, 1), ("S", 2, 1): ("S", 1, 1)}
    assert roots_of({}) == {}


def test_roots_of_survives_a_loop():  # chains should be a forest; a loop must stop, not hang
    k1, k2 = ("S", 1, 1), ("S", 2, 1)
    assert set(roots_of({k1: k2, k2: k1}).values()) == {k2}


def test_compile_records_the_blocker():
    book = book_of({"S": {"A1": ("=GAMMA(1)", 1), "A2": ("=A1*2", 2), "A3": ("=A2+1", 3)}})
    compiled = compile_book(book)
    assert compiled.unresolved[("S", 1, 1)] == "function GAMMA not supported yet"
    assert compiled.chains == {("S", 2, 1): ("S", 1, 1), ("S", 3, 1): ("S", 2, 1)}


def test_range_read_records_its_first_unfinished_cell():
    book = book_of({"S": {"A1": 1, "A2": ("=GAMMA(1)", 1), "A3": 3, "B1": ("=SUM(A1:A3)", 5)}})
    compiled = compile_book(book)
    assert compiled.chains == {("S", 1, 2): ("S", 2, 1)}


def test_chain_into_a_cell_with_no_saved_value():
    book = book_of({"S": {"A1": ("=1", None), "A2": ("=A1+1", 1)}})
    compiled = compile_book(book)
    assert compiled.unresolved[("S", 1, 1)].startswith("no saved value")
    assert compiled.chains == {("S", 2, 1): ("S", 1, 1)}


def test_impact_groups_by_root_reason_and_ranks_by_reach():
    book = book_of({"S": {"A1": ("=GAMMA(1)", 1), "A2": ("=A1+1", 2), "A3": ("=A2+1", 3),
                          "B1": ("=NOW()", 1), "C1": ("=B1+1", 2)}})
    compiled = compile_book(book)
    groups = impact(compiled.chains, verify(book, compiled))
    assert [g["reason"] for g in groups] == ["function GAMMA not supported yet", "volatile function NOW"]
    assert groups[0] == {"reason": "function GAMMA not supported yet", "cells": 3, "dependents": 2,
                         "examples": ["S!A1"]}
    assert groups[1] == {"reason": "volatile function NOW", "cells": 2, "dependents": 1, "examples": ["S!B1"]}


def test_a_cell_reading_two_blockers_is_counted_once():
    book = book_of({"S": {"A1": ("=GAMMA(1)", 1), "B1": ("=NOW()", 1), "C1": ("=A1+B1", 2)}})
    compiled = compile_book(book)
    groups = impact(compiled.chains, verify(book, compiled))
    assert sum(g["cells"] for g in groups) == 3  # C1 sits in one group only: the blocker the compile hit first
    assert {g["cells"] for g in groups} == {1, 2}


def test_report_ranks_the_reasons(tmp_path):
    verdicts = [
        Verdict(("S", 1, 1), "PASS", 1.0, 1.0),
        Verdict(("S", 2, 1), "UNRESOLVED", 1.0, None, "function GAMMA not supported yet"),
        Verdict(("S", 3, 1), "UNRESOLVED", 2.0, None, "depends on unresolved cell S!A2"),
        Verdict(("S", 4, 1), "UNRESOLVED", 3.0, None, "depends on unresolved cell S!A3"),
        Verdict(("S", 5, 1), "UNRESOLVED", 1.0, None, "volatile function NOW"),
    ]
    chains = {("S", 3, 1): ("S", 2, 1), ("S", 4, 1): ("S", 3, 1)}
    write_report(verdicts, [], tmp_path, "b.xlsx", chains=chains)
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "## Where to start" in md
    assert "| function GAMMA not supported yet | 3 | 2 | S!A2 |" in md
    assert "| volatile function NOW | 1 | 0 | S!A5 |" in md
    assert md.index("GAMMA") < md.index("NOW")
    data = json.loads((tmp_path / "report.json").read_text(encoding="utf-8"))
    assert data["unresolved_groups"][0]["reason"] == "function GAMMA not supported yet"
    cells = {c["cell"]: c for c in data["cells"]}
    assert cells["S!A4"]["blocked_by"] == "S!A2"
    assert cells["S!A2"]["blocked_by"] is None


def test_report_counts_use_thousands_separators(tmp_path):
    chains = {}
    verdicts = [Verdict(("S", 1, 1), "UNRESOLVED", 1.0, None, "function FOO not supported yet")]
    for r in range(2, 1237):
        chains[("S", r, 1)] = ("S", r - 1, 1)
        verdicts.append(Verdict(("S", r, 1), "UNRESOLVED", 1.0, None, f"depends on unresolved cell S!A{r - 1}"))
    write_report(verdicts, [], tmp_path, "b.xlsx", chains=chains)
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "| function FOO not supported yet | 1,236 | 1,235 | S!A1 |" in md


def test_report_without_chains_keeps_the_old_shape(tmp_path):
    write_report([Verdict(("S", 1, 1), "UNRESOLVED", 1.0, None, "volatile function NOW")], [], tmp_path, "b.xlsx")
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "Where to start" not in md
    data = json.loads((tmp_path / "report.json").read_text(encoding="utf-8"))
    assert "unresolved_groups" not in data
    assert all("blocked_by" not in c for c in data["cells"])


def test_convert_end_to_end_shows_where_to_start(tmp_path):
    wb = tmp_path / "b.xlsx"
    write_xlsx(wb, {"S": {"A1": ("=GAMMA(1)", 1), "A2": ("=A1+1", 2), "A3": ("=A2+1", 3)}})
    assert main(["convert", str(wb), "--out", str(tmp_path / "out")]) == 0
    md = (tmp_path / "out" / "REPORT.md").read_text(encoding="utf-8")
    assert "| function GAMMA not supported yet | 3 | 2 | S!A1 |" in md
