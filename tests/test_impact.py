"""Impact-ranked unresolveds: which reason, fixed first, reaches the most cells.

A cell that stops because it reads an unfinished cell carries 'depends on unresolved cell X'; the compiler records
X, so every chained cell is attributed to the first blocker its chain reaches and the report can rank the reasons.
"""
import json

from xlsxkit import book_of, write_xlsx

from cellproof.cli import main
from cellproof.compiler import compile_book
from cellproof.impact import impact, readers_of, roots_of, split_failures, wrong
from cellproof.model import Verdict
from cellproof.reader import load
from cellproof.report import cache_warning, write_report
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


def test_readers_of_finds_single_cells_ranges_and_whole_columns():
    book = book_of({"S": {
        "A1": ("=1+1", 2),
        "A2": 1, "A3": 2,
        "B1": ("=A1*3", 6),
        "C1": ("=SUM(A2:A3)", 3),
        "D1": ("=SUM(A:A)", 3),
    }})
    assert readers_of(book, [("S", 1, 1)]) == {("S", 1, 1): [("S", 1, 2), ("S", 1, 4)]}
    assert readers_of(book, [("S", 2, 1)]) == {("S", 2, 1): [("S", 1, 3), ("S", 1, 4)]}
    assert readers_of(book, []) == {}


def test_readers_of_crosses_sheets_and_names():
    book = book_of({"T": {"B2": ("=1+1", 2)}, "S": {"A1": ("=T!B2", 2), "A2": ("=Rate", 2)}},
                   names={"Rate": ("T", "B2")})
    assert readers_of(book, [("T", 2, 2)]) == {("T", 2, 2): [("S", 1, 1), ("S", 2, 1)]}


def test_readers_of_skips_a_formula_that_does_not_parse():
    book = book_of({"S": {"A1": ("=1", 1), "B1": ("=A1+", 1)}})
    assert readers_of(book, [("S", 1, 1)]) == {}


def test_wrong_lists_only_failed_cells_with_readers():
    book = book_of({"S": {"A1": ("=1+1", 3), "B1": ("=A1*3", 6), "C1": ("=1+1", 3)}})
    verdicts = verify(book, compile_book(book))
    assert wrong(book, verdicts) == [
        {"cell": "S!A1", "expected": 3.0, "actual": 2.0, "readers": ["S!B1"]},
    ]


def test_report_lists_wrong_answers_other_cells_read(tmp_path):
    book = book_of({"S": {"A1": ("=1+1", 3), "B1": ("=A1*3", 6), "C1": ("=1+1", 3)}})
    verdicts = verify(book, compile_book(book))
    write_report(verdicts, [], tmp_path, "b.xlsx", book=book)
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "## Wrong answers other cells read" in md
    assert "| S!A1 | 3.0 | 2.0 | S!B1 |" in md
    data = json.loads((tmp_path / "report.json").read_text(encoding="utf-8"))
    assert data["wrong_answers"] == [
        {"cell": "S!A1", "expected": 3.0, "actual": 2.0, "readers": ["S!B1"]},
    ]


def test_report_leaves_wrong_answers_out_without_a_book_or_readers(tmp_path):
    verdicts = [Verdict(("S", 1, 1), "FAIL", 3.0, 2.0)]
    write_report(verdicts, [], tmp_path, "b.xlsx")
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "Wrong answers other cells read" not in md
    data = json.loads((tmp_path / "report.json").read_text(encoding="utf-8"))
    assert "wrong_answers" not in data

    book = book_of({"S": {"A1": ("=1+1", 3)}})  # FAIL, but no other cell reads it
    write_report(verify(book, compile_book(book)), [], tmp_path, "b.xlsx", book=book)
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "Wrong answers other cells read" not in md
    data = json.loads((tmp_path / "report.json").read_text(encoding="utf-8"))
    assert "wrong_answers" not in data


# --- the FAIL split: primary differences vs ones that only read a difference -----------------------

def test_split_failures_separates_the_start_from_what_only_reads_it():
    book = book_of({"S": {"A1": ("=1+1", 3), "B1": ("=A1*3", 7), "C1": ("=B1+1", 8), "D1": ("=C1+1", 9)}})
    verdicts = verify(book, compile_book(book))
    assert {v.status for v in verdicts} == {"FAIL"}  # every one of the four disagrees with Excel
    split = split_failures(book, verdicts)
    assert split["primary"] == [{"cell": "S!A1", "expected": 3.0, "actual": 2.0, "dependents": 3}]
    assert split["inherited"] == 3
    assert split["computed_answers_excluded"] == 0


def test_split_failures_keeps_two_unconnected_primaries_apart():
    book = book_of({"S": {"A1": ("=1+1", 3), "B1": ("=A1*3", 7), "C1": ("=1+1", 5),
                          "D1": ("=C1*3", 30), "E1": ("=D1+1", 31)}})
    verdicts = verify(book, compile_book(book))
    assert {v.status for v in verdicts} == {"FAIL"}
    split = split_failures(book, verdicts)
    assert [p["cell"] for p in split["primary"]] == ["S!C1", "S!A1"]  # more dependents first
    assert [p["dependents"] for p in split["primary"]] == [2, 1]
    assert split["inherited"] == 3


def test_a_fail_that_many_cells_read_is_still_primary():
    book = book_of({"S": {"A1": ("=1/0", "#N/A"), **{f"B{r}": ("=A1+1", "#N/A") for r in range(1, 6)}}})
    verdicts = verify(book, compile_book(book))
    split = split_failures(book, verdicts)
    assert split["primary"] == [{"cell": "S!A1", "expected": "#N/A", "actual": "#DIV/0!", "dependents": 5}]
    assert split["inherited"] == 5


def test_a_cycle_is_unresolved_so_no_fail_loop_can_hide_a_primary():
    # `split_failures` assumes every inherited FAIL is downstream of a primary; that holds only because a cycle is
    # UNRESOLVED rather than FAIL. Pinned here so a future change to cycle handling cannot quietly break the split.
    book = book_of({"S": {"A1": ("=B1+1", 5), "B1": ("=A1+1", 6)}})
    verdicts = verify(book, compile_book(book))
    assert [v.status for v in verdicts] == ["UNRESOLVED", "UNRESOLVED"]
    assert split_failures(book, verdicts) == {
        "primary": [], "inherited": 0, "computed_answers_excluded": 0}


def test_split_failures_leaves_out_computed_answers():
    book = book_of({"S": {"A1": ("=1+1", 3)}})
    verdicts = [Verdict(("S", 1, 1), "FAIL", 3.0, 2.0)]
    assert split_failures(book, verdicts, {("S", 1, 1)}) == {
        "primary": [], "inherited": 0, "computed_answers_excluded": 1}


def test_split_failures_is_empty_without_fails():
    book = book_of({"S": {"A1": 2.0, "B1": ("=A1*2", 4)}})
    assert split_failures(book, verify(book, compile_book(book))) == {
        "primary": [], "inherited": 0, "computed_answers_excluded": 0}


def test_report_splits_failures_and_names_the_primary(tmp_path):
    book = book_of({"S": {"A1": ("=1+1", 3), "B1": ("=A1*3", 7), "C1": ("=B1+1", 8)}})
    verdicts = verify(book, compile_book(book))
    write_report(verdicts, [], tmp_path, "b.xlsx", book=book)
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "## Failures to look at" in md
    assert "| S!A1 | 3.0 | 2.0 | 2 |" in md
    assert "2 more failed cells read one of these" in md
    assert "## Failed cells" not in md  # the flat table is replaced, not repeated
    data = json.loads((tmp_path / "report.json").read_text(encoding="utf-8"))
    assert data["failures"]["primary"] == [
        {"cell": "S!A1", "expected": 3.0, "actual": 2.0, "dependents": 2}]
    assert data["failures"]["inherited"] == 2


def test_report_keeps_the_flat_table_when_there_is_no_book_to_split_with(tmp_path):
    write_report([Verdict(("S", 2, 1), "FAIL", 5.0, 4.0)], [], tmp_path, "b.xlsx")
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "## Failed cells" in md and "| S!A2 | 5.0 | 4.0 |" in md
    assert "Failures to look at" not in md


def test_cache_warning_reads_only_what_the_file_says(tmp_path):
    clean = tmp_path / "clean.xlsx"
    write_xlsx(clean, {"S": {"A1": 1.0}}, calc={"calcId": 191029}, app="Microsoft Excel")
    assert cache_warning(load(clean)) == ""

    manual = tmp_path / "manual.xlsx"
    write_xlsx(manual, {"S": {"A1": 1.0}}, calc={"calcMode": "manual", "calcId": 191029}, app="Microsoft Excel")
    said = cache_warning(load(manual))
    assert 'calcMode="manual"' in said and "Recalculate (F9)" in said and "may be out of date" in said

    pending = tmp_path / "pending.xlsx"
    write_xlsx(pending, {"S": {"A1": 1.0}}, calc={"calcId": 191029, "fullCalcOnLoad": True}, app="Microsoft Excel")
    assert "fullCalcOnLoad" in cache_warning(load(pending))

    unstamped = tmp_path / "unstamped.xlsx"
    write_xlsx(unstamped, {"S": {"A1": 1.0}}, app="Microsoft Excel")
    assert "calcId" in cache_warning(load(unstamped))

    other = tmp_path / "other.xlsx"
    write_xlsx(other, {"S": {"A1": 1.0}}, calc={"calcId": 191029}, app="LibreOffice")
    assert "LibreOffice" in cache_warning(load(other))


def test_warnings_with_no_fails_print_no_failure_table_and_no_stale_sentence(tmp_path):
    wb = tmp_path / "m.xlsx"
    write_xlsx(wb, {"S": {"A1": 2.0, "B1": ("=A1*2", 4.0)}}, calc={"calcMode": "manual"}, app="Microsoft Excel")
    book = load(wb)
    write_report(verify(book, compile_book(book)), [], tmp_path, "m.xlsx", book=book)
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert "Failed cells" not in md and "Failures to look at" not in md and "out of date" not in md
    assert "failures" not in json.loads((tmp_path / "report.json").read_text(encoding="utf-8"))


def test_the_warning_sits_above_the_failure_table(tmp_path):
    wb = tmp_path / "m.xlsx"
    write_xlsx(wb, {"S": {"A1": ("=1+1", 3.0), "B1": ("=A1*3", 7.0)}},
               calc={"calcMode": "manual", "calcId": 191029}, app="Microsoft Excel")
    book = load(wb)
    write_report(verify(book, compile_book(book)), [], tmp_path, "m.xlsx", book=book)
    md = (tmp_path / "REPORT.md").read_text(encoding="utf-8")
    assert md.index("out of date") < md.index("## Failures to look at") < md.index("| S!A1 | 3.0 | 2.0 | 1 |")
