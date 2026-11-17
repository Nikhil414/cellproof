"""`cellproof watch`: what changed between two runs, and the re-verify loop."""
from __future__ import annotations

import pytest

from pathlib import Path
from types import SimpleNamespace

from cellproof.model import Key, Verdict
from cellproof.watch import Run, changes, run_watch


def v(sheet="Sheet1", row=1, col=1, status="PASS", expected=None, actual=None, reason="") -> Verdict:
    return Verdict((sheet, row, col), status, expected, actual, reason)


def snap(*verdicts) -> dict:
    return {x.key: x for x in verdicts}


def stamp(mtime: int) -> SimpleNamespace:
    """What Path.stat() gives back, just the parts run_watch reads."""
    return SimpleNamespace(st_mtime_ns=mtime, st_size=0)


def test_no_changes_when_identical():
    old = snap(v(), v(row=2, col=2, status="FAIL", expected=1.0, actual=2.0))
    assert changes(old, snap(v(), v(row=2, col=2, status="FAIL", expected=1.0, actual=2.0))) == []


def test_broke_fail_names_the_values():
    old = snap(v())
    new = snap(v(status="FAIL", expected=10.0, actual=11.0))
    assert changes(old, new) == ["broke: Sheet1!A1 PASS -> FAIL (Excel 10.0, SQL 11.0)"]


def test_broke_unresolved_names_the_reason():
    old = snap(v())
    new = snap(v(status="UNRESOLVED", reason="SUMIF not translated"))
    assert changes(old, new) == ["broke: Sheet1!A1 PASS -> UNRESOLVED (SUMIF not translated)"]


def test_fixed_line_keeps_the_old_status():
    old = snap(v(row=2, col=2, status="UNRESOLVED", reason="XLOOKUP not translated"))
    new = snap(v(row=2, col=2))
    assert changes(old, new) == ["fixed: Sheet1!B2 UNRESOLVED -> PASS"]


def test_changed_between_fail_and_unresolved_both_ways():
    old = snap(v(row=2, col=2, status="FAIL", expected=1.0, actual=2.0))
    new = snap(v(row=2, col=2, status="UNRESOLVED", reason="new reason"))
    assert changes(old, new) == ["changed: Sheet1!B2 FAIL -> UNRESOLVED (new reason)"]
    assert changes(new, old) == ["changed: Sheet1!B2 UNRESOLVED -> FAIL (Excel 1.0, SQL 2.0)"]


def test_fail_that_keeps_its_status_only_reports_a_new_answer():
    old = snap(v(row=2, col=2, status="FAIL", expected=1.0, actual=2.0))
    same = changes(old, snap(v(row=2, col=2, status="FAIL", expected=1.0, actual=2.0)))
    assert same == []
    moved = changes(old, snap(v(row=2, col=2, status="FAIL", expected=1.0, actual=3.0)))
    assert moved == ["still FAIL: Sheet1!B2 (Excel 1.0, SQL was 2.0, now 3.0)"]


def test_new_cells_listed():
    old = snap(v(row=2, col=2))
    new = snap(v(row=2, col=2), v(row=3, col=2), v(row=4, col=2))
    assert changes(old, new) == ["2 new cells: Sheet1!B3, Sheet1!B4"]


def test_gone_cell_line():
    old = snap(v(row=2, col=2), v(row=3, col=2))
    new = snap(v(row=2, col=2))
    assert changes(old, new) == ["cells gone (1): Sheet1!B3"]


def test_new_unresolved_reason_counts_cells_now():
    """A reason already on the broke cells is not repeated here; this is for unresolved cells whose reason changed."""
    old = snap(v(), v(row=2, col=2, status="UNRESOLVED", reason="old reason"),
               v(row=3, col=2, status="UNRESOLVED", reason="old reason"))
    new = snap(v(), v(row=2, col=2, status="UNRESOLVED", reason="CONVERT not translated"),
               v(row=3, col=2, status="UNRESOLVED", reason="CONVERT not translated"))
    assert changes(old, new) == ["new unresolved reason, 2 cells: CONVERT not translated"]


def test_lists_capped_with_a_count():
    old = snap(v(), *[v(row=r, col=2) for r in range(2, 15)])
    new = snap(v(), *[v(row=r, col=2, status="FAIL", expected=1.0, actual=2.0) for r in range(2, 15)])
    lines = changes(old, new)
    assert lines[0].startswith("broke: Sheet1!B2")
    assert lines[-1] == "and 5 more broke cells"
    assert len(lines) == 9  # the first 8, then the count


def test_output_is_ascii():
    old = snap(v())
    new = snap(v(), v(row=2, col=2, status="UNRESOLVED", reason="cell reads 'Verkäufe'"))
    assert "\n".join(changes(old, new)).encode("ascii")
def test_run_watch_prints_baseline_then_stops():
    out: list[str] = []
    run_watch(Path("book.xlsx"), interval=0.0, runs=0,
              rerun=lambda: Run({"PASS": 2, "FAIL": 1, "UNRESOLVED": 0}, snap(v(), v(row=2, col=2)), 0.5, [], []),
              out=out.append, sleep=lambda s: None, stat=lambda: stamp(1))
    assert "2 pass, 1 fail, 0 unresolved" in "\n".join(out)
    assert "0.5s" in "\n".join(out)


def test_run_watch_reverifies_after_a_save():
    out: list[str] = []
    before = snap(v(), v(row=2, col=2, status="FAIL", expected=1.0, actual=2.0))
    after = snap(v(), v(row=2, col=2, status="FAIL", expected=1.0, actual=3.0))
    runs = {"n": 0}

    def rerun():
        runs["n"] += 1
        counts = {"PASS": 2, "FAIL": 1, "UNRESOLVED": 0}
        return Run(counts, before if runs["n"] == 1 else after, 0.1, [], [])

    stamps = iter([stamp(1), stamp(1), stamp(2), stamp(2), stamp(2), stamp(2)])
    run_watch(Path("book.xlsx"), interval=0.0, runs=1, rerun=rerun, out=out.append,
              sleep=lambda s: None, stat=lambda: next(stamps))
    text = "\n".join(out)
    assert "changed after save" in text
    assert "SQL was 2.0, now 3.0" in text


def test_run_watch_stays_quiet_when_nothing_changed():
    out: list[str] = []
    stamps = iter([stamp(1), stamp(1), stamp(2), stamp(2), stamp(2)])
    run_watch(Path("book.xlsx"), interval=0.0, runs=1,
              rerun=lambda: Run({"PASS": 1, "FAIL": 0, "UNRESOLVED": 0}, snap(v()), 0.1, [], []),
              out=out.append, sleep=lambda s: None, stat=lambda: next(stamps))
    text = "\n".join(out)
    assert "no change in the verdicts" in text  # the save happened, the answer did not
    assert "changed after save" not in text


def test_run_watch_keeps_the_last_result_when_the_book_cannot_be_read():
    out: list[str] = []
    state = {"n": 0}

    def rerun():
        state["n"] += 1
        if state["n"] == 2:
            return None
        return Run({"PASS": 1, "FAIL": 0, "UNRESOLVED": 0}, snap(v()), 0.1, [], [])

    stamps = iter([stamp(1), stamp(1), stamp(2), stamp(2), stamp(3), stamp(3)])
    run_watch(Path("book.xlsx"), interval=0.0, runs=1, rerun=rerun, out=out.append,
              sleep=lambda s: None, stat=lambda: next(stamps))
    text = "\n".join(out)
    assert "could not read" in text
    assert text.count("could not read") == 1  # the re-run after the next save recovers


def test_run_watch_reports_a_workbook_that_goes_away():
    out: list[str] = []
    state = {"n": 0}
    gone = iter([1, 1, None, 2, 2, 2])

    def stat():
        s = next(gone)
        if s is None:
            raise OSError("gone")
        return stamp(s)

    def rerun():
        state["n"] += 1
        return Run({"PASS": 1, "FAIL": 0, "UNRESOLVED": 0}, snap(v()), 0.1, [], [])

    assert run_watch(Path("book.xlsx"), interval=0.0, runs=1, rerun=rerun, out=out.append,
                     sleep=lambda s: None, stat=stat) == 0
    assert any("gone" in line for line in out)


def test_run_watch_exits_2_when_the_first_read_fails():
    out: list[str] = []
    err: list[str] = []
    assert run_watch(Path("book.xlsx"), interval=0.0, runs=None, rerun=lambda: None,
                     out=out.append, sleep=lambda s: None, stat=lambda: stamp(1),
                     err=err.append) == 2
    assert any("could not read" in line for line in err)  # a typo'd path is an error, on stderr
    assert out == []  # nothing was verified, so nothing was reported
