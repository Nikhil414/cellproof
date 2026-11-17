"""`cellproof watch`: verify again on every save, print only what changed."""
from __future__ import annotations

import time
from dataclasses import dataclass
from pathlib import Path

from cellproof.model import Key, Verdict, a1
from cellproof.report import readers_text

CAP = 8  # cells named per line before "and N more"


@dataclass
class Run:
    counts: dict[str, int]  # {"PASS": n, "FAIL": n, "UNRESOLVED": n}
    snapshot: dict[Key, Verdict]
    seconds: float
    groups: list[dict]  # unresolved groups (baseline "where to start"), [] when none
    wrong: list[dict]  # wrong answers other cells read, [] when none


def _cell(key: Key) -> str:
    return a1(key)


def _cap(cells: list[str]) -> str:
    return ", ".join(cells[:CAP]) + (f", and {len(cells) - CAP} more" if len(cells) > CAP else "")


def _capped(lines_here: list[str], label: str) -> list[str]:
    """A region flipping verdicts can mean hundreds of cells: name the first few, count the rest."""
    if len(lines_here) <= CAP:
        return lines_here
    return lines_here[:CAP] + [f"and {len(lines_here) - CAP} more {label} cells"]


def _values(x: Verdict) -> str:
    return f"(Excel {one(x.expected)}, SQL {one(x.actual)})"


def one(value) -> str:
    return "(no value)" if value is None else str(value)


def _ascii(text: str) -> str:
    """Console text stays ASCII: a reason or a value can hold any character the workbook held."""
    return text.encode("ascii", "replace").decode("ascii")


def changes(old: dict[Key, Verdict], new: dict[Key, Verdict]) -> list[str]:
    """What changed between two runs, worst first: broke, still-FAIL moves, FAIL<->UNRESOLVED, fixed, new and gone
    cells, and reasons that appear where none were. Empty when the verdicts are the same."""
    lines: list[str] = []
    broke: list[str] = []
    moved: list[str] = []
    flipped: list[str] = []
    fixed: list[str] = []
    added: list[str] = []
    gone: list[str] = []
    for key in sorted(set(old) | set(new)):
        was, now = old.get(key), new.get(key)
        if was is None:
            added.append(_cell(key))
            continue
        if now is None:
            gone.append(_cell(key))
            continue
        if was.status == now.status:
            if now.status == "FAIL" and (was.expected, was.actual) != (now.expected, now.actual):
                moved.append(f"still FAIL: {_cell(key)} (Excel {now.expected}, SQL was {was.actual}, now {now.actual})")
            continue
        if was.status == "PASS":
            broke.append(f"broke: {_cell(key)} PASS -> {now.status} "
                         + (_values(now) if now.status == "FAIL" else f"({now.reason})"))
        elif now.status == "PASS":
            fixed.append(f"fixed: {_cell(key)} {was.status} -> PASS")
        else:  # FAIL <-> UNRESOLVED: the translation did not change the answer, only what can be said about it
            flipped.append(f"changed: {_cell(key)} {was.status} -> {now.status} "
                           + (f"({now.reason})" if now.status == "UNRESOLVED" else _values(now)))
    lines += _capped(broke, "broke")
    lines += _capped(moved, "still FAIL")
    lines += _capped(flipped, "changed")
    lines += _capped(fixed, "fixed")
    if added:
        lines.append(f"{len(added)} new cells: {_cap(added)}")
    if gone:
        lines.append(f"cells gone ({len(gone)}): {_cap(gone)}")
    old_reasons = {x.reason for x in old.values() if x.status == "UNRESOLVED"}
    explained = {key for key in set(old) & set(new) if old[key].status != "UNRESOLVED"}
    now = [(x.reason, key) for key, x in new.items() if x.status == "UNRESOLVED"]
    for reason in sorted({r for r, key in now if key not in explained} - old_reasons):
        n = sum(1 for r, key in now if r == reason and key not in explained)
        lines.append(f"new unresolved reason, {n} cells: {reason}")
    return [_ascii(line) for line in lines]


def _counts_line(counts: dict[str, int]) -> str:
    return f"{counts['PASS']} pass, {counts['FAIL']} fail, {counts['UNRESOLVED']} unresolved"


def run_watch(path: Path, interval: float, runs: int | None, rerun, out=print, sleep=None,
              stat=None, err=print) -> int:
    """Print the first run, then verify again whenever the file's mtime or size changes, printing only what changed.
    `rerun` does the work and returns a Run, or None when the workbook could not be read (Excel may still be writing
    it): the result is retried once without waiting for another save, then the last successful result is kept until
    the next save. A workbook that is unreadable from the start is an error (exit 2), one that goes away mid-watch
    waits for it to come back. `runs` stops the loop after that many re-runs (0: baseline only); without it the loop
    runs until Ctrl+C."""
    stat = stat or path.stat
    sleep = sleep or time.sleep
    state: dict[Key, Verdict] | None = None
    old_counts: dict[str, int] | None = None
    stamp = None  # (mtime_ns, size) of the last verified save
    re_runs = 0
    failed = 0  # consecutive unreadable reads: the first one retries at once, the second waits for a new save
    first = True
    while True:
        if not first:
            sleep(interval)
            try:
                info = stat()
                this = (info.st_mtime_ns, info.st_size)
            except OSError:
                if stamp is not None:
                    out(f"{path.name} is gone; waiting for it to come back")
                    stamp = None
                continue
            if this == stamp:
                continue
        first = False
        stamp = _stat_stamp(stat)  # before the run: a save that lands while it runs is still seen next time
        result = rerun()
        if result is None:
            if state is None:  # never managed to read it: a typo'd path should not wait forever
                err(f"could not read {path}")
                return 2
            out("could not read the workbook (still saving?); keeping the last result")
            failed += 1
            if failed == 1:  # the save may have finished a moment ago: one retry without waiting for the next save
                stamp = None
            continue
        failed = 0
        if state is None:
            out(f"watching {path} ({_counts_line(result.counts)}, {result.seconds:.1f}s; Ctrl+C to stop)")
            for g in result.groups[:5]:
                out(_ascii(f"  {g['cells']:,}  {g['reason']}  ({', '.join(g['examples'])})"))
            for w in result.wrong[:5]:
                out(_ascii(f"  {len(w['readers']):,}  {w['cell']}  Excel {one(w['expected'])}, "
                           f"SQL {one(w['actual'])}  read by {readers_text(w['readers'])}"))
        else:
            lines = changes(state, result.snapshot)
            if lines:
                out(f"changed after save ({result.seconds:.1f}s): {_counts_line(result.counts)} "
                    f"(was {_counts_line(old_counts)})")
                for line in lines:
                    out(f"  {line}")
            else:
                out(f"no change in the verdicts ({_counts_line(result.counts)}, {result.seconds:.1f}s)")
        state = result.snapshot
        old_counts = result.counts
        if runs is not None and re_runs >= runs:
            return 0
        re_runs += 1


def _stat_stamp(stat) -> tuple | None:
    """The stamp to compare against, taken before each run: a save during a run is then still seen."""
    try:
        info = stat()
    except OSError:
        return None
    return (info.st_mtime_ns, info.st_size)
