"""Rank what to fix first: the unresolved reasons, and the wrong answers other cells read.

A cell that stops because it reads an unfinished cell carries the reason "depends on unresolved cell X"; the
compiler records X (Unsupported.dep, compiler.chains), so every chained cell can be attributed to the first
unresolved cell its chain reaches - the reason to fix. An attributed cell can hide a further reason of its own,
which surfaces once the blocker is translated: the counts say what a fix lets try again, never that it will PASS.

A FAIL needs the same treatment: the report lists every failed cell, but a wrong answer that feeds other formulas is
a different kind of problem from one that stands alone. `wrong` names the FAIL cells that other formulas read, one
hop only ("read by", never "blocks"): a reader can itself be wrong, so the count is a neighbourhood, not a
guarantee.

`split_failures` goes further and uses the same read graph the other way: a FAIL whose own inputs were all right is
*primary*, one that reads a FAIL is *inherited*. Both are still FAILs - no cell is ever called stale, because we
cannot see what Excel calculated before the edit - but the split says which ones are worth reading first.
"""
from __future__ import annotations

from bisect import bisect_left, bisect_right
from collections import Counter, defaultdict, deque

from cellproof.model import Book, Key, Unsupported, Verdict, a1
from cellproof.r1c1 import lexed
from cellproof.refs import resolve


def roots_of(chains: dict[Key, Key]) -> dict[Key, Key]:
    """Every chained cell -> the first unresolved cell its chain reaches.

    Memoized: a chain can be one cell deep per row down a whole column (50k rows), so re-walking every start
    would be quadratic."""
    out: dict[Key, Key] = {}
    for start in chains:
        if start in out:
            continue
        path, seen, key = [], {start}, start
        while key not in out:
            nxt = chains.get(key)
            if nxt is None or nxt in seen:  # the root, or a loop (chains should be a forest): stop here
                break
            path.append(key)
            key = nxt
            seen.add(key)
        root = out.get(key, key)
        for k in path:
            out[k] = root
        out[start] = root
    return out


def impact(chains: dict[Key, Key], verdicts: list[Verdict]) -> list[dict]:
    """The unresolved reasons ranked for the report: the reason, the cells attributed to it, how many of them
    sit below a root (its reach), and up to three root cells to look at. Every unresolved cell counts towards
    exactly one reason."""
    reasons = {v.key: v.reason for v in verdicts if v.status == "UNRESOLVED"}
    roots = roots_of(chains)
    cells: Counter = Counter()
    below: Counter = Counter()
    examples: dict[str, list[str]] = defaultdict(list)
    for key in sorted(reasons):
        root = roots.get(key, key)
        reason = reasons.get(root) or reasons[key]
        cells[reason] += 1
        if root != key:
            below[reason] += 1
        elif len(examples[reason]) < 3:
            examples[reason].append(a1(key))
    ranked = sorted(cells, key=lambda r: (-cells[r], r))
    return [{"reason": r, "cells": cells[r], "dependents": below[r], "examples": examples[r]} for r in ranked]


def readers_of(book: Book, keys) -> dict[Key, list[Key]]:
    """For each key, the formula cells that read it directly (one hop): a cell, a range or a whole column in a formula.

    Built from the lexed formulas (tokens already cached by the compile) and one bisect per column of the wanted
    cells, so a whole-column reference (`A:A`) costs O(log n), not O(rows). Sorted and deduplicated; a formula whose
    shape does not parse is skipped because the compiler already reports its real reason. A cell reading itself is
    left out: a cycle is not a reader anyone can fix by looking at the cell."""
    wanted = set(keys)
    if not wanted:
        return {}
    by_col: dict[tuple[str, int], list[int]] = defaultdict(list)
    for sheet, row, col in wanted:
        by_col[(sheet, col)].append(row)
    for rows in by_col.values():
        rows.sort()
    cols_of: dict[str, list[int]] = defaultdict(list)  # a sheet's wanted columns, sorted: a ref spans a few of them
    for sheet, col in by_col:
        cols_of[sheet].append(col)
    for cols in cols_of.values():
        cols.sort()
    out: dict[Key, set[Key]] = defaultdict(set)
    for key, (shape, texts) in lexed(book).items():
        if shape is None or not shape.parses():
            continue
        for text in texts:
            if text[:6].lower() == "_xlpm.":  # a LET name (see fn/logic.py): its value's references were lexed
                continue
            try:
                area = resolve(text, key[0], book)
            except (Unsupported, RecursionError):
                continue  # the compiler reports the real reason for this cell
            cols = cols_of.get(area.sheet, [])
            for col in cols[bisect_left(cols, area.c1):bisect_right(cols, area.c2)]:
                rows = by_col[(area.sheet, col)]
                for row in rows[bisect_left(rows, area.r1):bisect_right(rows, area.r2)]:
                    target = (area.sheet, row, col)
                    if target != key:
                        out[target].add(key)
    return {k: sorted(v) for k, v in out.items()}


def split_failures(book: Book, verdicts: list[Verdict],
                   computed: frozenset[Key] | set[Key] = frozenset()) -> dict:
    """The FAILs told apart: those whose inputs were all right, and those that read a FAIL.

    A FAIL reading no FAIL, directly or through other formula cells, is *primary*: the disagreement starts there, so
    it is either our translation or the value Excel saved. One that reads a FAIL is *inherited*, attributed to the
    nearest primary it reads, so fixing that primary is the move. Both stay FAILs; nothing is called stale, because
    the file is the only thing we trust and it does not say when it last calculated.

    A cell whose expected value came from the `formulas` library is left out and counted separately: it has no saved
    value to disagree with, so the split would be about the library, not the file.

    One call of `readers_of` with every formula's key does the work: it returns the read graph inverted (a cell to
    the formulas that read it), so a breadth-first walk out from the FAILs finds the inherited ones, and a second
    walk out from the primaries says which primary each inherited FAIL reaches first."""
    failed = [v for v in verdicts if v.status == "FAIL"]
    out: dict = {"primary": [], "inherited": 0, "computed_answers_excluded": 0}
    counted = [v for v in failed if v.key not in computed]
    out["computed_answers_excluded"] = len(failed) - len(counted)
    if not counted:
        return out
    by_key = {v.key: v for v in counted}
    fails = set(by_key)
    if len(fails) == 1:  # one FAIL reads no other: no graph worth building
        out["primary"] = [{"cell": a1(next(iter(fails))), "expected": counted[0].expected,
                           "actual": counted[0].actual, "dependents": 0}]
        return out
    readers = readers_of(book, [k for k, c in book.cells.items() if c.formula is not None])

    # Which FAILs a FAIL reads, at any depth: breadth-first over the read graph out from every FAIL. The walk
    # crosses cells that passed, because a difference can travel through a cell whose own saved value happens to
    # match, but the cells it crosses are not FAILs and are not counted.
    inherited: set[Key] = set()
    reached: set[Key] = set()
    queue = deque(sorted(fails))
    while queue:
        for reader in readers.get(queue.popleft(), ()):
            if reader in reached:
                continue
            reached.add(reader)
            if reader in fails:
                inherited.add(reader)
            queue.append(reader)
    primary = [k for k in sorted(fails) if k not in inherited]
    # A primary always exists: following the FAILs upstream through the read graph must end somewhere, because a
    # cycle of FAILs is impossible - a cycle is UNRESOLVED (see tests/test_impact.py). So every inherited FAIL is
    # downstream of a primary, and a breadth-first walk out from the primaries reaches all of them.
    owner: dict[Key, Key] = {k: k for k in primary}  # each cell reached -> the nearest primary upstream
    queue = deque(primary)
    while queue:
        key = queue.popleft()
        for reader in readers.get(key, ()):
            if reader not in owner:
                owner[reader] = owner[key]
                queue.append(reader)
    dependents: Counter = Counter(owner[k] for k in inherited if owner.get(k) != k)
    unreached = inherited - owner.keys()  # cannot happen (a cycle of FAILs is UNRESOLVED), but never lose a cell
    primary = sorted(primary + sorted(unreached))
    rows = [{"cell": a1(k), "expected": by_key[k].expected, "actual": by_key[k].actual,
             "dependents": dependents[k]} for k in primary]
    out["primary"] = sorted(rows, key=lambda r: (-r["dependents"], r["cell"]))
    out["inherited"] = len(inherited) - len(unreached)
    return out


def wrong(book: Book, verdicts: list[Verdict]) -> list[dict]:
    """The FAIL cells other formulas read, ranked by how many formulas read them, one hop only."""
    failed = [v for v in verdicts if v.status == "FAIL"]
    readers = readers_of(book, [v.key for v in failed])
    out = [{"cell": a1(v.key), "expected": v.expected, "actual": v.actual,
            "readers": [a1(k) for k in readers[v.key]]}
           for v in failed if v.key in readers]
    return sorted(out, key=lambda w: (-len(w["readers"]), w["cell"]))
