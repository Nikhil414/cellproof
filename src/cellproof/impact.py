"""Rank the unresolved reasons: which fix, done first, reaches the most cells.

A cell that stops because it reads an unfinished cell carries the reason "depends on unresolved cell X"; the
compiler records X (Unsupported.dep, compiler.chains), so every chained cell can be attributed to the first
unresolved cell its chain reaches - the reason to fix. An attributed cell can hide a further reason of its own,
which surfaces once the blocker is translated: the counts say what a fix lets try again, never that it will PASS.
"""
from __future__ import annotations

from collections import Counter, defaultdict

from cellproof.model import Key, Verdict, a1


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
