"""REPORT.md for people, report.json for tools."""
from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

from cellproof.impact import impact, roots_of
from cellproof.model import Key, Verdict, a1


def _md(value) -> str:
    return "(no value)" if value is None else str(value).replace("|", "\\|")


def write_report(verdicts: list[Verdict], findings: list[str], out: Path, source: str,
                 computed: frozenset[Key] | set[Key] = frozenset(), notes=(),
                 chains: dict[Key, Key] | None = None) -> Counter:
    counts = Counter({"PASS": 0, "FAIL": 0, "UNRESOLVED": 0})
    counts.update(v.status for v in verdicts)
    findings = list(dict.fromkeys(findings))
    groups = impact(chains, verdicts) if chains is not None else None
    roots = roots_of(chains) if chains else {}

    data = {
        "source": source,
        "counts": dict(counts),
        "findings": findings,
        "regions": list(notes),
        "cells": [{"cell": a1(v.key), "status": v.status, "expected": v.expected, "actual": v.actual,
                   "reason": v.reason, "answer": "computed" if v.key in computed else "excel-saved",
                   **({"blocked_by": a1(roots[v.key]) if v.key in roots else None} if chains is not None else {})}
                  for v in verdicts],
    }
    if groups is not None:
        data["unresolved_groups"] = groups
    (out / "report.json").write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    lines = [f"# cellproof report: {source}", "", "| Status | Cells |", "|---|---|"]
    lines += [f"| {status} | {counts[status]} |" for status in ("PASS", "FAIL", "UNRESOLVED")]
    if computed:
        lines += ["", f"{len(computed)} formula cells had no saved value. Their expected values were computed "
                      "by the `formulas` library, not by Excel, and are marked (computed)."]
    if groups:
        lines += ["", "## Where to start", "",
                  "The unresolved cells grouped by the reason to fix first, ranked by how many cells each reason "
                  "reaches. A cell whose reason is \"depends on unresolved cell\" counts towards the group of the "
                  "cell it reads, so fixing a reason at the top lets the most cells try again; one of those cells "
                  "can hold a further reason of its own, which surfaces next.", "",
                  "| Reason | Cells | Depend on it | First cells |", "|---|---|---|---|"]
        lines += [f"| {_md(g['reason'])} | {g['cells']:,} | {g['dependents']:,} | {', '.join(g['examples'])} |"
                  for g in groups]
    if notes:
        lines += ["", "## Regions", ""] + [f"- {n}" for n in notes]
    if findings:
        lines += ["", "## Findings in the workbook", ""] + [f"- {f}" for f in findings]
    failed = [v for v in verdicts if v.status == "FAIL"]
    if failed:
        lines += ["", "## Failed cells", "", "| Cell | Excel | SQL |", "|---|---|---|"]
        lines += [f"| {a1(v.key)} | {_md(v.expected)}{' (computed)' if v.key in computed else ''} | {_md(v.actual)} |"
                  for v in failed]
    unresolved = [v for v in verdicts if v.status == "UNRESOLVED"]
    if unresolved:
        lines += ["", "## Unresolved cells", "", "| Cell | Reason |", "|---|---|"]
        lines += [f"| {a1(v.key)} | {_md(v.reason)} |" for v in unresolved]
    (out / "REPORT.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return counts
