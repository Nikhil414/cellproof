"""REPORT.md for people, report.json for tools."""
from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

from cellproof.impact import impact, roots_of, split_failures, wrong
from cellproof.model import Book, Key, Verdict, a1, other_app


def _md(value) -> str:
    return "(no value)" if value is None else str(value).replace("|", "\\|")


def cache_warning(book: Book | None) -> str:
    """The file's own evidence that its saved values may be out of date, as one sentence, or "".

    Only what the workbook says about itself: a manual calculation mode, a full recalculation pending on load, no
    engine stamp, or a last-saver that is not Excel. Nothing is inferred from the values - calling a cell stale is
    exactly the guess this project must not make."""
    if book is None:
        return ""
    calc = book.calc
    why = []
    if calc.mode == "manual":
        why.append("it is set not to recalculate (calcMode=\"manual\"), so its saved values are whatever the last "
                   "full calculation left")
    if calc.full_on_load:
        why.append("it will recalculate everything when Excel next opens it (fullCalcOnLoad=\"1\")")
    if not calc.calc_id:
        why.append("it carries no calculation-engine stamp (calcId), so Excel probably did not write it last")
    app = other_app(book)
    if app:
        why.append(f"it was last saved by {app}, not Microsoft Excel, so its saved values follow that program's "
                   "rules")
    if not why:
        return ""
    said = "; ".join(why)
    return ("> **The file says its saved values may be out of date.** " + said[0].upper() + said[1:] + ". Recalculate "
            "(F9) and save in Excel, then run again. Until then a FAIL is not evidence that the SQL is wrong.")


def readers_text(names: list[str]) -> str:
    """The first five readers, then how many more: a whole column can read one wrong cell."""
    return ", ".join(names[:5]) + (f", and {len(names) - 5} more" if len(names) > 5 else "")


def write_report(verdicts: list[Verdict], findings: list[str], out: Path, source: str,
                 computed: frozenset[Key] | set[Key] = frozenset(), notes=(),
                 chains: dict[Key, Key] | None = None, book: Book | None = None,
                 groups: list[dict] | None = None, wrong_answers: list[dict] | None = None,
                 failures: dict | None = None) -> Counter:
    counts = Counter({"PASS": 0, "FAIL": 0, "UNRESOLVED": 0})
    counts.update(v.status for v in verdicts)
    findings = list(dict.fromkeys(findings))
    if groups is None and chains is not None:
        groups = impact(chains, verdicts)
    if wrong_answers is None and book is not None:
        wrong_answers = wrong(book, verdicts)
    if failures is None and book is not None:
        failures = split_failures(book, verdicts, computed)
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
    if wrong_answers:
        data["wrong_answers"] = wrong_answers
    if failures is not None and counts["FAIL"]:
        data["failures"] = failures
    (out / "report.json").write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    lines = [f"# cellproof report: {source}", "", "| Status | Cells |", "|---|---|"]
    lines += [f"| {status} | {counts[status]} |" for status in ("PASS", "FAIL", "UNRESOLVED")]
    if computed:
        lines += ["", f"{len(computed)} formula cells had no saved value. Their expected values were computed "
                      "by the `formulas` library, not by Excel, and are marked (computed)."]
    warning = cache_warning(book) if counts["FAIL"] else ""  # only beside FAILs: a clean file is left alone
    if warning:
        lines += ["", warning]
    if groups:
        lines += ["", "## Where to start", "",
                  "The unresolved cells grouped by the reason to fix first, ranked by how many cells each reason "
                  "reaches. A cell whose reason is \"depends on unresolved cell\" counts towards the group of the "
                  "cell it reads, so fixing a reason at the top lets the most cells try again; one of those cells "
                  "can hold a further reason of its own, which surfaces next.", "",
                  "| Reason | Cells | Depend on it | First cells |", "|---|---|---|---|"]
        lines += [f"| {_md(g['reason'])} | {g['cells']:,} | {g['dependents']:,} | {', '.join(g['examples'])} |"
                  for g in groups]
    if wrong_answers:
        lines += ["", "## Wrong answers other cells read", "",
                  "Excel has one value and cellproof's SQL has another, and other formulas read the wrong one, so the "
                  "difference can travel. These are direct readers only: a reader may add its own difference.", "",
                  "| Cell | Excel | SQL | Read by |", "|---|---|---|---|"]
        lines += [f"| {w['cell']} | {_md(w['expected'])} | {_md(w['actual'])} | {readers_text(w['readers'])} |"
                  for w in wrong_answers]
    if notes:
        lines += ["", "## Regions", ""] + [f"- {n}" for n in notes]
    shown = list(findings)
    if warning and other_app(book):  # the sentence above already says it: not twice on one page
        shown = [f for f in shown if not f.startswith(f"the file was last saved by {other_app(book)}")]
    if shown:
        lines += ["", "## Findings in the workbook", ""] + [f"- {f}" for f in shown]
    failed = [v for v in verdicts if v.status == "FAIL"]
    if failed:
        primary = failures["primary"] if failures is not None else []
        if primary:
            lines += ["", "## Failures to look at", "",
                      "Excel saved one value and cellproof's SQL has another. None of these cells reads a cell that "
                      "also failed, so the difference starts here: either the translation is wrong or the saved value "
                      "never matched its formula. \"Read by\" counts the failed cells whose difference may have "
                      "travelled from this one.", "",
                      "| Cell | Excel | SQL | Read by |", "|---|---|---|---|"]
            lines += [f"| {p['cell']} | {_md(p['expected'])} | {_md(p['actual'])} | {p['dependents']:,} |"
                      for p in primary]
            if failures["inherited"]:
                n = failures["inherited"]
                lines += ["", f"{n:,} more failed cell{'s' if n != 1 else ''} read{'s' if n == 1 else ''} one of "
                              "these: they may be right themselves, so fix the cells above first."]
        else:  # no book to split with (or no primary found): keep the flat list rather than lose the cells
            lines += ["", "## Failed cells", "", "| Cell | Excel | SQL |", "|---|---|---|"]
            lines += [f"| {a1(v.key)} | {_md(v.expected)}{' (computed)' if v.key in computed else ''} | "
                      f"{_md(v.actual)} |" for v in failed]
    unresolved = [v for v in verdicts if v.status == "UNRESOLVED"]
    if unresolved:
        lines += ["", "## Unresolved cells", "", "| Cell | Reason |", "|---|---|"]
        lines += [f"| {a1(v.key)} | {_md(v.reason)} |" for v in unresolved]
    (out / "REPORT.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return counts
