"""`cellproof diff old.xlsx new.xlsx`: what changed between two versions, and which of it broke an answer.

The verdicts come from the same region-mode run `convert` makes, because that is the run a person acts on. The SQL
comparison uses a second compile in cell mode (`compile_book(book, regions=False)`): there a formula's SQL depends on
its own text and its inputs, not on which neighbouring columns happened to group into a region, so a region
reshuffle does not read as hundreds of statements changing. Both compiles reuse the loaded book, so the cost is
compile time only.

Two cells are the same cell when they share a sheet name and an A1 address. A renamed sheet therefore reads as
everything removed and everything added: the truth, said rather than guessed around.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from openpyxl.utils import get_column_letter

from cellproof.compiler import compile_book
from cellproof.model import Book, Key, Verdict, a1
from cellproof.report import _md, cache_warning

# How a cell is written into the canonical SQL, so two cells can be compared by *which* cell they read.
# \x00 cannot occur in a sheet name, so a tag can never be confused with SQL text.
_TAG_IN_TEXT = re.compile(r"\x00([^\x00]*)\x00(\d+)\x00(\d+)\x00")
_TAG = r"\x00[^\x00]*\x00\d+\x00\d+\x00"
_CELL_READ = re.compile(r"FROM __cellproof_cells__ WHERE sheet = '((?:[^']|'')*)' "
                        r"AND row_num = (\d+) AND col_num = (\d+)")
_OPEN_READ = re.compile(rf"\(SELECT (v|t|k|e) FROM ({_TAG})\)")
_ALIAS = {"v": "num", "t": "txt", "k": "kind", "e": "err"}
_TXT_READ = re.compile(rf"\(SELECT COALESCE\(txt, ''\) FROM ({_TAG})\)")
_ERR_READ = re.compile(rf"\(SELECT CASE WHEN kind = 'e' THEN txt END FROM ({_TAG})\)")
_BOOL_READ = re.compile(rf"\(\(SELECT num FROM ({_TAG})\) <> 0\)")


def _tag(key: Key) -> str:
    return f"\x00{key[0]}\x00{key[1]}\x00{key[2]}\x00"


def _show(sql: str) -> str:
    """The tags back as A1 cells, so a stored difference is readable: `(SELECT num FROM <Sheet!B2>)`."""
    return _TAG_IN_TEXT.sub(lambda m: f"<{m.group(1)}!{get_column_letter(int(m.group(3)))}{int(m.group(2))}>", sql)


def _canonical(step, names: dict[str, str]) -> str:
    """One formula cell's SQL with every way of naming another cell reduced to the same words.

    Reads differ with the batch layout, not with what the formula means: a cell finished in an earlier batch is read
    from the table (`SELECT num ... WHERE row_num = 3`), one still open in the same batch from its CTE
    (`SELECT v FROM c7`), and a boolean keeps a `<> 0` wrapper on one path only. Normalising all of them to
    `(SELECT num FROM <Sheet!A3>)` makes a moving batch boundary, or a renumbered CTE, invisible; a cell that really
    reads something else keeps its difference. The error expression is part of the SQL for the same reason."""
    sql = step.expr.sql + " ||ERR|| " + step.expr.err
    for cte, tag in names.items():  # this batch's names -> the cells they hold
        sql = re.sub(rf"\bFROM {cte}\b", "FROM " + tag, sql)
    sql = _CELL_READ.sub(lambda m: "FROM " + _tag((m.group(1).replace("''", "'"), int(m.group(2)), int(m.group(3)))),
                         sql)
    sql = _OPEN_READ.sub(lambda m: f"(SELECT {_ALIAS[m.group(1)]} FROM {m.group(2)})", sql)
    sql = _TXT_READ.sub(r"(SELECT txt FROM \1)", sql)  # the seed reads '' back as NULL; not a difference
    sql = _ERR_READ.sub(r"(SELECT err FROM \1)", sql)
    sql = _BOOL_READ.sub(r"(SELECT num FROM \1)", sql)  # (num <> 0) is how a boolean is spelled off the table
    return _show(sql)


def _sqls(book: Book) -> dict[Key, str]:
    """Every formula cell's canonical SQL, from a cell-mode compile (region layout must not enter the comparison)."""
    compiled = compile_book(book, regions=False)
    out: dict[Key, str] = {}
    for batch in compiled.batches:
        names = {step.cte: _tag(step.key) for step in batch}
        for step in batch:
            out[step.key] = _canonical(step, names)
    return out


def _cells(book: Book) -> dict[Key, tuple[str | None, str, object]]:
    """Each cell's (formula text, kind, saved value) before any compile adds its hidden array-constant sheets."""
    sheets = set(book.sheets)
    return {k: (c.formula, c.kind, c.value) for k, c in book.cells.items() if k[0] in sheets}


def _text(formula: str | None) -> str | None:
    return None if formula is None else "=" + formula


def _changed(was: tuple, now: tuple) -> bool:
    """The saved value differs: a different kind (2 vs "2"), or a different value of the same kind."""
    return (was[1], was[2]) != (now[1], now[2])


_RANK = {"FAIL": 0, "UNRESOLVED": 1, "PASS": 2}


def _verdict_changes(old: dict[Key, Verdict], new: dict[Key, Verdict]) -> list[dict]:
    """The cells whose verdict differs. A new FAIL comes first: it is the headline.

    A cell that stopped being a formula, or that is gone, is already named under `text` or `removed`; a cell that
    became a formula and passes is already named under `text`. Both are left out here so the list stays news."""
    out = []
    for key in sorted(set(old) | set(new)):
        now = new.get(key)
        if now is None:
            continue
        was = old.get(key)
        if was is None and now.status == "PASS":
            continue
        if was is not None and was.status == now.status:
            continue
        out.append({"cell": a1(key), "was": was.status if was is not None else None, "now": now.status,
                    "expected": now.expected, "actual": now.actual, "reason": now.reason})
    return sorted(out, key=lambda v: (_RANK[v["now"]], v["cell"]))


def new_fails(data: dict) -> list[dict]:
    """The verdict changes that are a vendict getting worse: a FAIL where there was no FAIL before."""
    return [v for v in data["verdicts"] if v["now"] == "FAIL" and v["was"] != "FAIL"]


def _summary(counts: dict, fails: int, regions: int) -> str:
    if not any(counts.values()) and not regions:
        return "the two workbooks answer the same way"
    parts = [f"{fails} new FAIL{'' if fails == 1 else 's'}"] if fails else []
    for key, one, many in (("added", "added", "added"), ("removed", "removed", "removed"),
                           ("inputs", "input value changed", "input values changed"),
                           ("text", "formula text changed", "formula texts changed"),
                           ("sql", "formula reads different cells", "formulas read different cells"),
                           ("verdicts", "verdict changed", "verdicts changed")):
        if counts[key]:
            parts.append(f"{counts[key]} {one if counts[key] == 1 else many}")
    if regions:
        parts.append(f"{regions} region change{'' if regions == 1 else 's'}")
    return ", ".join(parts)


def compare(old: Book, new: Book, old_verdicts: dict[Key, Verdict] | None = None,
            new_verdicts: dict[Key, Verdict] | None = None, old_notes=(), new_notes=()) -> dict:
    """What changed between the two books; pass the verdicts the way `convert` computes them, or there is nothing to
    compare but the cells and their SQL."""
    before, after = _cells(old), _cells(new)  # before either compile adds the hidden array-constant sheets
    old_sql, new_sql = _sqls(old), _sqls(new)
    added = sorted(set(after) - set(before))
    removed = sorted(set(before) - set(after))
    inputs: list[dict] = []
    text: list[dict] = []
    sql: list[dict] = []
    for key in sorted(set(before) & set(after)):
        was_formula, _, _ = before[key]
        now_formula, _, _ = after[key]
        if was_formula != now_formula:
            text.append({"cell": a1(key), "was": _text(was_formula), "now": _text(now_formula)})
        elif was_formula is not None:
            if old_sql.get(key) != new_sql.get(key):  # same text, but what it reads moved
                sql.append({"cell": a1(key), "was": old_sql.get(key), "now": new_sql.get(key)})
        elif _changed(before[key], after[key]):
            inputs.append({"cell": a1(key), "was": before[key][2], "now": after[key][2]})
    changes = _verdict_changes(old_verdicts or {}, new_verdicts or {})
    stopped = sorted(set(old_notes) - set(new_notes))
    started = sorted(set(new_notes) - set(old_notes))
    counts = {"added": len(added), "removed": len(removed), "inputs": len(inputs), "text": len(text),
              "sql": len(sql), "verdicts": len(changes)}
    warning = cache_warning(new)
    return {
        "added": [a1(k) for k in added],
        "removed": [a1(k) for k in removed],
        "inputs": inputs,
        "text": text,
        "sql": sql,
        "verdicts": changes,
        "counts": counts,
        "regions": {"stopped": stopped, "started": started},
        "warning": warning,
        "summary": _summary(counts, len(new_fails({"verdicts": changes})), len(stopped) + len(started)),
    }


def _row(values) -> str:
    return "| " + " | ".join(_md(v) for v in values) + " |"


def _table(head: tuple[str, ...], rows: list[tuple], cap: int = 100) -> list[str]:
    lines = ["| " + " | ".join(head) + " |", "|" + "---|" * len(head)]
    lines += [_row(r) for r in rows[:cap]]
    if len(rows) > cap:
        lines += ["", f"{len(rows) - cap:,} more rows are in diff.json."]
    return lines


def _worth_a_page(data: dict) -> bool:
    return bool(data["counts"]["verdicts"]) or sum(data["counts"].values()) >= 10


def write_diff(data: dict, out: Path, old_name: str, new_name: str) -> Path | None:
    """diff.json always; DIFF.md when the difference is more than a summary. Returns the page written, or None."""
    out.mkdir(parents=True, exist_ok=True)
    (out / "diff.json").write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    if not _worth_a_page(data):
        return None
    counts = data["counts"]
    fails = new_fails(data)
    lines = [f"# cellproof diff: {old_name} -> {new_name}", ""]
    the_default = "no cell that passed before fails now: the differences below did not break an answer"
    if fails:
        named = ", ".join(f"{v['cell']} (Excel {_md(v['expected'])}, SQL {_md(v['actual'])})" for v in fails[:5])
        lines.append(f"{len(fails)} cell{'' if len(fails) == 1 else 's'} passed before and fail{'' if len(fails) == 1 else ''} "
                     f"now: {named}" + (f", and {len(fails) - 5} more" if len(fails) > 5 else "") +
                     ". Fixing one of them may fix the others, because a formula that reads a wrong answer is "
                     "usually wrong too.")
    else:
        lines.append(the_default[0].upper() + the_default[1:] + ".")
    lines.append("")
    lines += [f"- {label}: {n}" for n, label in ((counts["added"], "cells added"),
                                                 (counts["removed"], "cells removed"),
                                                 (counts["inputs"], "saved input values changed"),
                                                 (counts["text"], "formula texts changed"),
                                                 (counts["sql"], "formulas reading different cells"),
                                                 (counts["verdicts"], "verdicts changed")) if n]
    if data["warning"]:
        lines += ["", data["warning"]]
    if data["verdicts"]:
        lines += ["", "## Verdicts that changed", "",
                  "PASS, FAIL or UNRESOLVED per cell, from the same run `convert` makes. A FAIL says Excel's saved "
                  "value and cellproof's SQL disagree; it never says which of them is right.", ""]
        lines += _table(("Cell", "Was", "Now", "Excel", "SQL", "Reason"),
                        [(v["cell"], v["was"] or "(not a formula)", v["now"], v["expected"], v["actual"],
                          v["reason"]) for v in data["verdicts"]])
    if data["sql"]:
        lines += ["", "## Same formula, different cells read", "",
                  "The formula text is unchanged, but what it reads moved: a named range grew or shrank, a table "
                  "gained rows, a range's corner moved. Without this section the two versions look alike and answer "
                  "differently.", ""]
        lines += _table(("Cell", "Old SQL", "New SQL"), [(s["cell"], s["was"], s["now"]) for s in data["sql"]])
    if data["text"]:
        lines += ["", "## Formula text that changed", ""]
        lines += _table(("Cell", "Was", "Now"), [(t["cell"], t["was"] or "(was an input)", t["now"] or "(is now an "
                                                  "input)") for t in data["text"]])
    if data["inputs"]:
        lines += ["", "## Saved input values that changed", "",
                  "These are values the workbook held, not values cellproof computed: a FAIL elsewhere can come from "
                  "one of them.", ""]
        lines += _table(("Cell", "Was (Excel)", "Now (Excel)"), [(i["cell"], i["was"], i["now"])
                                                                 for i in data["inputs"]])
    if data["added"] or data["removed"]:
        lines += ["", "## Cells added and removed", ""]
        for title, cells in (("Added", data["added"]), ("Removed", data["removed"])):
            if cells:
                shown = ", ".join(cells[:50]) + (f", and {len(cells) - 50:,} more" if len(cells) > 50 else "")
                lines += [f"- **{title} ({len(cells):,})**: {shown}"]
        lines += ["", "Cells are matched by sheet name and A1 address only, so a renamed sheet reads as everything "
                      "removed and everything added."]
    regions = data["regions"]
    if regions["stopped"] or regions["started"]:
        lines += ["", "## Regions", "",
                  "The columns that verified as one block in one version and not the other (the SQL comparison "
                  "ignores this: it uses cell mode on both)."]
        lines += [f"- now: {n}" for n in regions["started"]] + [f"- before: {n}" for n in regions["stopped"]]
    path = out / "DIFF.md"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path
