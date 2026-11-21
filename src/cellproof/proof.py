"""A proof bundle: the SQL that answers the workbook, the report, and a seal - plus `--check`, which trusts none of it.

`cellproof prove book.xlsx` verifies the workbook exactly as `convert` does, then writes a folder holding `cells.sql`
(one runnable DuckDB script, the same chain the verifier ran), `report.json` and `REPORT.md` (what `convert` would
have written), and `proof.json` (the workbook's hash, the tool versions, the hash of each sibling, and one hash over
the set).

The bundle is a proof because a stranger needs only DuckDB: run `cells.sql`, compare each row with `report.json`, and
check the workbook's hash. Nothing in that act asks cellproof to agree with itself.

`cellproof prove --check dir [--workbook book.xlsx]` automates the same act: it re-hashes the files, optionally
re-reads the workbook, and runs `cells.sql` to see whether it still answers what `report.json` recorded. It is
written to do nothing beyond what a person could do by hand, so a disagreement between the two is a bug in `--check`.

The seal is not a signature. It catches a file that changed by accident or casually; it cannot stop a forger with
write access to the folder, who can reseal. What a forger still cannot hide is `cells.sql` disagreeing with
`report.json`, which is why `--check` runs it. `REPORT.md` says so, in one line, rather than letting "proof" promise
more than it can.
"""
from __future__ import annotations

import hashlib
import json
import platform
import sys
from datetime import datetime, timezone
from pathlib import Path

import duckdb
from openpyxl.utils.cell import coordinate_to_tuple

from cellproof import __version__
from cellproof.model import Book, Cell, Key, a1
from cellproof.query import build, selection
from cellproof.reader import UnreadableWorkbook, load
from cellproof.report import write_report
from cellproof.verifier import same

FORMAT = "cellproof-proof/1"
FILES = ("cells.sql", "report.json", "REPORT.md")  # the siblings proof.json seals
NOTE = ("_This bundle is sealed, not signed: the hashes catch an accident or a casual edit, not somebody who can "
        "write to this folder. It needs no trust in cellproof - run `cells.sql` in DuckDB and compare the rows with "
        "`report.json`, or run `cellproof prove --check` on the folder._")

STATUSES = ("PASS", "FAIL", "UNRESOLVED")


# --- the script ----------------------------------------------------------------------------------

def cells_sql(book: Book, compiled) -> str:
    """Every input cell, every batch, one SELECT: the chain the verifier ran, as a script anyone can run.

    The selection is the whole workbook, so the seed carries the input cells the formulas read *and* the input cells
    in the answer, and what the last SELECT returns is every cell that holds something."""
    answer = build(book, compiled, selection(book, None, None))
    return answer.sql.replace("(the reasons are on stderr)", "(the reasons are in report.json)")


# --- answers, compared one way only ---------------------------------------------------------------

def _answer(cell: Cell) -> tuple:
    """(kind, num, txt) - the way a cells-table row holds this saved value, so `verifier.same` can compare it."""
    if cell.kind == "b":
        return "b", 1.0 if cell.value else 0.0, None
    if cell.kind == "n":
        return "n", float(cell.value), None
    if cell.kind in ("s", "e"):
        return cell.kind, None, str(cell.value)
    return "blank", None, None


def _same(answer, value) -> bool:
    """Is this (kind, num, txt) answer the same as this plain value?

    `report.json` records an answer as a plain value, which loses its kind: a string that *is* an error code reads
    the same whether the cell was an error or text. So a string is tried as both, and only there - numbers and
    booleans are compared exactly, and the comparison itself is `verifier.same`, the one notion of "equal" here."""
    kind, num, txt = answer
    if isinstance(value, bool):
        return same(kind, num, txt, "b", value)
    if isinstance(value, (int, float)):
        return same(kind, num, txt, "n", value)
    if isinstance(value, str):
        return same(kind, num, txt, "e", value) or same(kind, num, txt, "s", value)
    return same(kind, num, txt, "blank", value)


def _key(text: str) -> Key:
    """'Model!F42' -> ("Model", 42, 6). A sheet name cannot contain "!", so the last one separates."""
    sheet, ref = text.rsplit("!", 1)
    row, col = coordinate_to_tuple(ref)
    return sheet, row, col


# --- the seal ------------------------------------------------------------------------------------

def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _bundle_hash(files: dict) -> str:
    """One number over the set: the sha256 of the sorted `name:hash` lines."""
    sealed = "".join(f"{name}:{files[name]}\n" for name in sorted(files))
    return hashlib.sha256(sealed.encode("utf-8")).hexdigest()


def _workbook_entry(workbook: Path) -> dict:
    return {"file": workbook.name, "bytes": workbook.stat().st_size, "sha256": _sha256(workbook)}


def manifest(workbook: Path, book: Book, counts, computed: int) -> dict:
    """Everything proof.json records but the seal of the sibling files, which `seal` adds."""
    return {
        "format": FORMAT,
        "workbook": _workbook_entry(workbook),
        "cellproof": __version__,
        "engine": {"name": "duckdb", "version": duckdb.__version__},
        "python": platform.python_version(),
        "platform": platform.platform(),
        "created": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "saved_values": {"calc_mode": book.calc.mode, "full_calc_on_load": book.calc.full_on_load,
                         "calc_id": book.calc.calc_id, "last_saved_by": book.app},
        "counts": {status: counts[status] for status in STATUSES},
        "computed_answers": computed,
        "files": {},
        "bundle": "",
    }


def seal(out: Path, workbook: Path, data: dict | None = None) -> dict:
    """Hash every sibling and the set, and write proof.json. Callable on its own: given no `data` it re-seals the
    manifest already in the folder, which is exactly what someone with write access to the folder can do - which is
    why `--check` also runs the SQL instead of stopping at the hashes."""
    if data is None:
        try:
            data = json.loads((out / "proof.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            data = {"format": FORMAT}
    data = dict(data)
    data.setdefault("format", FORMAT)
    if "workbook" not in data:
        data["workbook"] = _workbook_entry(workbook)
    files = {name: _sha256(out / name) for name in FILES}
    data["files"] = files
    data["bundle"] = _bundle_hash(files)
    (out / "proof.json").write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")
    return data


def note(out: Path) -> None:
    """The one line under REPORT.md's title that says what the seal is and is not."""
    path = out / "REPORT.md"
    head, _, rest = path.read_text(encoding="utf-8").partition("\n")
    path.write_text(f"{head}\n\n{NOTE}\n{rest}", encoding="utf-8")


def bundle(out: Path, workbook: Path, book: Book, compiled, verdicts: list, counts, computed: frozenset[Key]) -> dict:
    """Verify nothing, write everything: the four files and the seal. The caller has already verified; this only
    lays the result down so that someone else can check it."""
    out.mkdir(parents=True, exist_ok=True)
    (out / "cells.sql").write_text(cells_sql(book, compiled), encoding="utf-8")
    write_report(verdicts, compiled.findings, out, workbook.name, computed, compiled.notes, compiled.chains, book)
    note(out)
    return seal(out, workbook, manifest(workbook, book, counts, len(computed)))


# --- --check -------------------------------------------------------------------------------------

def _error(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)


def check(out: Path, workbook: Path | None = None) -> int:
    """Re-check a bundle without trusting the machine that made it: 0 when everything holds, 1 when an answer moved,
    2 when a file, a hash or the format is wrong."""
    manifest_path = out / "proof.json"
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as err:
        _error(f"cannot read {manifest_path}: {err}")
        return 2
    if data.get("format") != FORMAT:
        _error(f"{manifest_path} is not a {FORMAT} bundle (it says {data.get('format')!r})")
        return 2

    sealed = data.get("files") or {}
    for name in FILES:
        path = out / name
        if not path.is_file():
            _error(f"{name} is missing from {out}")
            return 2
        if sealed.get(name) != _sha256(path):
            _error(f"{name} was changed after the bundle was written")
            return 2
    if data.get("bundle") != _bundle_hash({name: sealed[name] for name in FILES}):
        _error(f"the seal of {manifest_path} does not cover the files next to it")
        return 2
    print(f"{len(FILES)} files intact")

    report_path = out / "report.json"
    try:
        report = json.loads(report_path.read_text(encoding="utf-8"))
        recorded = report["cells"]
    except (OSError, ValueError, KeyError) as err:
        _error(f"cannot read the cells of {report_path}: {err}")
        return 2

    code = 0
    if workbook is not None:
        code = _check_workbook(data, recorded, workbook)
        if code == 2:
            return 2

    code = _check_sql(out, recorded) or code
    if code == 0:
        print("the bundle checks out")
    return code


def _check_workbook(data: dict, recorded: list, workbook: Path) -> int:
    """The workbook's hash must be the one in the manifest, and the values it holds must be the ones the report
    records as Excel's - so a bundle cannot pair a report with a workbook that did not produce it."""
    try:
        digest = _sha256(workbook)
    except OSError as err:
        _error(f"cannot read {workbook}: {err}")
        return 2
    if digest != (data.get("workbook") or {}).get("sha256"):
        _error(f"{workbook} is not the workbook this bundle was made from")
        return 2
    print("workbook hash matches")
    try:
        book = load(workbook)
    except UnreadableWorkbook as err:
        _error(str(err))
        return 2
    formulas = {k: c for k, c in book.cells.items() if c.formula is not None}
    want = {_key(cell["cell"]): cell for cell in recorded}
    for key in sorted(set(want) - set(formulas)):  # a formula cell the report knows and the workbook does not
        _error(f"{want[key]['cell']} is a formula cell in the report but not in {workbook.name}")
        return 1
    for key in sorted(set(formulas) - set(want)):  # a formula cell the workbook holds and the report never saw
        _error(f"{a1(key)} is a formula cell in {workbook.name} but not in the report")
        return 1
    matched = 0
    for key, cell in formulas.items():
        if want[key].get("answer") == "computed":
            continue  # the file has no saved value here: the report's expected value is a library's, not Excel's
        if not _same(_answer(cell), want[key]["expected"]):
            _error(f"{a1(key)} holds a different saved value in {workbook.name} than the report records")
            return 1
        matched += 1
    print(f"{matched} saved values match the report")
    return 0


def _check_sql(out: Path, recorded: list) -> int:
    """Run cells.sql on a fresh DuckDB and compare it with what report.json says the answers were."""
    try:
        sql = (out / "cells.sql").read_text(encoding="utf-8")
        rows = duckdb.connect().execute(sql).fetchall()
    except (OSError, duckdb.Error) as err:
        _error(f"cells.sql did not run: {str(err).splitlines()[0]}")
        return 1
    got = {(sheet, row, col): (kind, num, txt) for sheet, row, col, kind, num, txt in rows}
    matched = total = unresolved = 0
    moved = []
    for cell in recorded:
        if cell.get("status") == "UNRESOLVED":
            unresolved += 1  # a hole in cells.sql, not an answer: report.json says why
            continue
        total += 1
        row = got.get(_key(cell["cell"]))
        if row is not None and _same(row, cell.get("actual")):
            matched += 1
        else:
            moved.append(cell["cell"])
    print(f"cells.sql reproduces {matched} of {total} answers")
    for name in moved:
        _error(f"{name} answers differently in cells.sql than report.json records")
    if unresolved:
        print(f"{unresolved} cell(s) could not be translated: cells.sql does not answer them "
              "(report.json says why)")
    return 1 if moved else 0
