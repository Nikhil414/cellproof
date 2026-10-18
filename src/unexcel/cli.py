"""unexcel convert WORKBOOK [--out DIR] [--dbt]"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from unexcel.compiler import NO_VALUE, compile_book
from unexcel.compute import fill_missing
from unexcel.emitter import run_dbt, write_project
from unexcel.reader import UnreadableWorkbook, load
from unexcel.report import write_report
from unexcel.verifier import SqlFailed, verify


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")  # Windows consoles can't print every character
    parser = argparse.ArgumentParser(prog="unexcel", description="Turn an Excel workbook into a verified dbt project.")
    sub = parser.add_subparsers(dest="command", required=True)
    convert = sub.add_parser("convert", help="convert a workbook and verify every formula cell")
    convert.add_argument("workbook", type=Path)
    convert.add_argument("--out", type=Path, help="output folder (default: <workbook>_dbt next to the file)")
    convert.add_argument("--dbt", action="store_true", help="also run `dbt build` on the result")
    args = parser.parse_args(argv)

    try:
        book = load(args.workbook)
    except UnreadableWorkbook as err:
        print(f"error: {err}", file=sys.stderr)
        return 2
    book, computed, why = fill_missing(book, args.workbook)
    compiled = compile_book(book)
    if why:
        compiled.unresolved.update({k: f"{r} ({why})" for k, r in compiled.unresolved.items() if r == NO_VALUE})
    try:
        verdicts = verify(book, compiled)
    except SqlFailed as err:
        print(f"error: {err}", file=sys.stderr)
        return 3
    out = args.out or args.workbook.with_name(args.workbook.stem + "_dbt")
    write_project(book, compiled, out, args.workbook.stem)
    counts = write_report(verdicts, compiled.findings, out, args.workbook.name, computed, compiled.notes)
    if computed:
        print(f"{len(computed)} expected values computed by the formulas library (the file had none saved)")
    print(f"{counts['PASS']} pass, {counts['FAIL']} fail, {counts['UNRESOLVED']} unresolved -> {out / 'REPORT.md'}")
    code = 1 if counts["FAIL"] else 0
    if args.dbt:
        ok, log = run_dbt(out)
        print("dbt build: ok" if ok else f"dbt build failed:\n{log}")
        code = code if ok else 1
    return code
