"""cellproof convert WORKBOOK [--out DIR] [--dbt]"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from cellproof.compiler import NO_VALUE, compile_book
from cellproof.compute import fill_missing
from cellproof.emitter import run_dbt, write_project
from cellproof.reader import UnreadableWorkbook, load
from cellproof.report import write_report
from cellproof.verifier import SqlFailed, verify_book


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")  # Windows consoles can't print every character
    parser = argparse.ArgumentParser(prog="cellproof", description="Turn an Excel workbook into a verified dbt project.")
    sub = parser.add_subparsers(dest="command", required=True)
    convert = sub.add_parser("convert", help="convert a workbook and verify every formula cell")
    convert.add_argument("workbook", type=Path)
    convert.add_argument("--out", type=Path, help="output folder (default: <workbook>_dbt next to the file)")
    convert.add_argument("--dbt", action="store_true", help="also run `dbt build` on the result")
    convert.add_argument("--target", choices=["duckdb", "postgres"], default="duckdb",
                         help="verify on DuckDB (default) or also on Postgres (needs: pip install cellproof[postgres])")
    convert.add_argument("--pg-url", help="the Postgres to verify on (default: a Postgres of cellproof's own)")
    args = parser.parse_args(argv)

    try:
        book = load(args.workbook)
    except UnreadableWorkbook as err:
        print(f"error: {err}", file=sys.stderr)
        return 2
    book, computed, why = fill_missing(book, args.workbook)
    def compile_fn(b):
        compiled = compile_book(b)
        if why:
            compiled.unresolved.update({k: f"{r} ({why})" for k, r in compiled.unresolved.items() if r == NO_VALUE})
        return compiled
    try:
        compiled, verdicts, failures = verify_book(book, compile_fn)
    except SqlFailed as err:
        print(f"error: {err}", file=sys.stderr)
        return 3
    for failure in failures:  # those cells are UNRESOLVED in the report; the run goes on
        print(f"error: {failure}", file=sys.stderr)
    out = args.out or args.workbook.with_name(args.workbook.stem + "_dbt")
    dbt = None  # (ok, log) once dbt build ran
    if args.target == "postgres":
        try:
            from cellproof import postgres
            from psycopg import OperationalError
        except ImportError:
            print("error: the Postgres target needs: pip install cellproof[postgres]", file=sys.stderr)
            return 2
        try:
            with postgres.connect(args.pg_url) as con:
                duck = sum(v.status == "PASS" for v in verdicts)
                compiled, verdicts = postgres.verify_book(book, compile_fn, con)
                compiled.findings.append(f"verified on Postgres: these verdicts are Postgres's ({duck} cells PASS on DuckDB)")
                write_project(book, compiled, out, args.workbook.stem, target="postgres")
                if args.dbt:  # while this Postgres runs: cellproof's own one stops when the block ends
                    dbt = run_dbt(out, env=postgres.dbt_env(con))
        except OperationalError as err:
            print(f"error: could not connect to Postgres: {str(err).strip().splitlines()[0]}", file=sys.stderr)
            return 2
    else:
        write_project(book, compiled, out, args.workbook.stem)
        if args.dbt:
            dbt = run_dbt(out)
    counts = write_report(verdicts, compiled.findings, out, args.workbook.name, computed, compiled.notes,
                          compiled.chains)
    if computed:
        print(f"{len(computed)} expected values computed by the formulas library (the file had none saved)")
    print(f"{counts['PASS']} pass, {counts['FAIL']} fail, {counts['UNRESOLVED']} unresolved -> {out / 'REPORT.md'}")
    code = 3 if failures else 1 if counts["FAIL"] else 0
    if dbt is not None:
        ok, log = dbt
        print("dbt build: ok" if ok else f"dbt build failed:\n{log}")
        code = code if ok else 1
    return code
