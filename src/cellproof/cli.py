"""cellproof convert WORKBOOK [--out DIR] [--dbt] | cellproof sql WORKBOOK [--sheet S] [--range A1:A9] [--out FILE]
| cellproof watch WORKBOOK [--interval SECONDS]"""
from __future__ import annotations

import argparse
import sys
import time
from collections import Counter
from pathlib import Path

from cellproof import fastread, proof, reader
from cellproof.compiler import NO_VALUE, compile_book
from cellproof.compute import fill_missing
from cellproof.emitter import run_dbt, write_project
from cellproof.impact import impact, split_failures, wrong
from cellproof.query import QueryError, build, check, selection
from cellproof.reader import SUPPORTED, UnreadableWorkbook, load
from cellproof.report import readers_text, write_report
from cellproof.verifier import SqlFailed, verify_book
from cellproof.watch import Run, run_watch


def _load(args):
    """The workbook, read with the fast loader when --fast, or with openpyxl when that loader cannot.

    The fast loader reads each worksheet once instead of twice, and raises UnreadableWorkbook rather than guess
    when it meets something it does not handle; openpyxl then reads the file, so --fast can only be slower."""
    path = Path(args.workbook)
    if not getattr(args, "fast", False) or path.suffix.lower() not in SUPPORTED:
        return load(path)
    try:
        return fastread.load(path)
    except UnreadableWorkbook as err:
        print(f"note: openpyxl is reading {path.name} instead ({err})", file=sys.stderr)
    return load(path)


def _value(value) -> str:
    return "(no value)" if value is None else str(value)


def _summarise(groups: list[dict], wrong_answers: list[dict], file=None) -> None:
    """What to fix first, on the console: the ranked unresolved reasons and the wrong answers other cells read."""
    file = file or sys.stdout
    if groups:
        print("Where to start (unresolved, ranked by how many cells each blocks):", file=file)
        for g in groups[:5]:
            print(f"  {g['cells']:,}  {g['reason']}  ({', '.join(g['examples'])})", file=file)
    if wrong_answers:
        print("Wrong answers other cells read (direct readers only):", file=file)
        for w in wrong_answers[:5]:
            print(f"  {len(w['readers']):,}  {w['cell']}  Excel {_value(w['expected'])}, SQL {_value(w['actual'])}"
                  f"  read by {readers_text(w['readers'])}", file=file)


def _sql(args) -> int:
    """`cellproof sql`: compile the whole workbook, print only what the selection needs. SQL to stdout, the reasons
    an asked-for cell is missing to stderr, so `cellproof sql book.xlsx > run.sql` yields a file you can run."""
    try:
        book = _load(args)
    except UnreadableWorkbook as err:
        print(f"error: {err}", file=sys.stderr)
        return 2
    try:  # the selection is checked before the workbook is compiled: a typo'd range should not cost a compile
        check(book, args.sheet, args.range)
    except QueryError as err:
        print(f"error: {err}", file=sys.stderr)
        return 2
    book, computed, why = fill_missing(book, args.workbook)
    compiled = compile_book(book)
    if why:
        compiled.unresolved.update({k: f"{r} ({why})" for k, r in compiled.unresolved.items() if r == NO_VALUE})
    try:  # again, now that the compiler has added the hidden array-constant sheets (selection excludes them)
        target = selection(book, args.sheet, args.range)
    except QueryError as err:  # unreachable: check() passed on the same arguments
        print(f"error: {err}", file=sys.stderr)
        return 2
    answer = build(book, compiled, target)
    if computed:
        print(f"{len(computed)} expected values computed by the formulas library (the file had none saved)",
              file=sys.stderr)
    if answer.unresolved:
        print(f"{len(answer.unresolved):,} of the {len(answer.selection):,} cells asked for could not be translated "
              f"and are not in this SQL (where to start):", file=sys.stderr)
        _summarise(answer.groups, [], file=sys.stderr)
    if args.stats:
        total = len(compiled.steps)
        print(f"the workbook compiles {total:,} formula cells in {len(compiled.batches):,} statement(s); "
              f"this SQL keeps {len(answer.computed):,} of them in {answer.batches:,}, "
              f"reading {answer.inputs:,} input cells", file=sys.stderr)
    if args.out is None:
        _write_sql(answer.sql)
    else:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(answer.sql, encoding="utf-8")
        print(f"wrote {args.out} ({answer.batches:,} statement(s))", file=sys.stderr)
    return 1 if answer.unresolved else 0


def _watch(args) -> int:
    """`cellproof watch`: verify the workbook, then verify again on every save, printing only what changed."""
    path = args.workbook

    def rerun():
        started = time.monotonic()
        try:
            book = _load(args)
            book, computed, why = fill_missing(book, path)

            def compile_fn(b):
                compiled = compile_book(b)
                if why:
                    compiled.unresolved.update({k: f"{r} ({why})" for k, r in compiled.unresolved.items()
                                                if r == NO_VALUE})
                return compiled

            compiled, verdicts, failures = verify_book(book, compile_fn)
            for failure in failures:  # those cells are UNRESOLVED; the run goes on
                print(f"error: {failure}", file=sys.stderr)
            counts = Counter({"PASS": 0, "FAIL": 0, "UNRESOLVED": 0})
            counts.update(v.status for v in verdicts)
            return Run(counts, {v.key: v for v in verdicts}, time.monotonic() - started,
                       impact(compiled.chains, verdicts), wrong(book, verdicts))
        except Exception as err:  # an unreadable workbook is normal here (Excel is still saving); keep watching
            first = str(err).strip().splitlines()[0] if str(err).strip() else repr(err)
            print(f"error: {first}", file=sys.stderr)
            return None

    return run_watch(path, args.interval, args.runs, rerun, err=lambda line: print(line, file=sys.stderr))


def _prove(args) -> int:
    """`cellproof prove`: verify the workbook as `convert` does, then write a bundle someone else can check."""
    if args.check is not None:
        workbook = args.check_workbook
        return proof.check(Path(args.check), workbook)
    if args.workbook is None:
        print("error: prove needs a workbook (or --check DIR to re-check a bundle)", file=sys.stderr)
        return 2
    try:
        book = _load(args)
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
    for failure in failures:  # those cells are UNRESOLVED in the report; the bundle is still worth writing
        print(f"error: {failure}", file=sys.stderr)
    out = args.out or args.workbook.with_name(args.workbook.stem + "_proof")
    counts = Counter({"PASS": 0, "FAIL": 0, "UNRESOLVED": 0})
    counts.update(v.status for v in verdicts)
    proof.bundle(out, args.workbook, book, compiled, verdicts, counts, computed)
    if computed:
        print(f"{len(computed)} expected values computed by the formulas library (the file had none saved)")
    print(f"{counts['PASS']} pass, {counts['FAIL']} fail, {counts['UNRESOLVED']} unresolved -> {out / 'REPORT.md'}")
    print(f"the bundle is in {out} (check it with: cellproof prove --check {out})")
    return 3 if failures else 1 if counts["FAIL"] else 0


def _write_sql(text: str) -> None:
    """The SQL as UTF-8 bytes: sheet names and text values can hold any character, and a console encoding that
    cannot print them (Windows' cp1252, or an ASCII redirect) must replace them - a replaced sheet name would make
    the script answer the wrong cells, which is worse than mojibake on the way through."""
    buffer = getattr(sys.stdout, "buffer", None)
    if buffer is None:
        sys.stdout.write(text)  # a captured stream without bytes (rare): let it decide
        return
    sys.stdout.flush()
    buffer.write(text.encode("utf-8"))
    buffer.flush()


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")  # Windows consoles can't print every character
    parser = argparse.ArgumentParser(prog="cellproof", description="Turn an Excel workbook into a verified dbt project.")
    sub = parser.add_subparsers(dest="command", required=True)
    convert = sub.add_parser("convert", help="convert a workbook and verify every formula cell")
    convert.add_argument("workbook", type=Path)
    convert.add_argument("--out", type=Path, help="output folder (default: <workbook>_dbt next to the file)")
    convert.add_argument("--fast", action="store_true",
                         help="read with the single-pass loader (falls back to openpyxl if it cannot)")
    convert.add_argument("--dbt", action="store_true", help="also run `dbt build` on the result")
    convert.add_argument("--target", choices=["duckdb", "postgres"], default="duckdb",
                         help="verify on DuckDB (default) or also on Postgres (needs: pip install cellproof[postgres])")
    convert.add_argument("--pg-url", help="the Postgres to verify on (default: a Postgres of cellproof's own)")
    convert.add_argument("--stats", action="store_true",
                         help="print a tally of the unresolved cells' reasons, most common first")
    sql = sub.add_parser("sql", help="print the SQL for a selection of cells, and nothing it does not need",
                         description="Print a runnable DuckDB script that answers the cells you name: the input cells "
                                     "they read, the formulas in between, and a final SELECT. Exit 1 when an "
                                     "asked-for cell could not be translated (the SQL then lacks it).")
    sql.add_argument("workbook", type=Path)
    sql.add_argument("--fast", action="store_true",
                     help="read with the single-pass loader (falls back to openpyxl if it cannot)")
    sql.add_argument("--sheet", help="the sheet to answer for (default: the whole workbook)")
    sql.add_argument("--range", dest="range",
                     help="an A1 rectangle on --sheet, like D2:D1000 (default: all of it)")
    sql.add_argument("--out", type=Path, help="write the SQL to this file instead of stdout")
    sql.add_argument("--stats", action="store_true", help="print how much SQL was needed to stderr")
    watch = sub.add_parser("watch", help="verify again on every save; print only what changed",
                           description="Verify the workbook, then check the file and verify again on every save, "
                                       "printing only what changed between runs: cells that broke or were fixed, "
                                       "new FAILs, verdicts that flipped, new unresolved reasons. Runs until Ctrl+C.")
    watch.add_argument("workbook", type=Path)
    watch.add_argument("--fast", action="store_true",
                       help="read with the single-pass loader (falls back to openpyxl if it cannot)")
    watch.add_argument("--interval", type=float, default=1.0, help="seconds between checks (default: 1.0)")
    watch.add_argument("--runs", type=int, help="stop after this many re-runs (default: until Ctrl+C)")
    prove = sub.add_parser("prove", help="write a bundle someone else can check, or re-check one",
                           description="Verify the workbook as `convert` does, then write a folder holding cells.sql "
                                       "(one runnable DuckDB script answering every cell), report.json, REPORT.md, "
                                       "and proof.json (the workbook's hash, the tool versions and a seal). "
                                       "`prove --check DIR` re-checks a bundle without trusting the machine that "
                                       "made it. The SQL is about 10 MB for 200k formula cells; no compression, "
                                       "because a proof should be readable.")
    prove.add_argument("workbook", type=Path, nargs="?")
    prove.add_argument("--out", type=Path, help="the bundle folder (default: <workbook>_proof next to the file)")
    prove.add_argument("--fast", action="store_true",
                       help="read with the single-pass loader (falls back to openpyxl if it cannot)")
    prove.add_argument("--check", metavar="DIR", help="re-check the bundle in DIR instead of writing one")
    prove.add_argument("--workbook", dest="check_workbook", type=Path,
                       help="with --check: also re-read the workbook this bundle names")
    args = parser.parse_args(argv)
    if args.command == "sql":
        return _sql(args)
    if args.command == "prove":
        return _prove(args)
    if args.command == "watch":
        try:
            return _watch(args)
        except KeyboardInterrupt:  # Ctrl+C is how a watcher ends: not an error
            print("", file=sys.stderr)
            return 0

    try:
        book = _load(args)
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
    groups = impact(compiled.chains, verdicts)
    wrong_answers = wrong(book, verdicts)
    split = split_failures(book, verdicts, computed)
    counts = write_report(verdicts, compiled.findings, out, args.workbook.name, computed, compiled.notes,
                          compiled.chains, book, groups, wrong_answers, split)
    if computed:
        print(f"{len(computed)} expected values computed by the formulas library (the file had none saved)")
    print(f"{counts['PASS']} pass, {counts['FAIL']} fail, {counts['UNRESOLVED']} unresolved -> {out / 'REPORT.md'}")
    _summarise(groups, wrong_answers)
    if args.stats:
        if counts["FAIL"]:
            print(f"Failures: {len(split['primary']):,} to look at, {split['inherited']:,} that read one of them"
                  + (f", {split['computed_answers_excluded']:,} excluded (expected values from the formulas library)"
                     if split["computed_answers_excluded"] else ""))
        reasons = Counter(v.reason for v in verdicts if v.status == "UNRESOLVED")
        if reasons:
            print("Unresolved by reason (top 15):")
            for reason, n in reasons.most_common(15):
                print(f"  {n}  {reason}")
        else:
            print("no unresolved cells")
    code = 3 if failures else 1 if counts["FAIL"] else 0
    if dbt is not None:
        ok, log = dbt
        print("dbt build: ok" if ok else f"dbt build failed:\n{log}")
        code = code if ok else 1
    return code
