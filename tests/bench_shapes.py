"""Timing tool for the slow sheet shapes. Not a test: run it by hand.

    python tests/bench_shapes.py SHAPE ROWS [--dbt] [--batch-weight N]

SHAPE is running_formula (SUM over a formula column), running_input (SUM over an input column),
vlookup (a VLOOKUP into a 1000-row table on every row) or flat (A*1.1 on every row).
"""
import argparse
import tempfile
import time
from pathlib import Path

from xlsxkit import write_xlsx

import unexcel.compiler as compiler
from unexcel.emitter import run_dbt, write_project
from unexcel.reader import load
from unexcel.verifier import verify


def make(shape: str, n: int) -> dict:
    d = {}
    if shape == "running_formula":
        total = 0.0
        for r in range(1, n + 1):
            b = (r % 97) * 1.5
            total += b
            d[f"A{r}"] = float(r % 97)
            d[f"B{r}"] = (f"=A{r}*1.5", b)
            d[f"C{r}"] = (f"=SUM($B$1:B{r})", total)
        return {"Data": d}
    if shape == "running_input":
        total = 0.0
        for r in range(1, n + 1):
            total += float(r % 97)
            d[f"A{r}"] = float(r % 97)
            d[f"C{r}"] = (f"=SUM($A$1:A{r})", total)
        return {"Data": d}
    if shape == "vlookup":
        table = {}
        for r in range(1, 1001):
            table[f"A{r}"] = float(r)
            table[f"B{r}"] = float(r * 2)
        for r in range(1, n + 1):
            k = float(r % 1000 + 1)
            d[f"A{r}"] = k
            d[f"B{r}"] = (f"=VLOOKUP(A{r},Tbl!$A$1:$B$1000,2,FALSE)", k * 2)
        return {"Data": d, "Tbl": table}
    if shape == "flat":
        for r in range(1, n + 1):
            d[f"A{r}"] = float(r)
            d[f"B{r}"] = (f"=A{r}*1.1", r * 1.1)
        return {"Data": d}
    raise SystemExit(f"unknown shape {shape!r}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("shape")
    ap.add_argument("rows", type=int)
    ap.add_argument("--dbt", action="store_true", help="also write the dbt project, run dbt build, report its size")
    ap.add_argument("--batch-weight", type=int, help="override compiler.BATCH_WEIGHT (once it exists)")
    args = ap.parse_args()
    if args.batch_weight:
        compiler.BATCH_WEIGHT = args.batch_weight
    tmp = Path(tempfile.mkdtemp())
    wb = tmp / "bench.xlsx"
    write_xlsx(wb, make(args.shape, args.rows))
    t0 = time.perf_counter()
    book = load(wb)
    t1 = time.perf_counter()
    compiled = compiler.compile_book(book)
    t2 = time.perf_counter()
    verdicts = verify(book, compiled)
    t3 = time.perf_counter()
    passed = sum(v.status == "PASS" for v in verdicts)
    line = (f"{args.shape} n={args.rows} read={t1 - t0:.1f} compile={t2 - t1:.1f} verify={t3 - t2:.1f} "
            f"batches={len(compiled.batches)} pass={passed}/{len(verdicts)}")
    if args.dbt:
        out = tmp / "out"
        t4 = time.perf_counter()
        write_project(book, compiled, out, "bench")
        ok, _ = run_dbt(out)
        db = out / "unexcel.duckdb"
        line += f" dbt_ok={ok} dbt={time.perf_counter() - t4:.1f}s db_mb={db.stat().st_size / 1e6:.1f}"
    print(line, flush=True)


if __name__ == "__main__":
    main()
