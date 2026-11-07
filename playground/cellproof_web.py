"""cellproof in the browser (Pyodide). Converts a workbook exactly as `cellproof convert` does and hands the page the
report and the dbt project as a zip. Nothing leaves the browser: no network call is made with the workbook.

The newest DuckDB built for Pyodide is 1.2, which only knows the arrow lambda (`x -> ...`); cellproof writes DuckDB
2.0's `lambda x: ...`, which 1.2 cannot parse. Both mean the same, so on DuckDB < 1.3 every statement's lambdas are
rewritten on the way in."""
from __future__ import annotations

import contextlib
import io
import json
import re
import shutil
import zipfile
from pathlib import Path

import duckdb

from cellproof import cli, region, verifier

# `lambda x, i:` -> `(x, i) ->` (same precedence: the body runs to the end of the argument either way)
# ponytail: textual rewrite; a workbook string literal containing "lambda x:" would be rewritten too (a FAIL, never
# a false PASS). Drop the whole shim once DuckDB publishes a Pyodide build >= 1.3.
_LAMBDA = re.compile(r"\blambda\s+([A-Za-z_]\w*(?:\s*,\s*[A-Za-z_]\w*)*)\s*:")


def arrow(sql: str) -> str:
    return _LAMBDA.sub(lambda m: f"({m.group(1)}) ->", sql)


class _Con:
    def __init__(self, con):
        self._con = con

    def execute(self, sql, *args, **kw):
        self._con.execute(arrow(sql), *args, **kw)
        return self

    def __getattr__(self, name):
        return getattr(self._con, name)


class _DuckDB:
    def __getattr__(self, name):
        return getattr(duckdb, name)

    @staticmethod
    def connect(*args, **kw):
        return _Con(duckdb.connect(*args, **kw))


if tuple(int(p) for p in duckdb.__version__.split(".")[:2]) < (1, 3):
    verifier.duckdb = region.duckdb = _DuckDB()


def run(data: bytes, name: str) -> tuple[str, bytes]:
    """Convert one workbook. -> (JSON of {code, log, report, report_md}, zip of the dbt project and the report)."""
    work = Path("/tmp/cellproof")
    shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True)
    book = work / Path(name).name
    book.write_bytes(data)
    out = work / (book.stem + "_dbt")
    log = io.StringIO()
    with contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
        try:
            code = cli.main(["convert", str(book), "--out", str(out)])
        except Exception as err:  # noqa: BLE001 - shown on the page, never swallowed
            print(f"error: {type(err).__name__}: {err}")
            code = 3
    result = {"code": code, "log": log.getvalue()}
    report = out / "report.json"
    if report.exists():
        result["report"] = json.loads(report.read_text(encoding="utf-8"))
        result["report_md"] = (out / "REPORT.md").read_text(encoding="utf-8")
    buf = io.BytesIO()
    if out.exists():
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for f in sorted(out.rglob("*")):
                if f.is_file():
                    z.write(f, f.relative_to(work))
    return json.dumps(result, ensure_ascii=False), buf.getvalue()
