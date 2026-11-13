"""Write the compiled workbook as a dbt project: on DuckDB, or on Postgres (`target="postgres"`).

The models run the SQL the verifier ran: seed input_cells -> cellproof_cells/cells_0001 -> ... Each batch model
holds only its own rows; a checkpoint model (cp_NNNN) unions everything every 64 batches and after the last
batch, so no model reads more than ~65 sources. workbook_cells keeps the formula cells only.

On Postgres each model is the DuckDB model translated by cellproof.postgres (the SQL the Postgres verification ran),
the helper functions are created in the target schema by an on-run-start macro, and the profile reads its connection
from CELLPROOF_PG_* environment variables.
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
from pathlib import Path

from cellproof.compiler import Compiled, Step, render_batch, write_seed
from cellproof.expr import CELLS
from cellproof.region import WIDE
from cellproof.model import Book

PROJECT = """name: {name}
version: "1.0.0"
config-version: 2
profile: cellproof
flags:
  send_anonymous_usage_stats: false
seeds:
  {name}:
    +column_types:
      sheet: varchar
      row_num: integer
      col_num: integer
      kind: varchar
      num: double
      txt: varchar
models:
  {name}:
    +materialized: table
"""

PROFILES = """cellproof:
  target: dev
  outputs:
    dev:
      type: duckdb
      path: cellproof.duckdb
"""

PROFILES_POSTGRES = """cellproof:
  target: dev
  outputs:
    dev:
      type: postgres
      host: "{{ env_var('CELLPROOF_PG_HOST', 'localhost') }}"
      port: "{{ env_var('CELLPROOF_PG_PORT', '5432') | as_number }}"
      user: "{{ env_var('CELLPROOF_PG_USER', 'postgres') }}"
      password: "{{ env_var('CELLPROOF_PG_PASSWORD', '') }}"
      dbname: "{{ env_var('CELLPROOF_PG_DBNAME', 'postgres') }}"
      schema: "{{ env_var('CELLPROOF_PG_SCHEMA', '__NAME__') }}"
      threads: 1
"""

SCHEMA = """version: 2
models:
  - name: workbook_cells
    description: "Every formula cell of the workbook, recomputed in SQL. One row per cell."
    columns:
      - name: sheet
        data_tests: [not_null]
      - name: row_num
        data_tests: [not_null]
      - name: col_num
        data_tests: [not_null]
"""


def project_name(stem: str) -> str:
    name = re.sub(r"[^A-Za-z0-9]+", "_", stem).strip("_").lower() or "workbook"
    return f"wb_{name}" if name[0].isdigit() else name


def model_name(k: int) -> str:
    return f"cells_{k:04d}"


CHECKPOINT = 64  # a model reads the latest checkpoint plus at most this many batch models


def checkpoint_name(k: int) -> str:
    return f"cp_{k:04d}"


def sources_for(k: int) -> list[str]:
    """Models whose union is every cell known before batch k (1-based): latest checkpoint, then later batches."""
    base = (k - 1) // CHECKPOINT * CHECKPOINT
    first = "input_cells" if base == 0 else checkpoint_name(base)
    return [first] + [model_name(i) for i in range(base + 1, k)]


def _union(models: list[str]) -> str:
    return "\nUNION ALL\n".join("SELECT * FROM {{ ref('" + m + "') }}" for m in models)


def _escape(sql: str) -> str:
    """Braces from workbook text are escaped so Jinja renders them literally."""
    return re.sub(r"[{}]", lambda m: "{{ '" + m.group() + "' }}", sql)


def dbt_sql(steps: list[Step], sources: list[str]) -> str:
    """One batch as a dbt model: the new rows only, reading all earlier cells through the CTE all_cells."""
    body = _escape(render_batch(steps, "all_cells", own_only=True))
    return body.replace("WITH\n", f"WITH\nall_cells AS (\n{_union(sources)}\n),\n", 1)


def dbt_region_wide(batch, sources: list[str]) -> str:
    """The wide, readable model of a region: it reads every earlier cell through the CTE all_cells."""
    body = _escape(batch.wide_sql.replace(CELLS, "all_cells"))
    return body.replace("WITH\n", f"WITH\nall_cells AS (\n{_union(sources)}\n),\n", 1)


def dbt_region_cells(batch) -> str:
    """The long rows of a region: a mechanical unpivot of its wide model."""
    return _escape(batch.unpivot(WIDE)).replace(WIDE, "{{ ref('" + batch.model + "') }}")


def final_sql(last: str) -> str:
    """The formula cells: every row of the last batch that is not an input."""
    return ("SELECT c.* FROM {{ ref('" + last + "') }} AS c\nWHERE NOT EXISTS (\n"
            "  SELECT 1 FROM {{ ref('input_cells') }} AS i\n"
            "  WHERE i.sheet = c.sheet AND i.row_num = c.row_num AND i.col_num = c.col_num\n)")


_REF = re.compile(r"\{\{ ref\('([^']+)'\) \}\}")
_BRACE = re.compile(r"\{\{ '([{}])' \}\}")
_PLACEHOLDER = "cellproof_ref__"


def to_postgres(model: str, rows: bool = False) -> str:
    """A DuckDB dbt model -> the Postgres one: refs become placeholder tables while sqlglot translates (it cannot read
    Jinja), workbook braces are escaped again, and pg_temp helpers move to the target schema. `rows`: a model of cell
    rows, whose columns get their names (Postgres calls every unnamed column "?column?" and cannot create the table)."""
    from cellproof.postgres import translate
    sql = _BRACE.sub(lambda m: m.group(1), _REF.sub(lambda m: _PLACEHOLDER + m.group(1), model))
    if rows:
        sql = f"SELECT * FROM (\n{sql}\n) AS cellproof_rows(sheet, row_num, col_num, kind, num, txt)"
    sql = _escape(translate(sql))
    sql = re.sub(_PLACEHOLDER + r"(\w+)", lambda m: "{{ ref('" + m.group(1) + "') }}", sql)
    return re.sub(r"(?i)\bpg_temp\.", "{{ target.schema }}.", sql)


def helpers_macro() -> str:
    """The Postgres helpers (DuckDB's semantics) as a macro creating them in the target schema; raw, so nothing in
    their bodies is read as Jinja."""
    from cellproof.postgres import PRELUDE
    body = "{% endraw %}{{ target.schema }}.{% raw %}".join(PRELUDE.split("pg_temp."))
    return ("{% macro cellproof_helpers() %}\nCREATE SCHEMA IF NOT EXISTS {{ target.schema }};\n{% raw %}"
            + body + "{% endraw %}\n{% endmacro %}\n")


def write_project(book: Book, compiled: Compiled, out: Path, stem: str, target: str = "duckdb") -> None:
    name = project_name(stem)
    pg = target == "postgres"
    fix = to_postgres if pg else (lambda sql: sql)
    rows = (lambda sql: to_postgres(sql, rows=True)) if pg else fix
    batches_dir = out / "models" / "cellproof_cells"
    (out / "seeds").mkdir(parents=True, exist_ok=True)
    batches_dir.mkdir(parents=True, exist_ok=True)
    for old in batches_dir.glob("*.sql"):  # a rerun must not leave batches or checkpoints from a bigger workbook
        old.unlink()
    project = PROJECT.format(name=name)
    if pg:
        project = project.replace("num: double", "num: double precision").replace(": varchar", ": text")
        project += 'on-run-start: ["{{ cellproof_helpers() }}"]\n'
        (out / "macros").mkdir(exist_ok=True)
        (out / "macros" / "cellproof_helpers.sql").write_text(helpers_macro(), encoding="utf-8")
    (out / "dbt_project.yml").write_text(project, encoding="utf-8")
    (out / "profiles.yml").write_text(PROFILES_POSTGRES.replace("__NAME__", name) if pg else PROFILES, encoding="utf-8")
    write_seed(book, out / "seeds" / "input_cells.csv")
    n = len(compiled.batches)
    last = "input_cells"
    schema = SCHEMA
    used: set[str] = set()
    for k, batch in enumerate(compiled.batches, 1):
        if hasattr(batch, "wide_sql"):
            base, i = batch.model, 1
            while batch.model in used:  # sheet names that slug alike must not share a model
                i += 1
                batch.model = f"{base}_{i}"
            used.add(batch.model)
            (batches_dir / f"{batch.model}.sql").write_text(fix(dbt_region_wide(batch, sources_for(k))) + "\n", encoding="utf-8")
            (batches_dir / f"{model_name(k)}.sql").write_text(rows(dbt_region_cells(batch)) + "\n", encoding="utf-8")
            schema += (f'  - name: {batch.model}\n    description: "Region: rows of one sheet that repeat one formula down their rows, '
                       f'computed in one statement. Wide and readable; its {model_name(k)} twin holds the same values as long rows."\n')
        elif hasattr(batch, "render_rows"):  # a spill: its one statement, reading earlier cells as all_cells
            body = _escape(batch.render_rows("all_cells"))
            body = body.replace("WITH\n", f"WITH\nall_cells AS (\n{_union(sources_for(k))}\n),\n", 1)
            (batches_dir / f"{model_name(k)}.sql").write_text(rows(body) + "\n", encoding="utf-8")
        else:
            (batches_dir / f"{model_name(k)}.sql").write_text(rows(dbt_sql(batch, sources_for(k))) + "\n", encoding="utf-8")
        if k % CHECKPOINT == 0 or k == n:
            last = checkpoint_name(k)
            (batches_dir / f"{last}.sql").write_text(fix(_union(sources_for(k) + [model_name(k)])) + "\n", encoding="utf-8")
    (out / "models" / "workbook_cells.sql").write_text(fix(final_sql(last)) + "\n", encoding="utf-8")
    (out / "models" / "schema.yml").write_text(schema, encoding="utf-8")


def run_dbt(out: Path, env: dict[str, str] | None = None) -> tuple[bool, str]:
    """End-to-end check: dbt seeds the inputs and builds the models (`env` adds variables, e.g. postgres.dbt_env)."""
    exe = shutil.which("dbt")
    if exe is None:
        return False, "dbt is not installed - run: pip install cellproof[dbt]"
    env = {**os.environ, **(env or {}), "DBT_SEND_ANONYMOUS_USAGE_STATS": "False"}  # dbt phones home by default; cellproof never does
    proc = subprocess.run([exe, "build", "--project-dir", str(out), "--profiles-dir", str(out)],
                          capture_output=True, text=True, cwd=out, env=env)
    return proc.returncode == 0, (proc.stdout + proc.stderr)[-4000:]
