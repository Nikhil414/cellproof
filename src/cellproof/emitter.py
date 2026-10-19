"""Write the compiled workbook as a dbt project (DuckDB profile included).

The models run the SQL the verifier ran: seed input_cells -> cellproof_cells/cells_0001 -> ... Each batch model
holds only its own rows; a checkpoint model (cp_NNNN) unions everything every 64 batches and after the last
batch, so no model reads more than ~65 sources. workbook_cells keeps the formula cells only.
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


def write_project(book: Book, compiled: Compiled, out: Path, stem: str) -> None:
    name = project_name(stem)
    batches_dir = out / "models" / "cellproof_cells"
    (out / "seeds").mkdir(parents=True, exist_ok=True)
    batches_dir.mkdir(parents=True, exist_ok=True)
    for old in batches_dir.glob("*.sql"):  # a rerun must not leave batches or checkpoints from a bigger workbook
        old.unlink()
    (out / "dbt_project.yml").write_text(PROJECT.format(name=name), encoding="utf-8")
    (out / "profiles.yml").write_text(PROFILES, encoding="utf-8")
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
            (batches_dir / f"{batch.model}.sql").write_text(dbt_region_wide(batch, sources_for(k)) + "\n", encoding="utf-8")
            (batches_dir / f"{model_name(k)}.sql").write_text(dbt_region_cells(batch) + "\n", encoding="utf-8")
            schema += (f'  - name: {batch.model}\n    description: "Region: rows of one sheet that repeat one formula down their rows, '
                       f'computed in one statement. Wide and readable; its {model_name(k)} twin holds the same values as long rows."\n')
        else:
            (batches_dir / f"{model_name(k)}.sql").write_text(dbt_sql(batch, sources_for(k)) + "\n", encoding="utf-8")
        if k % CHECKPOINT == 0 or k == n:
            last = checkpoint_name(k)
            (batches_dir / f"{last}.sql").write_text(_union(sources_for(k) + [model_name(k)]) + "\n", encoding="utf-8")
    (out / "models" / "workbook_cells.sql").write_text(final_sql(last) + "\n", encoding="utf-8")
    (out / "models" / "schema.yml").write_text(schema, encoding="utf-8")


def run_dbt(out: Path) -> tuple[bool, str]:
    """End-to-end check: dbt seeds the inputs and builds the models on DuckDB."""
    exe = shutil.which("dbt")
    if exe is None:
        return False, "dbt is not installed - run: pip install cellproof[dbt]"
    env = {**os.environ, "DBT_SEND_ANONYMOUS_USAGE_STATS": "False"}  # dbt phones home by default; cellproof never does
    proc = subprocess.run([exe, "build", "--project-dir", str(out), "--profiles-dir", str(out)],
                          capture_output=True, text=True, cwd=out, env=env)
    return proc.returncode == 0, (proc.stdout + proc.stderr)[-4000:]
