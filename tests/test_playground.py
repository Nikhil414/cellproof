"""The browser playground's bridge (playground/cellproof_web.py), run natively: same conversion as the CLI."""
import io
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "playground"))

import cellproof_web  # noqa: E402


def test_lambdas_become_arrows_for_duckdb_1_2():
    assert cellproof_web.arrow("list_transform(v, lambda x, i: x / i)") == "list_transform(v, (x, i) -> x / i)"
    assert cellproof_web.arrow("list_filter(l, lambda x: x <> '')") == "list_filter(l, (x) -> x <> '')"


def test_run_gives_the_cli_report_and_the_dbt_project():
    res, z = cellproof_web.run((ROOT / "examples" / "functions.xlsx").read_bytes(), "functions.xlsx")
    r = json.loads(res)
    assert r["code"] == 0 and r["report"]["counts"] == {"PASS": 74, "FAIL": 0, "UNRESOLVED": 2}
    names = zipfile.ZipFile(io.BytesIO(z)).namelist()
    assert "functions_dbt/REPORT.md" in names and "functions_dbt/dbt_project.yml" in names


def test_an_unreadable_file_is_reported_not_raised():
    r = json.loads(cellproof_web.run(b"not a workbook", "x.xlsx")[0])
    assert r["code"] == 2 and "report" not in r
