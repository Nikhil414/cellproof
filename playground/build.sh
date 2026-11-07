#!/usr/bin/env bash
# Builds the browser playground into a folder (default _site): the page, the worker, the bridge, the example
# workbook and a wheel of this checkout's cellproof. Serve the folder over http(s) (Pyodide cannot load from file://).
set -euo pipefail
out="${1:-_site}"
rm -rf "$out" && mkdir -p "$out"
python -m pip wheel --no-deps --quiet -w "$out" .
python -m pip download --no-deps --only-binary :all: --quiet -d "$out" openpyxl==3.1.5 et-xmlfile==2.0.0
cp playground/index.html playground/worker.js playground/cellproof_web.py examples/functions.xlsx "$out"/
# DuckDB for Pyodide (duckdb/duckdb-pyodide; jsDelivr mirrors the repo and is fast, duckdb.github.io can crawl)
duck=duckdb-1.2.0-cp312-cp312-pyodide_2024_0_wasm32.whl
if [ -f "${WHEEL_CACHE:-/nonexistent}/$duck" ]; then cp "$WHEEL_CACHE/$duck" "$out/"; else  # WHEEL_CACHE: a folder holding it
curl -sSfL --retry 3 -o "$out/$duck" "https://cdn.jsdelivr.net/gh/duckdb/duckdb-pyodide@main/wheels/$duck" \
  || curl -sSfL --retry 3 -o "$out/$duck" "https://duckdb.github.io/duckdb-pyodide/wheels/$duck"
fi
echo "48ace5ee2bf26d13943e2260c3143ae3eb87c727c268f6020d5424e2fbe6d2ea  $out/$duck" | shasum -a 256 -c -
(cd "$out" && ls cellproof-*.whl openpyxl-*.whl et_xmlfile-*.whl) > "$out/wheels.txt"  # what the worker installs
echo "built $out:" $(cat "$out/wheels.txt")
