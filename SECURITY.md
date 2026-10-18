# Security policy

## Reporting a vulnerability

Please do not open a public issue. Use GitHub's private reporting: the repository's **Security** tab, then
**Report a vulnerability**. Include a minimal workbook or steps that reproduce it. You will get an answer within a
week.

## What unexcel promises

unexcel is meant to run on workbooks you did not write. It aims to guarantee:

- **No network access** with workbook data. Conversion runs locally on DuckDB. The generated dbt project turns off
  dbt's anonymous usage statistics, and `--dbt` runs dbt with them off too.
- **No code execution from a workbook.** Formulas are parsed into a tree and translated; they are never evaluated as
  Python. VBA macros in `.xlsm` files are ignored. Workbook text reaches SQL only as escaped string literals; column
  headers become SQL names only after being reduced to `[a-z0-9_]`; `{` and `}` are escaped so dbt's Jinja templating
  cannot run workbook text.
- **Bounded input.** A file that unzips to more than 2 GB is refused; formulas nested too deeply or whose SQL would
  exceed 1 MB are UNRESOLVED instead of exhausting memory.

## Known residual risks

- **XML parsing** is done by openpyxl with Python's XML parser, which protects against external entities and, with
  the expat versions shipped with supported Pythons, against entity-expansion bombs. Install `defusedxml` for extra
  hardening; openpyxl uses it automatically.
- **Resource use:** a large but legitimate workbook can still take minutes and gigabytes of memory.
- **The `compute` extra** uses the third-party `formulas` library to calculate expected values. It does not execute
  workbook text as Python, but it is a larger code surface; leave it uninstalled for untrusted files if you do not need
  it.
- **`--out`**: a rerun deletes old `*.sql` files in `<out>/models/unexcel_cells/`. Give each conversion its own folder.
- **The generated CSV seed** (`seeds/input_cells.csv`) holds the workbook's text as data. Opening it in a spreadsheet
  program may interpret text starting with `=` as a formula, as with any CSV.
