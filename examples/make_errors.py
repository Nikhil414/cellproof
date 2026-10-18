"""Writes examples/errors_unsaved.xlsx: one cell per Excel error rule unexcel relies on.

Open it in real Microsoft Excel, then File > Save As examples/errors.xlsx.
Excel stores its own answers in the file; tests/test_excel_saved.py checks unexcel against them.
"""
from pathlib import Path

import openpyxl

CELLS = {
    "A1": 0, "A2": -8, "A3": "South",
    "B1": "=1/A1", "B2": "=A1^-1", "B3": "=A1^0", "B4": "=A2^(1/3)", "B5": "=#N/A",
    "B6": "=B5+1", "B7": "=B1+B5", "B8": "=IFERROR(1/A1,-1)", "B9": "=IF(1/A1>1,1,2)",
    # SUM reports the first error row by row: E1 (#DIV/0!) comes before D2 (#N/A)
    "D1": 1, "E1": "=1/0", "D2": "=#N/A", "E2": 2, "F1": "=SUM(D1:E2)",
    # VLOOKUP: a missing key is #N/A; a matched row returns the error its cell holds
    "G1": "West", "H1": 0.05, "G2": "South", "H2": "=1/0",
    "I1": '=VLOOKUP("North",G1:H2,2,FALSE)', "I2": "=VLOOKUP(A3,G1:H2,2,FALSE)",
}

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Errors"
for ref, value in CELLS.items():
    ws[ref] = value
wb.save(Path(__file__).parent / "errors_unsaved.xlsx")
