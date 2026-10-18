"""Writes examples/commission_unsaved.xlsx (formulas, no saved values).

Open it in real Microsoft Excel, then File > Save As examples/commission.xlsx.
Excel stores its own answers in the file; tests/test_excel_saved.py checks unexcel against them.
"""
from pathlib import Path

import openpyxl

wb = openpyxl.Workbook()
sales = wb.active
sales.title = "Sales"
for row in [["Rep", "Region", "Amount"], ["Asha", "West", 80000], ["asha", "West", 40000],
            ["Vikram", "South", 60000], ["Vikram", "South", "15,000"]]:
    sales.append(row)
c = wb.create_sheet("Commission")
c.append(["Rep", "Total", "Bonus"])
c.append(["Asha", "=SUMIFS(Sales!C:C,Sales!A:A,A2)", "=IF(B2>100000,B2*5%,B2*2%)"])
c.append(["Vikram", "=SUMIFS(Sales!C:C,Sales!A:A,A3)", 1500])
c["B5"] = "=SUM(B2:B3)"
wb.save(Path(__file__).parent / "commission_unsaved.xlsx")
