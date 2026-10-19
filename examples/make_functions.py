"""Writes examples/functions_unsaved.xlsx: the Excel behaviours Plan 2b could not confirm without Excel.

Open it in real Microsoft Excel (Windows or Mac), then File > Save As examples/functions.xlsx.
tests/test_excel_saved.py then checks cellproof's answers against Excel's own.
"""
from pathlib import Path

import openpyxl

DATA = {  # column A: a mixed range; B: numbers; C: a formula that returns ""; D: sorted brackets; E: unsorted
    "A1": 5, "A2": "5", "A3": "apple", "A4": "Banana", "A5": True, "A7": "a*b", "A8": "=1/0",
    "B1": 10, "B2": 20, "B3": 30, "B4": 40, "B5": 50, "B6": 60, "B7": 70, "B8": 80,
    "C1": '=IF(B1>0,"","x")', "D1": 0, "D2": 100, "D3": 500, "E1": 30, "E2": 10, "E3": 20, "H1": 5, "H2": 5, "H3": "#N/A",
}
CHECKS = {  # the open question each one answers
    "G1": '=COUNTIF(A1:A8,5)',            # does a number criterion also count the text "5"?
    "G2": '=COUNTIF(A1:A8,"<>5")',        # are empty and error cells counted by "<>5"?
    "G3": '=COUNTIF(A1:C1,"<>")',         # is a formula's "" counted as not empty?
    "G4": '=COUNTIF(A1:A8,">b")',         # text comparison ignores case
    "G5": '=SUMIF(A1:A8,"<>apple",B1:B8)',  # a matched error row (A8 is #DIV/0!, B8 a number)
    "G6": '=SUMIF(B1:B8,">30",A1:A8)',    # sum range holding an error in a matched row
    "G7": '="x"&0.00001',                 # number -> text for small numbers
    "G8": '="x"&1E+15',
    "G9": '="x"&123456789012345678',
    "G10": '=TRIM("  a   b  ")',
    "G11": '=PROPER("hELLO o\'neil 2nd")',
    "G12": '=MATCH(15,E1:E3)',            # approximate match on unsorted data
    "G13": '=MATCH(150,D1:D3,-1)',
    "G14": '=VLOOKUP("b",A3:B4,2)',       # approximate text match
    "G15": '=DAYS360(DATE(2026,1,30),DATE(2026,2,28))',
    "G16": '=DAYS360(DATE(2026,2,28),DATE(2026,3,31))',
    "G17": '=TEXT(-0.001,"0.00")',
    "G18": '=MOD(1,0.1)',
    "G19": '=ROUNDUP(0.1*3,1)',
    "G20": '=CHAR(128)',
    "G21": '=1<"a"',
    "G22": '="a"<TRUE',
    "G23": '=IFERROR(1/0,"")',
    "G24": '=AVERAGE(A1:A3,"3")',
    "G25": '=INDEX(F1:F3,2)',             # an empty cell found
    "G26": '=_xlfn.XLOOKUP(250,D1:D3,B1:B3,0,-1)',
    "G27": '=SUBTOTAL(9,B1:B3,G28)',
    "G28": '=SUBTOTAL(9,B1:B2)',
    "G29": '=SUM(A8,B1)',                 # error order: the first error wins
    "G30": '=DATE(1900,3,1)',
    "G31": '=WEEKDAY(DATE(2026,1,1))',
    "G32": '=TEXT(DATE(2026,3,6),"dd-mmm-yyyy")',
    "G33": '=DATE(1900,2,29)',
    "G34": '=MAX(1/0,H3)',                 # scalar error before a range error
    "G35": '=MAX(H3,1/0)',
    "G36": '=ROUNDUP(1,1000)',
    "G37": '=ROUNDUP(1,-1000)',
    "G38": '=EDATE(2958465,1)',
    "G39": '=EDATE(1,-1)',
    "G40": '=EOMONTH(2958465,1)',
    "G41": '=LEN(REPT("a",32767))',
    "G42": '=LEN(REPT("a",32768))',
    "G43": '=VLOOKUP(10,B1:C1,2,FALSE)',    # lookup of a formula returning empty text
    "G44": '=MEDIAN(B1:B2,B1:B1)',
    "G45": '=_xlfn.XLOOKUP(5,H1:H2,B1:B2,"none",-1,1)',
    "G46": '=_xlfn.XLOOKUP(5,H1:H2,B1:B2,"none",1,-1)',
    # Added after the first Windows run (2026-10-14): where does a number stop being written as plain digits?
    "G47": '="x"&0.000001',
    "G48": '="x"&0.0000001',
    "G49": '="x"&0.000000001',
    "G50": '="x"&0.0000000001',
    "G51": '="x"&0.00000000001',
    "G52": '="x"&1E+18',
    "G53": '="x"&1E+19',
    "G54": '="x"&1E+20',
    "G55": '="x"&1E+21',
    "G56": '="x"&1E+22',
    # Which error does Excel report when a typed-in error and an error inside a cell meet? (MAX(H3,1/0) is #DIV/0!)
    "G57": '=SUM(H3,1/0)',
    "G58": '=SUM(1/0,H3)',
    "G59": '=MIN(H3,1/0)',
    "G60": '=AVERAGE(H3,1/0)',
    "G61": '=PRODUCT(H3,1/0)',
    "G62": '=SUM(H2:H3,1/0)',
    # DAYS360, the US method, month ends other than the one seen so far
    "G63": '=DAYS360(DATE(2026,1,15),DATE(2026,4,30))',
    "G64": '=DAYS360(DATE(2026,1,15),DATE(2026,2,28))',
    "G65": '=DAYS360(DATE(2024,2,29),DATE(2024,3,31))',
    "G66": '=DAYS360(DATE(2026,1,31),DATE(2026,2,28))',
    # MOD of values that are not exactly representable; the 1900 date system around 29 February 1900
    "G67": '=MOD(0.3,0.1)',
    "G68": '=MOD(-1,0.1)',
    "G69": '=DATE(1900,1,60)',
    "G70": '=DATE(1900,2,30)',
    "G71": '=DATE(1900,3,0)',
    "G72": '=TEXT(-0.0004,"0.0%")',
    # Approximate match on unsorted data: cellproof does not guess, these record what Excel's binary search answers
    "G73": '=MATCH(15,E1:E3,-1)',
    "G74": '=MATCH(25,E1:E3,1)',
}

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Functions"
for ref, value in {**DATA, **CHECKS}.items():
    ws[ref] = value
wb.save(Path(__file__).parent / "functions_unsaved.xlsx")
