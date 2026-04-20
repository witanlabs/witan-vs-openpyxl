"""Task: read Sheet1 data from formulas.xls and report B3 value."""
import sys, openpyxl
PATH = sys.argv[1]
try:
    wb = openpyxl.load_workbook(PATH)
    ws = wb[wb.sheetnames[0]]
    print(f"ok: B3 = {ws['B3'].value}")
except Exception as e:
    print(f"FAIL ({type(e).__name__}): {e}")
