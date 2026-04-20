"""openpyxl what-if: change Inputs!B4 (bonus rate) from 0.1 to 0.2, report Model!B7 (net income)."""
import sys, shutil, openpyxl
SRC, DST = sys.argv[1], sys.argv[2]
shutil.copy(SRC, DST)
wb = openpyxl.load_workbook(DST)
wb["Inputs"]["B4"] = 0.2
wb.save(DST)

wb2 = openpyxl.load_workbook(DST, data_only=True)
print(f"openpyxl Model!B7 (net income)  = {wb2['Model']['B7'].value!r}")
print(f"openpyxl Model!B3 (profit)      = {wb2['Model']['B3'].value!r}")
print(f"openpyxl Model!B4 (bonus)       = {wb2['Model']['B4'].value!r}")
