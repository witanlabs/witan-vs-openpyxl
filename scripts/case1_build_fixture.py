"""Build pricing.xlsx with an NPV depending on Assumptions!B5, then open in Excel
   via xlwings to force calc and cache the computed value (so openpyxl with
   data_only=True has something to return)."""
import os, sys, openpyxl, xlwings as xw

OUT = sys.argv[1]
wb = openpyxl.Workbook()
a = wb.active
a.title = "Assumptions"
a["A5"] = "Discount rate"
a["B5"] = 0.08

cf = wb.create_sheet("Cashflows")
cf["A1"] = "Year"; cf["B1"] = "Cashflow"
# 10-year stream of cashflows
flows = [-100000, 20000, 25000, 27000, 30000, 32000, 34000, 35000, 36000, 38000, 40000]
for i, v in enumerate(flows):
    cf.cell(row=i+2, column=1, value=i)
    cf.cell(row=i+2, column=2, value=v)

s = wb.create_sheet("Summary")
s["A1"] = "NPV @ rate"
# NPV of positive flows + initial outlay
s["E23"] = "=Cashflows!B2 + NPV(Assumptions!B5, Cashflows!B3:B12)"

wb.save(OUT)

# Open in Excel to cache computed values
app = xw.App(visible=False)
try:
    book = app.books.open(os.path.abspath(OUT))
    app.calculate()
    cached = book.sheets["Summary"].range("E23").value
    book.save()
    book.close()
finally:
    app.quit()

print(f"cached Summary!E23 (rate=8%) = {cached:.2f}")
