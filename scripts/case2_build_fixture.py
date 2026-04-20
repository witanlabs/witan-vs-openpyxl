"""Build circular.xlsx: seed with openpyxl, then open+iterate+save via Excel."""
import os, sys, openpyxl, xlwings as xw
OUT = os.path.abspath(sys.argv[1])

# Step 1: create with openpyxl, seed circular cells with explicit 0 values in <v>
wb = openpyxl.Workbook()
wb.calculation.iterate = True
wb.calculation.iterateCount = 200
wb.calculation.iterateDelta = 0.0001

ip = wb.active; ip.title = "Inputs"
ip["A1"] = "Revenue"; ip["B1"] = 100000
ip["A2"] = "Opex ratio"; ip["B2"] = 0.4
ip["A3"] = "Tax rate"; ip["B3"] = 0.3
ip["A4"] = "Bonus rate"; ip["B4"] = 0.1

m = wb.create_sheet("Model")
m["A1"] = "Revenue";             m["B1"] = "=Inputs!B1"
m["A2"] = "Opex";                m["B2"] = "=Model!B1*Inputs!B2"
m["A3"] = "Profit before bonus"; m["B3"] = "=Model!B1-Model!B2-Model!B4"
m["A4"] = "Bonus";               m["B4"] = "=Model!B3*Inputs!B4"
m["A5"] = "Profit after bonus";  m["B5"] = "=Model!B3"
m["A6"] = "Tax";                 m["B6"] = "=Model!B5*Inputs!B3"
m["A7"] = "Net income";          m["B7"] = "=Model!B5-Model!B6"
wb.save(OUT)

# Step 2: open in Excel, run many iterations, save cached values
app = xw.App(visible=False)
try:
    app.display_alerts = False
    app.api.iteration = True
    app.api.max_iteration = 500
    app.api.max_change = 0.00001
    book = app.books.open(OUT)
    for _ in range(20):
        app.calculate()
    sh = book.sheets["Model"]
    print("B3 =", sh.range("B3").value)
    print("B4 =", sh.range("B4").value)
    print("B6 =", sh.range("B6").value)
    print("B7 =", sh.range("B7").value)
    book.save()
    book.close()
finally:
    app.quit()
