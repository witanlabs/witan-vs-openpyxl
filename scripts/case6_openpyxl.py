"""Reproduce openpyxl issue #2311 — single-series LineChart misinterprets category axis.
Exact repro from the issue, adapted to the project's layout."""
import sys
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference

OUT = sys.argv[1]
wb = Workbook()
ws = wb.active
ws.title = "Data"

ws.append(["Services", "2025-01", "2025-02", "2025-03", "2025-04", "2025-05"])
ws.append(["Compute",  100,        200,        150,        180,        220])

chart = LineChart()
chart.title = "Compute usage (openpyxl single-series)"
cats = Reference(ws, min_col=2, max_col=6, min_row=1)
vals = Reference(ws, min_col=2, max_col=6, min_row=2)
chart.add_data(vals, titles_from_data=False)
chart.set_categories(cats)

ws.add_chart(chart, "A5")
wb.save(OUT)
print(f"saved {OUT}")
