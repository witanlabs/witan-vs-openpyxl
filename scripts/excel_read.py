"""Generic: open a workbook in Excel, force recalc, and print selected addresses.
Usage: python excel_read.py <file> <addr1> [<addr2> ...]"""
import os, sys, xlwings as xw

PATH, ADDRS = sys.argv[1], sys.argv[2:]
app = xw.App(visible=False)
try:
    app.display_alerts = False
    book = app.books.open(os.path.abspath(PATH))
    app.calculate()
    for a in ADDRS:
        if "!" in a:
            sheet, addr = a.split("!", 1)
        else:
            sheet, addr = book.sheets[0].name, a
        v = book.sheets[sheet].range(addr).value
        print(f"  {sheet}!{addr} = {v}")
    book.close()
finally:
    app.quit()
