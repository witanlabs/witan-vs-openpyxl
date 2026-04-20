"""openpyxl what-if attempt: change Assumptions!B5 to 0.12 and report Summary!E23."""
import sys, shutil, openpyxl
SRC, DST = sys.argv[1], sys.argv[2]
shutil.copy(SRC, DST)

# Attempt 1: read cached value (data_only=True). Must reopen after write to read.
wb = openpyxl.load_workbook(DST)  # formulas
wb["Assumptions"]["B5"] = 0.12
wb.save(DST)

wb2 = openpyxl.load_workbook(DST, data_only=True)
cached = wb2["Summary"]["E23"].value
formula_wb = openpyxl.load_workbook(DST, data_only=False)
formula = formula_wb["Summary"]["E23"].value

print(f"openpyxl reports Summary!E23 = {cached!r} (data_only=True, cached)")
print(f"openpyxl formula string       = {formula!r}")
