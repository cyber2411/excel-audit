"""Build samples/bakery-demo.xlsx (run: python samples/make_sample.py).

A tiny home-bakery workbook with one deliberate fault per check, so the
README examples are reproducible:
- Sales!D5 adds instead of multiplying (inconsistent column formula)
- Summary!B2 points at a sheet that does not exist (broken ref)
- Summary!B3 contains #REF! (broken ref)
- Summary!B4/B5 use TODAY() and INDIRECT() (volatile functions)
"""

from pathlib import Path

import openpyxl

wb = openpyxl.Workbook()
sales = wb.active
sales.title = "Sales"
sales.append(["Item", "Price", "Qty", "Total"])
sales.append(["Chocolate Cake", 25, 2, "=B2*C2"])
sales.append(["Banana Bread", 12, 5, "=B3*C3"])
sales.append(["Cupcakes x6", 18, 3, "=B4*C4"])
sales.append(["Brownies", 15, 4, "=B5+C5"])  # <-- the odd one out

summary = wb.create_sheet("Summary")
summary["A1"] = "Metric"
summary["B1"] = "Value"
summary["A2"] = "Missing sheet ref"
summary["B2"] = "='Stock List'!A1"
summary["A3"] = "Deleted ref"
summary["B3"] = "=SUM(#REF!)"
summary["A4"] = "Today"
summary["B4"] = "=TODAY()"
summary["A5"] = "Indirect lookup"
summary["B5"] = '=INDIRECT("Sales!A2")'

out = Path(__file__).with_name("bakery-demo.xlsx")
wb.save(out)
print(f"wrote {out}")
