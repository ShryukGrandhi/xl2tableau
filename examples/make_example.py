"""Write examples/sales_model.xlsx: a 200-row sales model with one hardcoded override.

    python examples/make_example.py
    xl2tableau convert examples/sales_model.xlsx -o out
"""
import os

from xl2tableau.sample import build_sales_workbook

if __name__ == "__main__":
    wb, _ = build_sales_workbook(200, seed=7, override_row=9)
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sales_model.xlsx")
    wb.save(path)
    print(f"wrote {path}")
