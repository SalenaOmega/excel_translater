"""
generate_sample_data.py
Create sample_data/ with two Excel files, one CSV, and one HTML table.
Run this once to get demo data before running main.py.
"""
from __future__ import annotations

import os
import random
from datetime import date

import pandas as pd

random.seed(42)

CATEGORIES = ["Electronics", "Groceries", "Clothing", "Home", "Toys"]
PRODUCTS = {
    "Electronics": ["Headphones", "USB Cable", "Power Bank", "Webcam"],
    "Groceries":   ["Coffee", "Rice 5kg", "Olive Oil", "Tea"],
    "Clothing":    ["T-Shirt", "Jeans", "Socks", "Jacket"],
    "Home":        ["Mug", "Towel Set", "Lamp", "Cushion"],
    "Toys":        ["Puzzle", "Toy Car", "Board Game", "Doll"],
}
REGIONS = ["North", "South", "East", "West"]


def make_rows(year: int, month: int, n: int) -> list[dict]:
    rows = []
    for _ in range(n):
        cat = random.choice(CATEGORIES)
        prod = random.choice(PRODUCTS[cat])
        qty = random.randint(1, 20)
        price = round(random.uniform(5, 200), 2)
        rows.append({
            "Date ": date(year, month, random.randint(1, 28)).strftime("%Y/%m/%d"),
            "Category": cat,
            "Product": f"  {prod} ",   # intentional leading/trailing spaces
            "Region": random.choice(REGIONS),
            "Quantity": qty,
            "Unit Price": price,
            "Amount": round(qty * price, 2),
        })
    df = pd.DataFrame(rows)
    # inject one duplicate and one blank row to demonstrate cleaning
    df = pd.concat([df, df.iloc[[0]]], ignore_index=True)
    df.loc[len(df)] = [None] * df.shape[1]
    # inject one anomaly row (negative amount)
    df.loc[len(df)] = [
        date(year, month, 15).strftime("%Y/%m/%d"),
        "Electronics", "  Webcam ", "North", -5, 99.0, -495.0,
    ]
    return df


def main() -> None:
    out = "sample_data"
    os.makedirs(out, exist_ok=True)

    # Two Excel files
    for year, month, n in [(2026, 1, 40), (2026, 2, 35)]:
        df = make_rows(year, month, n)
        path = os.path.join(out, f"sales_{year}_{month:02d}.xlsx")
        df.to_excel(path, index=False)
        print(f"Wrote {path}")

    # One CSV with different column name aliases to show rename_fields
    df_csv = make_rows(2026, 3, 30).rename(columns={
        "Date ": "date", "Amount": "销售额", "Quantity": "qty"
    })
    csv_path = os.path.join(out, "sales_2026_03.csv")
    df_csv.to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"Wrote {csv_path}")

    # One HTML file with a static table
    df_html = make_rows(2026, 4, 20)
    html_path = os.path.join(out, "sales_2026_04.html")
    df_html.to_html(html_path, index=False)
    print(f"Wrote {html_path}")


if __name__ == "__main__":
    main()
