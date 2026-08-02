"""
One-off script used to generate data/sample_sales_data.csv for the demo.
Not part of the reporting pipeline itself — kept here for transparency/reproducibility.
"""
import numpy as np
import pandas as pd

rng = np.random.default_rng(42)

N = 3000
start = pd.Timestamp("2024-01-01")
end = pd.Timestamp("2025-12-31")
days = (end - start).days

regions = ["North", "South", "East", "West", "Central"]
region_w = [0.22, 0.20, 0.18, 0.24, 0.16]

categories = {
    "Electronics": ["Headphones", "Smartphone", "Laptop", "Smartwatch", "Tablet"],
    "Furniture": ["Office Chair", "Desk", "Bookshelf", "Sofa", "Bed Frame"],
    "Clothing": ["T-Shirt", "Jacket", "Jeans", "Sneakers", "Cap"],
    "Home & Kitchen": ["Blender", "Cookware Set", "Vacuum Cleaner", "Air Fryer", "Lamp"],
    "Stationery": ["Notebook", "Pen Set", "Backpack", "Whiteboard", "Desk Organizer"],
}
cat_list = list(categories.keys())
cat_w = [0.24, 0.16, 0.26, 0.18, 0.16]

segments = ["Consumer", "Corporate", "Home Office"]
seg_w = [0.5, 0.3, 0.2]

base_price = {
    "Electronics": (60, 900), "Furniture": (80, 700), "Clothing": (10, 90),
    "Home & Kitchen": (15, 200), "Stationery": (3, 40),
}

rows = []
for i in range(1, N + 1):
    order_date = start + pd.Timedelta(days=int(rng.integers(0, days + 1)))
    # mild seasonality: more sales in Nov/Dec
    if order_date.month in (11, 12) and rng.random() < 0.35:
        order_date = order_date.replace(day=min(order_date.day, 28))

    region = rng.choice(regions, p=region_w)
    category = rng.choice(cat_list, p=cat_w)
    product = rng.choice(categories[category])
    segment = rng.choice(segments, p=seg_w)

    lo, hi = base_price[category]
    unit_price = round(rng.uniform(lo, hi), 2)
    quantity = int(rng.integers(1, 8))
    discount = float(rng.choice([0, 0, 0, 0.05, 0.1, 0.15, 0.2], p=[0.35,0.15,0.15,0.12,0.1,0.08,0.05]))
    sales = round(unit_price * quantity * (1 - discount), 2)
    margin_rate = rng.uniform(0.05, 0.35)
    profit = round(sales * margin_rate - (5 if discount > 0.1 else 0), 2)

    rows.append({
        "Order ID": f"ORD-{10000+i}",
        "Order Date": order_date.strftime("%Y-%m-%d"),
        "Customer Name": f"Customer {rng.integers(1, 900)}",
        "Segment": segment,
        "Region": region,
        "Category": category,
        "Sub Category": product,
        "Quantity": quantity,
        "Unit Price": unit_price,
        "Discount": discount,
        "Sales": sales,
        "Profit": profit,
    })

df = pd.DataFrame(rows).sort_values("Order Date").reset_index(drop=True)
df.to_csv("data/sample_sales_data.csv", index=False)
print(df.shape)
print(df.head())
