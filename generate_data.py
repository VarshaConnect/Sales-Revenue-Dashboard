
import pandas as pd
import random
from datetime import datetime, timedelta

# Make results reproducible
random.seed(42)

# Sample data
products = {
    "Laptop": ("Electronics", 50000, 35000),
    "Mobile": ("Electronics", 20000, 14000),
    "Tablet": ("Electronics", 25000, 18000),
    "Headphones": ("Accessories", 3000, 1500),
    "Keyboard": ("Accessories", 2000, 1000),
    "Mouse": ("Accessories", 1000, 500),
    "Chair": ("Furniture", 8000, 5000),
    "Desk": ("Furniture", 12000, 8000)
}

regions = ["North", "South", "East", "West", "Central"]

data = []

start_date = datetime(2025, 1, 1)

for i in range(1, 501):

    product = random.choice(list(products.keys()))
    category, price, cost = products[product]

    quantity = random.randint(1, 10)

    order_date = start_date + timedelta(
        days=random.randint(0, 364)
    )

    revenue = quantity * price
    total_cost = quantity * cost
    profit = revenue - total_cost

    data.append({
        "Order ID": f"ORD{i:04d}",
        "Order Date": order_date,
        "Product": product,
        "Category": category,
        "Region": random.choice(regions),
        "Quantity": quantity,
        "Unit Price": price,
        "Revenue": revenue,
        "Cost": total_cost,
        "Profit": profit
    })

# Create DataFrame
df = pd.DataFrame(data)

# Export to CSV
df.to_csv("sales_data.csv", index=False)

print("Sales dataset generated successfully!")
print("Total Records:", len(df))
print(df.head())
