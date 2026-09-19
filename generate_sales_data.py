import csv
import random
from datetime import datetime, timedelta
import os

# Product catalog with prices
products = {
    'Laptop': {'category': 'Electronics', 'price': 899, 'cost': 540},
    'Smartphone': {'category': 'Electronics', 'price': 699, 'cost': 420},
    'Tablet': {'category': 'Electronics', 'price': 549, 'cost': 330},
    'Monitor': {'category': 'Electronics', 'price': 449, 'cost': 270},
    'Headphones': {'category': 'Accessories', 'price': 149, 'cost': 60},
    'Desk Chair': {'category': 'Furniture', 'price': 399, 'cost': 240},
    'Keyboard': {'category': 'Accessories', 'price': 79, 'cost': 40},
    'Mouse': {'category': 'Accessories', 'price': 49, 'cost': 25},
    'Webcam': {'category': 'Accessories', 'price': 129, 'cost': 65},
    'Standing Desk': {'category': 'Furniture', 'price': 599, 'cost': 360},
}

regions = ['North', 'South', 'East', 'West']
sales_reps = ['Jane Smith', 'John Doe', 'Bob Wilson', 'Alice Brown']

# Generate data for 500 lines
data = []
start_date = datetime(2024, 1, 1)

for i in range(500):
    # Random date in 2024
    days_offset = random.randint(0, 364)
    date = start_date + timedelta(days=days_offset)
    
    # Random product
    product_name = random.choice(list(products.keys()))
    product_info = products[product_name]
    
    # Random region and sales rep
    region = random.choice(regions)
    sales_rep = random.choice(sales_reps)
    
    # Random units sold (vary by product type)
    if product_name in ['Laptop', 'Smartphone']:
        units = random.randint(5, 35)
    elif product_name in ['Tablet', 'Monitor']:
        units = random.randint(3, 20)
    elif product_name in ['Headphones', 'Keyboard', 'Mouse']:
        units = random.randint(10, 60)
    else:
        units = random.randint(2, 15)
    
    # Calculate financials
    unit_price = product_info['price']
    revenue = units * unit_price
    cost_per_unit = product_info['cost']
    total_cost = units * cost_per_unit
    profit = revenue - total_cost
    
    data.append({
        'Date': date.strftime('%Y-%m-%d'),
        'Product': product_name,
        'Category': product_info['category'],
        'Region': region,
        'Sales_Rep': sales_rep,
        'Units_Sold': units,
        'Unit_Price': unit_price,
        'Revenue': revenue,
        'Cost': total_cost,
        'Profit': profit
    })

# Sort by date
data.sort(key=lambda x: x['Date'])

# Write to CSV
csv_file = 'sales_data.csv'
fieldnames = ['Date', 'Product', 'Category', 'Region', 'Sales_Rep', 'Units_Sold', 'Unit_Price', 'Revenue', 'Cost', 'Profit']

with open(csv_file, 'w', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(data)

print(f"Generated {len(data)} lines of sales data to {csv_file}")

# Also create Excel file
try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Data"
    
    # Write headers
    for col, field in enumerate(fieldnames, 1):
        cell = ws.cell(row=1, column=col, value=field)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.font = Font(bold=True, color="FFFFFF")
    
    # Write data
    for row_idx, row_data in enumerate(data, 2):
        for col_idx, field in enumerate(fieldnames, 1):
            ws.cell(row=row_idx, column=col_idx, value=row_data[field])
    
    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 20)
        ws.column_dimensions[column].width = adjusted_width
    
    excel_file = 'sales_data.xlsx'
    wb.save(excel_file)
    print(f"Created Excel file: {excel_file}")
    
except ImportError:
    print("openpyxl not installed. Only CSV file created.")
    print("To create Excel file, install openpyxl: pip install openpyxl")

print("\nData generation complete!")
print(f"- CSV: sales_data.csv ({len(data)} rows)")
print(f"- Excel: sales_data.xlsx ({len(data)} rows)")
