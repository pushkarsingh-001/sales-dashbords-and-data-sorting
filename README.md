# sales-dashbords-and-data-sorting

This project generates a synthetic sales dataset for 2024 and provides a simple data source suitable for dashboards, sorting, and reporting exercises.

## Overview

The repository contains:

- `generate_sales_data.py` – Python script that creates the sales dataset
- `sales_data.csv` – generated sales records for analysis
- `README.md` – project documentation

## Dataset contents

Each row in `sales_data.csv` includes:

- Date
- Product
- Category
- Region
- Sales Rep
- Units Sold
- Unit Price
- Revenue
- Cost
- Profit

The dataset includes product categories such as:

- Electronics
- Accessories
- Furniture

Products in the dataset include laptops, smartphones, tablets, monitors, headphones, desks, keyboards, mice, webcams, and chairs.

## How the data is generated

The script creates 500 records for the year 2024 by:

1. Selecting a random date in 2024
2. Choosing a product and category
3. Assigning a region (`North`, `South`, `East`, `West`)
4. Assigning a sales representative
5. Generating units sold based on product type
6. Calculating revenue, cost, and profit
7. Sorting records by date
8. Exporting the final data to CSV

## Run the generator

```bash
python generate_sales_data.py
```

This creates:

- `sales_data.csv`
- `sales_data.xlsx` (if `openpyxl` is installed)

## Requirements

- Python 3.x
- Optional: `openpyxl` for Excel export

Install the optional dependency with:

```bash
pip install openpyxl
```

## Example use cases

- Build dashboards in Excel, Power BI, or Tableau
- Sort and filter data by date, region, product, or sales rep
- Analyze sales performance by category
- Review profit and cost trends
- Use the dataset as sample data for data analysis practice

## Notes

This project is intended for demonstration, learning, and reporting exercises. The data is synthetic and generated algorithmically rather than pulled from a real business system.
