import pandas as pd

# Load the Excel file
file_path = 'Test - Advance Excel.xlsx'  # Update with the correct file path
xls = pd.ExcelFile(file_path)

# Load the 'Sales Data' and 'Retailers info' sheets
sales_data = pd.read_excel(xls, sheet_name='Sales Data', skiprows=1)
retailer_data = pd.read_excel(xls, sheet_name='Retailers info')

# Rename columns for easier access
sales_data.columns = ['Product ID', 'Order Date', 'Quantities Sold', 'Price Per Unit']
retailer_data.columns = ['Product ID', 'Retailer']  # Assuming there's no 'Cost Price'

# Merge sales data with retailer data
merged_data = pd.merge(sales_data, retailer_data, on='Product ID')

# Convert 'Order Date' to datetime format
merged_data['Order Date'] = pd.to_datetime(merged_data['Order Date'], errors='coerce')

# Drop rows with invalid 'Order Date'
merged_data = merged_data.dropna(subset=['Order Date'])

# Group data by Product ID to create separate sheets for each product
product_groups = merged_data.groupby('Product ID')

# Create an Excel writer object
output_file = 'Product_Order_Retailer_Report.xlsx'
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # Iterate through each product group and save to separate sheets
    for i, (product_id, product_data) in enumerate(product_groups, start=1):
        # Create a sheet name with product index (e.g., 5.1, 5.2, ...)
        sheet_name = f'5.{i}'
        # Select relevant columns: 'Order Date' and 'Retailer'
        product_report = product_data[['Order Date', 'Retailer']]
        # Write the product data to the sheet
        product_report.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Product order dates and retailer names report generated successfully! Saved to {output_file}")
