import pandas as pd

# Load the Excel file
file_path = """Test - Advance Excel.xlsx"""  # Update with the correct file path
xls = pd.ExcelFile(file_path)

# Load the 'Sales Data' and 'Retailers info' sheets
sales_data = pd.read_excel(xls, sheet_name='Sales Data', skiprows=1)
retailer_data = pd.read_excel(xls, sheet_name='Retailers info')

# Rename columns for easier access
sales_data.columns = ['Product ID', 'Order Date', 'Quantities Sold', 'Price Per Unit']
retailer_data.columns = ['Product ID', 'Retailer']

# Convert 'Order Date' to datetime format, ignoring errors for invalid dates
sales_data['Order Date'] = pd.to_datetime(sales_data['Order Date'], errors='coerce')

# Drop rows with invalid 'Order Date'
sales_data = sales_data.dropna(subset=['Order Date'])

# Extract product category from 'Product ID' (first letter represents the category)
sales_data['Product Category'] = sales_data['Product ID'].str[0]

# Merge sales data with retailer data
merged_data = pd.merge(sales_data, retailer_data, on='Product ID')

# Group by 'Retailer' and 'Product Category' to generate retailer-wise sales report
retailer_sales_report = merged_data.groupby(['Retailer', 'Product Category']).agg({'Quantities Sold': 'sum'}).reset_index()

# Save the retailer-wise sales report to an Excel file
retailer_sales_report.to_excel('Retailer_Sales_Report.xlsx', index=False)

print("Retailer-wise sales report generated successfully!")
