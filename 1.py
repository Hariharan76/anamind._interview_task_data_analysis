import pandas as pd

# Load the Excel file
file_path = """Test - Advance Excel.xlsx"""  # Update with the correct file path
xls = pd.ExcelFile(file_path)

# Load the 'Sales Data' sheet and skip the first row (header)
sales_data = pd.read_excel(xls, sheet_name='Sales Data', skiprows=1)

# Rename columns for easier access
sales_data.columns = ['Product ID', 'Order Date', 'Quantities Sold', 'Price Per Unit']

# Convert 'Order Date' to datetime format, ignoring errors for invalid dates
sales_data['Order Date'] = pd.to_datetime(sales_data['Order Date'], errors='coerce')

# Drop rows with invalid 'Order Date'
sales_data = sales_data.dropna(subset=['Order Date'])

# Extract year and month from 'Order Date'
sales_data['Year'] = sales_data['Order Date'].dt.year
sales_data['Month'] = sales_data['Order Date'].dt.month

# Extract product category from 'Product ID' (first letter represents the category)
sales_data['Product Category'] = sales_data['Product ID'].str[0]

# Filter data for the year 2022
sales_data_2022 = sales_data[sales_data['Year'] == 2022]

# Group by 'Product Category' and 'Month' to generate the monthly sales report
monthly_sales_report = sales_data_2022.groupby(['Product Category', 'Month']).agg({'Quantities Sold': 'sum'}).reset_index()

# Save the monthly sales report to an Excel file
monthly_sales_report.to_excel('Monthly_Sales_Report.xlsx', index=False)

print("Monthly sales report generated successfully!")
