import pandas as pd

# Load the Excel file
file_path = 'Test - Advance Excel.xlsx'  # Update with the correct file path
xls = pd.ExcelFile(file_path)

# Load the 'Sales Data' and 'Retailers info' sheets
sales_data = pd.read_excel(xls, sheet_name='Sales Data', skiprows=1)
retailer_data = pd.read_excel(xls, sheet_name='Retailers info')

# Rename columns for easier access
sales_data.columns = ['Product ID', 'Order Date', 'Quantities Sold', 'Price Per Unit']
retailer_data.columns = ['Product ID', 'Retailer']  # Assuming 'Cost Price' is not needed for this task

# Convert 'Order Date' to datetime format
sales_data['Order Date'] = pd.to_datetime(sales_data['Order Date'], errors='coerce')

# Drop rows with invalid 'Order Date'
sales_data = sales_data.dropna(subset=['Order Date'])

# Extract month and year from 'Order Date'
sales_data['Month'] = sales_data['Order Date'].dt.to_period('M')

# Merge sales data with retailer data to get retailer names
merged_data = pd.merge(sales_data, retailer_data, on='Product ID')

# Calculate total sales for each item
merged_data['Total Sales'] = merged_data['Quantities Sold'] * merged_data['Price Per Unit']

# Group data by month and retailer, then sum the total sales for each retailer per month
monthly_sales_by_retailer = merged_data.groupby(['Month', 'Retailer'])['Total Sales'].sum().reset_index()

# Find the retailer with the highest sales for each month
highest_sales_by_month = monthly_sales_by_retailer.loc[monthly_sales_by_retailer.groupby('Month')['Total Sales'].idxmax()]

# Save the result to an Excel file
highest_sales_by_month.to_excel('Highest_Sales_By_Month.xlsx', index=False)

print("Retailers with the highest sales for each month report generated successfully!")
