import pandas as pd

# Load the Excel file
file_path = 'Test - Advance Excel.xlsx'  # Update with the correct file path
xls = pd.ExcelFile(file_path)

# Load the 'Sales Data' and 'Retailers info' sheets
sales_data = pd.read_excel(xls, sheet_name='Sales Data', skiprows=1)
retailer_data = pd.read_excel(xls, sheet_name='Retailers info')

# Check the columns in 'Retailers info' to avoid mismatch
print(retailer_data.columns)

# Rename columns for easier access based on actual data
# Update according to the number of columns in 'Retailers info'
if len(retailer_data.columns) == 2:
    retailer_data.columns = ['Product ID', 'Retailer']  # If no 'Cost Price' column exists
elif len(retailer_data.columns) == 3:
    retailer_data.columns = ['Product ID', 'Retailer', 'Cost Price']

# Rename columns for 'Sales Data'
sales_data.columns = ['Product ID', 'Order Date', 'Quantities Sold', 'Price Per Unit']

# Convert 'Order Date' to datetime format, ignoring errors for invalid dates
sales_data['Order Date'] = pd.to_datetime(sales_data['Order Date'], errors='coerce')

# Drop rows with invalid 'Order Date'
sales_data = sales_data.dropna(subset=['Order Date'])

# Merge sales data with retailer data
merged_data = pd.merge(sales_data, retailer_data, on='Product ID')

# Calculate total sales for each item
merged_data['Total Sales'] = merged_data['Quantities Sold'] * merged_data['Price Per Unit']

# Sort the data by total sales to find top and bottom selling items
sorted_data = merged_data.sort_values(by='Total Sales', ascending=False)

# Get the top 10 and bottom 10 best-selling items
top_10 = sorted_data.head(10)
bottom_10 = sorted_data.tail(10)

# Combine both lists into a single dataframe
best_worst_sellers = pd.concat([top_10, bottom_10])

# Select relevant columns: 'Product ID', 'Retailer', 'Quantities Sold', 'Total Sales'
# Check if 'Cost Price' exists before including it in the report
if 'Cost Price' in best_worst_sellers.columns:
    best_worst_sellers_report = best_worst_sellers[['Product ID', 'Retailer', 'Quantities Sold', 'Total Sales', 'Cost Price']]
else:
    best_worst_sellers_report = best_worst_sellers[['Product ID', 'Retailer', 'Quantities Sold', 'Total Sales']]

# Save the report to an Excel file
best_worst_sellers_report.to_excel('Best_Worst_Sellers_Report.xlsx', index=False)

print("Top 10 and Bottom 10 best-selling items report generated successfully!")
