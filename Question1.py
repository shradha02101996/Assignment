import pandas as pd

# Load the data from the Excel file
file_path = 'DA_Task_Data.xlsx'
xls = pd.ExcelFile(file_path)

# Read the relevant sheets into DataFrames
contracts_df = pd.read_excel(xls, 'contracts')
areas_df = pd.read_excel(xls, 'areas')
owners_df = pd.read_excel(xls, 'owners')
tenants_df = pd.read_excel(xls, 'tenants')
units_df = pd.read_excel(xls, 'units')

# Convert date columns to datetime
contracts_df['contract_start_date'] = pd.to_datetime(contracts_df['contract_start_date'])
contracts_df['contract_end_date'] = pd.to_datetime(contracts_df['contract_end_date'])

# Filter contracts for the year 2022
contracts_2022_df = contracts_df[(contracts_df['contract_start_date'].dt.year == 2022) & (contracts_df['contract_end_date'].dt.year == 2022)]

# Merge contracts with units to get property details
merged_df = contracts_2022_df.merge(units_df, on='property_number', how='left')

# Merge with areas to get area names
merged_df = merged_df.merge(areas_df, on='area_number', how='left')

# Calculate contract duration in months
merged_df['contract_duration_months'] = ((merged_df['contract_end_date'] - merged_df['contract_start_date']) / pd.Timedelta(days=30)).astype(int)

# Calculate rent per sqm per month
merged_df['rent_per_sqm_per_month'] = merged_df['contract_amount'] / merged_df['contract_duration_months'] / merged_df['property_area']

# Extract month from contract start date
merged_df['month'] = merged_df['contract_start_date'].dt.month

# Select relevant columns for the final table
final_table = merged_df[['contract_amount', 'rent_per_sqm_per_month', 'contract_duration_months', 'property_type', 'property_sub_type', 'area_name', 'month']]

# Display the final table
print(final_table)
