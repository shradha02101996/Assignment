import pandas as pd

# Load the data from the Excel file
file_path = 'DA_Task_Data.xlsx'
xls = pd.ExcelFile(file_path)

# Read the relevant sheets into DataFrames
contracts_df = pd.read_excel(xls, 'contracts')
areas_df = pd.read_excel(xls, 'areas')
units_df = pd.read_excel(xls, 'units')

# Convert date columns to datetime
contracts_df['contract_start_date'] = pd.to_datetime(contracts_df['contract_start_date'])
contracts_df['contract_end_date'] = pd.to_datetime(contracts_df['contract_end_date'])

# Merge contracts with units to get property details
merged_df = contracts_df.merge(units_df, on='property_number', how='left')

# Merge with areas to get area names
merged_df = merged_df.merge(areas_df, left_on='area_number', right_on='area_number', how='left')

# Sort by tenant number and contract start date
merged_df = merged_df.sort_values(by=['tenant_number', 'contract_start_date'])

# Create a new column for the previous contract amount and area name
merged_df['prev_contract_amount'] = merged_df.groupby('tenant_number')['contract_amount'].shift(1)
merged_df['prev_area_name'] = merged_df.groupby('tenant_number')['area_name'].shift(1)

# Filter out rows where there is no previous contract
filtered_df = merged_df.dropna(subset=['prev_contract_amount', 'prev_area_name'])

# Determine the movement type
def determine_movement_type(row):
    if row['contract_amount'] > row['prev_contract_amount']:
        return 'Upgrade'
    elif row['contract_amount'] < row['prev_contract_amount']:
        return 'Downgrade'
    else:
        return 'Same Level'

filtered_df['movement_type'] = filtered_df.apply(determine_movement_type, axis=1)

# Select relevant columns for the final table
final_table = filtered_df[['tenant_number', 'prev_area_name', 'area_name', 'contract_start_date', 'movement_type']]
final_table.columns = ['Tenant Number', 'Area Moved From', 'Area Moved To', 'Moving Date', 'Movement Type']

# Display the final table
print(final_table)
