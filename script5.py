import pandas as pd

# Load CSVs into DataFrames
areas_df = pd.read_csv('areas.csv')
contracts_df = pd.read_csv('cleaned_contracts.csv')
units_df = pd.read_csv('units.csv')

# Convert date columns to datetime
contracts_df['start_date'] = pd.to_datetime(contracts_df['contract_start_date'])
contracts_df['end_date'] = pd.to_datetime(contracts_df['contract_end_date'])
contracts_df['contract_status'] = contracts_df['contract_status'].str.strip().str.lower()

# Filter for contracts active in 2022
active_contracts_2022 = contracts_df[
    (contracts_df['start_date'] <= '2022-12-31') &
    (contracts_df['end_date'] >= '2022-01-01') &
    (contracts_df['contract_status'].isin(['active', 'renewed']))
]

# Calculate months in 2022
def calculate_months_in_2022(start_date, end_date):
    start = max(pd.to_datetime(start_date), pd.to_datetime('2022-01-01'))
    end = min(pd.to_datetime(end_date), pd.to_datetime('2022-12-31'))
    if start > end:
        return 0
    return (end.year - start.year) * 12 + end.month - start.month + 1

active_contracts_2022.loc[:, 'months_in_2022'] = active_contracts_2022.apply(
    lambda row: calculate_months_in_2022(row['start_date'], row['end_date']),
    axis=1
)

# Join with areas
merged_df = active_contracts_2022.merge(areas_df, on='area_number')

# Join with units
merged_df = merged_df.merge(units_df, on='property_number')

# Check columns and create any missing columns if necessary
print("Merged DataFrame columns:", merged_df.columns)

# If needed, calculate contract_duration_months
if 'contract_duration_months' not in merged_df.columns:
    merged_df['contract_duration_months'] = (
        (merged_df['end_date'] - merged_df['start_date']).dt.days // 30
    )

# Calculate 2022 Rental Amount
merged_df['2022_Rental_Amount'] = merged_df['rent_per_month'] * merged_df['months_in_2022']
merged_df['Rent_per_sqm'] = merged_df['rent_per_month'] / merged_df['property_area']

# Select relevant columns
result_df = merged_df[['2022_Rental_Amount', 'Rent_per_sqm', 'contract_duration_months', 'property_type', 'property_sub_type', 'area_name', 'start_date']]

# Rename columns as needed
result_df.columns = ['2022 Rental Amount', 'Rents per sqm', 'Contract Duration', 'Property Type', 'Property Sub Type', 'Area Name', 'Month']

# Save or display the result
result_df.to_csv('result_2022_rental_amount.csv', index=False)
print(result_df)
