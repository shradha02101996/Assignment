import pandas as pd
import re
from datetime import datetime

# Load the data
df = pd.read_csv('contracts.csv')

def clean_amount(amount):
    # Remove currency symbols and commas, then convert to integer
    amount = amount.strip().upper()
    amount = re.sub(r'[^\d]', '', amount)
    return int(amount)

def calculate_months(start_date, end_date):
    # Calculate the number of months between two dates
    start = datetime.strptime(start_date, '%Y-%m-%d')
    end = datetime.strptime(end_date, '%Y-%m-%d')
    return (end.year - start.year) * 12 + end.month - start.month

# Clean and convert contract_amount
df['contract_amount'] = df['contract_amount'].apply(clean_amount)

# Convert dates and calculate duration in months
df['contract_start_date'] = pd.to_datetime(df['contract_start_date'])
df['contract_end_date'] = pd.to_datetime(df['contract_end_date'])
df['duration_months'] = df.apply(lambda row: calculate_months(row['contract_start_date'].strftime('%Y-%m-%d'), row['contract_end_date'].strftime('%Y-%m-%d')), axis=1)

# Calculate rent per month
df['rent_per_month'] = df['contract_amount'] / df['duration_months']

# Ensure all required columns are included in the final CSV
columns_to_save = [
    'contract_number', 'tenant_number', 'contract_start_date', 'contract_end_date',
    'property_number', 'contract_amount', 'contract_status', 'contract_termination_date',
    'owner_number', 'area_number', 'duration_months', 'rent_per_month'
]

# Save the complete data to CSV
df.to_csv('cleaned_contracts.csv', columns=columns_to_save, index=False)

# Aggregate total rent per month for each tenant
total_rent_per_tenant = df.groupby('tenant_number')['rent_per_month'].sum().reset_index()

# Save the aggregated data to CSV
total_rent_per_tenant.to_csv('total_rent_per_tenant.csv', index=False)

print("Data cleaned, analyzed, and saved to 'cleaned_contracts.csv' and 'total_rent_per_tenant.csv'.")
