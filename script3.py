import pandas as pd
from datetime import datetime

# Load the cleaned data
df = pd.read_csv('cleaned_contracts.csv')

def calculate_rent_in_year(start_date, end_date, rent_per_month, year):
    # Calculate months in a specific year
    start = datetime.strptime(start_date, '%Y-%m-%d')
    end = datetime.strptime(end_date, '%Y-%m-%d')
    
    # Calculate the range of months for the specified year
    year_start = datetime(year, 1, 1)
    year_end = datetime(year, 12, 31)
    
    # Calculate overlapping period
    actual_start = max(start, year_start)
    actual_end = min(end, year_end)
    
    if actual_start <= actual_end:
        # Calculate the number of months in the overlap
        months = (actual_end.year - actual_start.year) * 12 + actual_end.month - actual_start.month + 1
        return months * rent_per_month
    else:
        return 0

# Initialize counters
total_months = 0
total_rent = 0

# Loop through each contract and calculate the rent earned in 2022
for _, row in df.iterrows():
    if pd.isna(row['contract_end_date']):
        # If no end date, assume the contract was active at the end of 2022
        end_date = '2022-12-31'
    else:
        end_date = row['contract_end_date']
    
    # Check if the contract overlaps with 2022
    if datetime.strptime(row['contract_start_date'], '%Y-%m-%d').year <= 2022:
        months_in_2022 = calculate_rent_in_year(row['contract_start_date'], end_date, row['rent_per_month'], 2022)
        if months_in_2022 > 0:
            total_months += months_in_2022 / row['rent_per_month']
            total_rent += months_in_2022

print(f"Total months in 2022: {int(total_months)}")
print(f"Total rent earned in 2022: {total_rent:.2f}")
