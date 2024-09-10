import pandas as pd
from datetime import datetime

# Load the cleaned data
df = pd.read_csv('cleaned_contracts.csv')

def calculate_rent_in_year(start_date, end_date, rent_per_month, year):
    """
    Calculate the rent earned during a specific year for a given contract.
    """
    start = datetime.strptime(start_date, '%Y-%m-%d')
    end = datetime.strptime(end_date, '%Y-%m-%d')
    
    # Define the start and end of the specified year
    year_start = datetime(year, 1, 1)
    year_end = datetime(year, 12, 31)
    
    # Determine the actual start and end dates for the overlap with the specified year
    actual_start = max(start, year_start)
    actual_end = min(end, year_end)
    
    # Check if there's an overlap
    if actual_start <= actual_end:
        # Calculate the number of months in the overlap
        months = (actual_end.year - actual_start.year) * 12 + actual_end.month - actual_start.month + 1
        return months * rent_per_month, months
    else:
        return 0, 0

# Initialize counters
total_months = 0
total_rent = 0

# Prepare a log for detailed information
log_details = []

# Loop through each contract and calculate the rent earned in 2022
for _, row in df.iterrows():
    contract_number = row['contract_number']
    tenant_number = row['tenant_number']
    start_date = row['contract_start_date']
    end_date = row['contract_end_date'] if not pd.isna(row['contract_end_date']) else '2022-12-31'
    rent_per_month = row['rent_per_month']
    
    # Calculate the rent earned in 2022
    rent_earned, months_in_2022 = calculate_rent_in_year(start_date, end_date, rent_per_month, 2022)
    
    # Log the details of each contract's calculation
    log_details.append({
        'contract_number': contract_number,
        'tenant_number': tenant_number,
        'start_date': start_date,
        'end_date': end_date,
        'rent_per_month': rent_per_month,
        'months_in_2022': months_in_2022,
        'rent_earned_in_2022': rent_earned
    })
    
    # Print details for this contract
    print(f"Contract Number: {contract_number}")
    print(f"  Tenant Number: {tenant_number}")
    print(f"  Start Date: {start_date}")
    print(f"  End Date: {end_date}")
    print(f"  Rent per Month: {rent_per_month}")
    print(f"  Months in 2022: {months_in_2022}")
    print(f"  Rent Earned in 2022: {rent_earned:.2f}\n")
    
    # Update total counters
    if rent_earned > 0:
        total_months += months_in_2022
        total_rent += rent_earned

# Print the overall results
print(f"Total months covered in 2022: {int(total_months)}")
print(f"Total rent earned in 2022: {total_rent:.2f}")

# Save log details to a CSV file
log_df = pd.DataFrame(log_details)
log_df.to_csv('rent_earned_2022_log.csv', index=False)

print("Detailed log of rent earned in 2022 saved to 'rent_earned_2022_log.csv'.")
