import pandas as pd

# Load the Excel file
excel_file = 'data.xlsx'

# Read the Excel file
excel_data = pd.ExcelFile(excel_file)

# Loop through each sheet and save it as a separate CSV
for sheet_name in excel_data.sheet_names:
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    # Save the dataframe to a CSV file
    csv_filename = f"{sheet_name}.csv"
    df.to_csv(csv_filename, index=False)
    print(f"Sheet '{sheet_name}' saved as '{csv_filename}'.")

print("All sheets have been converted to CSV files.")
