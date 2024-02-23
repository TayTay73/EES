import pandas as pd
import numpy as np
import os

# This section is for creating a file name
def get_valid_year(prompt="What is the Year(YYYY)? "):
    while True:
        try:
            year = int(input(prompt))
            if year >= 2000 and year <= 2099:
                return year
            else:
                print("Invalid year. Please enter a 4-digit year.")
        except ValueError:
            print("Invalid input. Please enter a number.")

def get_valid_month(prompt="What month(MMM) is this for? "):
    valid_months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    while True:
        month = input(prompt).capitalize()
        if month in valid_months:
            return month
        else:
            print("Invalid month. Please enter a valid month name.")

file_year = get_valid_year()
file_month = get_valid_month()

custom_name = f"{file_year} {file_month}"
filename = f'Cost Adjustment {custom_name}.xlsx'

def get_valid_file_path(prompt):
    while True:
        path = input(prompt)
        if os.path.exists(path) and os.path.isfile(path):
            return path
        else:
            print("Invalid file path. Please enter a valid path to an existing file.")

# Load last month's data
last_month_file_path = get_valid_file_path("Enter the full path to last month's Excel file: ")
try:
    last_month_data = pd.read_excel(last_month_file_path)
    print("Last month's Excel file read successfully!")
    print(last_month_data.head())
except (FileNotFoundError, pd.errors.ParserError, PermissionError) as e:
    print(f"Error reading last month's file: {e}")

# Load new data
new_data_file_path = get_valid_file_path("Enter the full path to the new Excel file: ")
try:
    new_data = pd.read_excel(new_data_file_path)
    print("New Excel file read successfully!")
    print(new_data.head())
except (FileNotFoundError, pd.errors.ParserError, PermissionError) as e:
    print(f"Error reading new data file: {e}")


# Filter rows based on Employee ID presence in new_data
rows_to_delete = last_month_data[~last_month_data['Employee ID'].isin(new_data['Employee ID'])]
last_month_data.drop(rows_to_delete.index, inplace=True)  # Drop rows and reset index

# Alternatively, using query:
# last_month = last_month.query("Employee ID not in @new_data['Employee ID ']")

# Filter new IDs not present in last_month
new_ids = new_data[~new_data['Employee ID'].isin(last_month_data['Employee ID'])]

# Ensure columns match between DataFrames before combining
if last_month_data.columns.tolist() != new_ids.columns.tolist():
    print("Warning: Column names differ between DataFrames. Make sure they match before combining.")
    # Optionally attempt to align columns or raise an error

# Combine last_month and new_ids, potentially appending or concatenating:
combined_df = pd.concat([last_month_data, new_ids], ignore_index=True)

combined_df = combined_df.fillna(0)

# Re-arranges columns for last_month locking the first row
combined_df['Store Total Last Month'] = combined_df['Store Total']
combined_df['SS Last Month'] = combined_df['SS Total']
combined_df['2023 Training Costs Last Month'] = combined_df['Total Training Costs 2023']

# Move specific columns from new data to old data
combined_df['Overtime Hours'] = new_data['Overtime Hours']
combined_df['Regular Hours'] = new_data['Regular Hours']
combined_df['Regular Pay'] = new_data['Regular Pay']
combined_df['Overtime Pay'] = new_data['Overtime Pay']
combined_df['Commission Pay'] = new_data['Commission Pay']
combined_df['Gross Pay'] = new_data['Gross Pay']

# Recalculate formulas (replace with your specific formula logic)
combined_df['Total Hours'] = combined_df['Regular Hours'] + combined_df['Overtime Hours']
combined_df['Rate + Comm'] = (combined_df['Commission Pay'] / combined_df['Total Hours']) + combined_df['Hourly Rate']
combined_df['Gross Pay *1.2'] = combined_df['Gross Pay'] * 1.2
combined_df['Store This Month'] = combined_df['Gross Pay *1.2'] - combined_df['SS This Month']
combined_df['Store Total'] = combined_df['Store Total Last Month'] + combined_df['Store This Month']
combined_df['SS Total'] = combined_df['SS Last Month'] + combined_df['SS This Month']
combined_df['Total Training Costs'] = combined_df['Store Total'] + combined_df['SS Total']
combined_df['Total Training Costs 2023'] = combined_df['2023 Training Costs Last Month'] + combined_df['Gross Pay *1.2']

# Output  Excel
combined_df.to_excel(filename, index=False)