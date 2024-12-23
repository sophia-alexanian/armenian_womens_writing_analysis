import pandas as pd
import os

# List of Excel files you want to process
excel_files = ["Armenian_Weekly_IndividualContr2_0.xlsx", "EVN_Report_IndividualContr2_0.xlsx", "MirrorSpectator_IndividualContr2_0.xlsx", "Horizon_Weekly_IndividualContr.xlsx"]

# Column name to extract (replace with the actual column name)
column_name = 'First Name'

# Set to hold unique values
unique_values = set()

# Loop through each Excel file
for file in excel_files:
    # Check if the file exists
    if os.path.exists(file):
        # Read the Excel file
        df = pd.read_excel(file)

        # Check if the column exists in the file
        if column_name in df.columns:
            # Add unique values from the column to the set
            unique_values.update(df[column_name].dropna().unique())
        else:
            print(f"Column '{column_name}' not found in {file}")
    else:
        print(f"File {file} does not exist")

# Convert the set of unique values to a DataFrame
unique_df = pd.DataFrame(list(unique_values), columns=[column_name])

# Save the unique values to a new Excel file
unique_df.to_excel("unique_author_names.xlsx", index=False)

print("Unique values have been saved to 'unique_author_names.xlsx'")
