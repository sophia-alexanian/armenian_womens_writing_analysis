import pandas as pd

# Load the source Excel file
source_file = "Horizon_Weekly_Articles.xlsx"
df = pd.read_excel(source_file)

# List of columns to check for emptiness
columns_to_check = ['Title', 'Author', 'Publish Date', 'Keywords', 'URL']

# Replace blank strings with NaN
df[columns_to_check] = df[columns_to_check].replace("", pd.NA)

# Drop rows where any of the specified columns is NaN
filtered_df = df.dropna(subset=columns_to_check)

# Save the filtered data to a new Excel file
output_file = "Horizon_Weekly_IndividualContr.xlsx"
filtered_df.to_excel(output_file, index=False)

print(f"Filtered data saved to {output_file}")
