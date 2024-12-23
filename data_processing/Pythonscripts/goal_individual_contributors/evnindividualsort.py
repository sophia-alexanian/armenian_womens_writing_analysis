import pandas as pd

# Load the source Excel file
source_file = "EVN_Report_Articles.xlsx"
df = pd.read_excel(source_file)

# Clean up column values to remove leading/trailing spaces or line breaks
df['Author'] = df['Author'].str.strip()

# Remove rows where Author is "Evn Report" or matches a specific unwanted string
df = df[~df['Author'].isin(["Evn Report", ".Wp-Block-Co-Authors-Plus-Coauthors.Is-Layout-Flow"])]

# List of columns to check for emptiness
columns_to_check = ['Title', 'Author', 'Publish Date', 'Keywords', 'URL']

# Replace blank strings with NaN
df[columns_to_check] = df[columns_to_check].replace("", pd.NA)

# Drop rows where any of the specified columns is NaN
filtered_df = df.dropna(subset=columns_to_check)

# Save the filtered data to a new Excel file
output_file = "EVN_Report_IndividualContr.xlsx"
filtered_df.to_excel(output_file, index=False)

print(f"Filtered data saved to {output_file}")




