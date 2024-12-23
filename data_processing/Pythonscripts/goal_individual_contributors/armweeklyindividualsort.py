import pandas as pd

# Load the source Excel file
source_file = "Armenian_Weekly_Articles.xlsx"
df = pd.read_excel(source_file)

# Exact matches to exclude
exact_exclusions = [
    "Guest Contributor",
    "Camp Haiastan",
    ".Wp-Block-Co-Authors-Plus-Coauthors.Is-Layout-Flow",
    "Weekly Staff",
]

# Keywords to exclude if present in 'Author'
keywords_exclusions = ["Armenian", "Anca", "Anc", "Rhode Island", "Eastern", "Western", "Homenetmen", "Hamazkayin", "Ayf", "Arf", "Activist", "Collective", "Committee", "Institute"]

# Remove rows with exact matches in 'Author'
df = df[~df['Author'].isin(exact_exclusions)]

# Remove rows where 'Author' contains any of the keywords
df = df[~df['Author'].str.contains('|'.join(keywords_exclusions), case=False, na=False)]

# List of columns to check for emptiness
columns_to_check = ['Title', 'Author', 'Publish Date', 'Keywords', 'URL']

# Replace blank strings with NaN
df[columns_to_check] = df[columns_to_check].replace("", pd.NA)

# Drop rows where any of the specified columns is NaN
filtered_df = df.dropna(subset=columns_to_check)

# Save the filtered data to a new Excel file
output_file = "Armenian_Weekly_IndividualContr.xlsx"
filtered_df.to_excel(output_file, index=False)

print(f"Filtered data saved to {output_file}")
