import pandas as pd

# Load the source Excel file
source_file = "Horizon_Weekly_Articles.xlsx"
df = pd.read_excel(source_file)

# Remove rows where Author is NaN or "N/A"
filtered_df = df[df['Author'].notna()]  # Exclude rows with NaN in the Author column
filtered_df = filtered_df[filtered_df['Author'] != "N/A"]  # Exclude rows with Author as "N/A"

# Save the filtered data to a new Excel file
output_file = "Horizon_Weekly_IndividualContr.xlsx"
filtered_df.to_excel(output_file, index=False)

print(f"Filtered data saved to {output_file}")


