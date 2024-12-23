import pandas as pd

# Load the source Excel file
source_file = "The_Mirror_Spectator_Articles.xlsx"
df = pd.read_excel(source_file)

# Replace items in the 'Author' column longer than 5 words (this is where it scraped location info)
df['Author'] = df['Author'].apply(
    lambda x: "The Armenian Mirror-Spectator" if isinstance(x, str) and len(x.split()) > 5 else x
)

# Remove rows where 'Author' is "The Armenian Mirror-Spectator"
df = df[df['Author'] != "The Armenian Mirror-Spectator"]

# List of columns to check for emptiness
columns_to_check = ['Title', 'Author', 'Publish Date', 'Keywords', 'URL']

# Replace blank strings with NaN
df[columns_to_check] = df[columns_to_check].replace("", pd.NA)

# Drop rows where any of the specified columns is NaN
filtered_df = df.dropna(subset=columns_to_check)

# Save the modified data to a new Excel file
output_file = "MirrorSpectator_IndividualContr.xlsx"
df.to_excel(output_file, index=False)

print(f"Modified data saved to {output_file}")
