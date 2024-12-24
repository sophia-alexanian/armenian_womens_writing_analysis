import pandas as pd

# Load the spreadsheet
file_path = 'Armenian_Womens_Articles.xlsx'  # Replace with your file's path
sheet_name = 'Sheet1'  # Replace with your sheet name if needed

# Read the data
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Normalize the data by splitting the 'Keywords' column
df_normalized = df.assign(Keywords=df['Keywords'].str.split(', ')).explode('Keywords')

# Save the normalized data back to a new Excel file
output_file_path = 'Normalized_Armenian_Womens_Articles.xlsx'
df_normalized.to_excel(output_file_path, index=False)

print(f"Normalized data saved to {output_file_path}")
