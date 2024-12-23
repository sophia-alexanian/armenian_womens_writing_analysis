import pandas as pd

# Load Excel file
excel_file = "Armenian_Womens_Articles.xlsx"

# Convert to CSV
df = pd.read_excel(excel_file)
df.to_csv("Armenian_womens_articles.csv", index=False)
