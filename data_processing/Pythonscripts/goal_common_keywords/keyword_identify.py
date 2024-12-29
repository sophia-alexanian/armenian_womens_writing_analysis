import pandas as pd
from collections import Counter

# Function to filter and save common keywords with their count
def save_common_keywords(file_path, column_name, threshold=9):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Ensure the specified column exists
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in the Excel file.")

    # Extract the keywords
    keywords = df[column_name].dropna().astype(str)

    # Count keyword occurrences
    keyword_counts = Counter(keywords)

    # Filter keywords based on the threshold
    common_keywords = {key: count for key, count in keyword_counts.items() if count >= threshold}

    # Create a DataFrame for the common keywords
    result_df = pd.DataFrame(list(common_keywords.items()), columns=["Keyword", "Count"])

    # Save the results to a new Excel file
    result_df.to_excel("common_keywords.xlsx", index=False)
    print("Common keywords and their counts saved to 'common_keywords.xlsx'.")

# Example usage
if __name__ == "__main__":
    # File path to the Excel file
    excel_file_path = "Normalized_Armenian_Womens_Articles.xlsx"  # Replace with your file path
    
    # Column name containing the keywords
    keyword_column = "Keywords"  # Replace with your column name

    try:
        # Save the common keywords and their counts
        save_common_keywords(excel_file_path, keyword_column)
    except Exception as e:
        print(f"An error occurred: {e}")

