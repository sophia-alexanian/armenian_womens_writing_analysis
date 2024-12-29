import pandas as pd
from collections import Counter

# Function to filter entries with the most common keywords
def filter_entries_with_common_keywords(file_path, column_name, threshold=10):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Ensure the specified column exists
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in the Excel file.")

    # Extract the keywords
    keywords = df[column_name].dropna().astype(str)

    # Count keyword occurrences
    keyword_counts = Counter(keywords)

    # Identify common keywords based on the threshold
    common_keywords = {key for key, count in keyword_counts.items() if count >= threshold}

    # Filter entries containing the common keywords
    filtered_entries = df[df[column_name].isin(common_keywords)]

    return filtered_entries

# Example usage
if __name__ == "__main__":
    # File path to the Excel file
    excel_file_path = "Normalized_Armenian_Womens_Articles.xlsx"  # Replace with your file path
    
    # Column name containing the keywords
    keyword_column = "Keywords"  # Replace with your column name

    try:
        # Get the filtered entries with common keywords
        filtered_entries = filter_entries_with_common_keywords(excel_file_path, keyword_column)

        # Print the results
        print(filtered_entries)

        # Save the results to a new Excel file
        filtered_entries.to_excel("Armenian_Women_Filtered_Entries.xlsx", index=False)
        print("Filtered entries saved to 'Armenian_Women_Filtered_Entries.xlsx'.")
    except Exception as e:
        print(f"An error occurred: {e}")
