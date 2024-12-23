import pandas as pd

def filter_female_authors(input_files, output_file, author_gender_column):
    master_data = []

    # Iterate through each input file
    for file in input_files:
        try:
            # Read the data into a DataFrame
            data = pd.read_excel(file)

            # Filter rows where the author gender column indicates "Female"
            filtered_data = data[data[author_gender_column].str.lower() == 'female']

            # Append the filtered data to the master list
            master_data.append(filtered_data)
        except Exception as e:
            print(f"Error processing file {file}: {e}")

    # Combine all filtered data into a single DataFrame
    if master_data:
        master_df = pd.concat(master_data, ignore_index=True)

        # Save the combined DataFrame to the output file
        master_df.to_excel(output_file, index=False)
        print(f"Filtered data saved to {output_file}")
    else:
        print("No data to save.")

# List of input spreadsheet file paths
input_files = [
    'Armenian_Weekly_Gendered.xlsx',
    'EVN_Report_Gendered.xlsx',
    'Horizon_Weekly_Gendered.xlsx',
    'MirrorSpectator_Gendered.xlsx'
]

# Output file path for the master spreadsheet
output_file = 'Armenian_Womens_Articles.xlsx'

# Column name indicating author gender
author_gender_column = 'Gender'

# Call the function to filter and save the data
filter_female_authors(input_files, output_file, author_gender_column)
