import pandas as pd

def extract_column_to_csv(excel_file, column_name, output_csv):
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
        if column_name not in df.columns:
            print(f"Column '{column_name}' not found in the Excel file.")
            return

        column_data = df[[column_name]]

        column_data.to_csv(output_csv, index=False)

        print(f"Column '{column_name}' has been successfully written to '{output_csv}'.")
    except FileNotFoundError:
        print(f"The file '{excel_file}' does not exist.")
    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":

    excel_file_path = "Normalized_Armenian_Womens_Articles.xlsx"  

    
    column_name_to_extract = "Keywords"  

    
    output_csv_path = "totalwomenswriting.csv"  

    
    extract_column_to_csv(excel_file_path, column_name_to_extract, output_csv_path)
