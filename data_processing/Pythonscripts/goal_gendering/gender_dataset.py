import pandas as pd

def create_gender_dictionary(input_file):
    df = pd.read_excel(input_file)
    name_gender_dict = dict(zip(df['First Name'], df['Gender']))
    return name_gender_dict

def assign_genders(input_files, output_files, name_gender_dict):
    for input_file, output_file in zip(input_files, output_files):
        df = pd.read_excel(input_file)
        df['Gender'] = df['First Name'].map(name_gender_dict).fillna('Unknown')
        df.to_excel(output_file, index=False)

def main():
    # Input spreadsheet with names and genders
    input_dictionary_file = 'unique_gendered_author_names.xlsx'

    # Input files for first names
    input_files = [
        'Horizon_Weekly_IndividualContr.xlsx',
        'Armenian_Weekly_IndividualContr2_0.xlsx',
        'MirrorSpectator_IndividualContr2_0.xlsx',
        'EVN_Report_IndividualContr2_0.xlsx'
    ]

    # Output files for spreadsheets with assigned genders
    output_files = [
        'Horizon_Weekly_Gendered.xlsx',
        'Armenian_Weekly_Gendered.xlsx',
        'MirrorSpectator_Gendered.xlsx',
        'EVN_Report_Gendered.xlsx'
    ]

    # Create the gender dictionary
    name_gender_dict = create_gender_dictionary(input_dictionary_file)

    # Assign genders to the names in spreadsheets
    assign_genders(input_files, output_files, name_gender_dict)


main()
