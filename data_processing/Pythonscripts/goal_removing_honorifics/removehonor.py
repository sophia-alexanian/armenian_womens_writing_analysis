import pandas as pd
import re

# Define a function to remove honorifics
def remove_honorific(name):
    # Define a list of common honorifics (allow for multiple honorifics in sequence)
    honorifics = r'\b(Mr\.|Ms\.|Dr\.|Prof\.|Mrs\.|Rev\.|Fr\.|His Holiness Catholicos\.|Very\.|Archpriest\.)\b'
    
    # Use re.sub to replace all honorifics in the name, possibly with multiple honorifics before the name
    cleaned_name = re.sub(r'(\b(Mr\.|Ms\.|Dr\.|Prof\.|Mrs\.|Rev\.|Fr\.|His Holiness Catholicos\.|Very\.|Archpriest\.)\s*)+', '', name).strip()

    return cleaned_name

# Read the Excel file
file_path = 'MirrorSpectator_IndividualContr.xlsx'  # Replace with your file path
df = pd.read_excel(file_path)

# Apply the function to remove honorifics from the 'Author' column
df['Author'] = df['Author'].apply(remove_honorific)

# Save the modified DataFrame to a new Excel file
df.to_excel('MirrorSpectator_IndividualContr2_0.xlsx', index=False)

print("Honorifics removed and file saved as 'MirrorSpectator_IndividualContr2_0.xlsx'.")


