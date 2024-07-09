import pandas as pd

# Define the input and output file paths
input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Different Format images\WavinFull Description.xlsx'
output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Different Format images\WavinFull Description_Processed.xlsx'
sheet_name = 'Sheet1'

# Read the Excel file
df = pd.read_excel(input_file, sheet_name=sheet_name)

# Print the DataFrame columns to check for the 'Product specifications' column
print("DataFrame columns:", df.columns)

# Strip leading/trailing spaces from column names
df.columns = df.columns.str.strip()

# Check for 'Product specifications' column presence
if 'Product specifications' not in df.columns:
    raise KeyError("Column 'Product specifications' not found in the Excel file.")

# Function to split the description into specified sections
def split_description(description):
    if not isinstance(description, str):
        return {'Identification specifications': '', 'General specifications': '', 'Logistic specifications': '', 'Measurement specifications': ''}
    
    sections = ['General specifications', 'Logistic specifications', 'Measurement specifications']
    section_data = {'Identification specifications': ''}
    
    # Initialize the start position
    start_pos = 0
    for section in sections:
        # Find the start and end position for each section
        section_start = description.find(section, start_pos)
        if section_start == -1:
            continue
        section_data[sections[sections.index(section) - 1 if sections.index(section) > 0 else 0]] = description[start_pos:section_start].strip()
        start_pos = section_start
    
    # Add the final section
    section_data[sections[-1]] = description[start_pos:].strip()
    
    return section_data

# Apply the function to the DataFrame
split_data = df['Product specifications'].apply(split_description)

# Create a new DataFrame from the split data
split_df = pd.DataFrame(split_data.tolist())

# Combine the original DataFrame with the new split DataFrame
result_df = pd.concat([df, split_df], axis=1)

# Save the new DataFrame to an Excel file
result_df.to_excel(output_file, sheet_name=sheet_name, index=False)

print("Descriptions have been successfully split and saved in the Excel file.")
