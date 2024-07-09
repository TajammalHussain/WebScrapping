import pandas as pd

# Path to the Excel file
file_path = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Template.xlsx'

# Load data from the Excel file
try:
    df = pd.read_excel(file_path, sheet_name='Sheet1', engine='openpyxl')
except FileNotFoundError:
    print(f"Error: File '{file_path}' not found.")
    exit()
except Exception as e:
    print(f"Error occurred while reading the file: {e}")
    exit()

# Print column names to verify
print("Columns in DataFrame:")
print(df.columns)

# Assuming 'General specifications' is the column containing data
column_name = 'Details'

# Check if the column exists in the DataFrame
if column_name not in df.columns:
    print(f"Error: Column '{column_name}' not found in the Excel sheet.")
    exit()

# Function to convert string representation to dictionary
def str_to_dict(s):
    return eval(s)

# Define the desired order
desired_order = [
    {"Colour": "n/a", "Main Material": "Steel", "Material": "Steel", "Surface Protection": "Galvanic/electrolytic zinc plated"},
    {"Colour": "Black", "Main Material": "PVC", "Model": "1part", "Shape": "Straight"},
    {"Colour": "Grey", "Main Material": "PVC", "Model": "1part", "Shape": "Straight"},
    {"Colour": "White", "Main Material": "PVC", "Model": "1part", "Shape": "Straight"},
    {"Colour": "Black", "Main Material": "PVC", "Model": "2part", "Shape": "Straight"},
    {"Colour": "Grey", "Main Material": "PVC", "Model": "1part", "Shape": "Straight"},
    {"Colour": "Black", "Main Material": "PVCU", "Model": "1part", "Shape": "Straight"},
    {"Colour": "Grey", "Main Material": "PVCU", "Model": "2part", "Shape": "Straight"},
    {"Colour": "White", "Main Material": "PVCU", "Model": "2part", "Shape": "Straight"},
    {"Colour": "Grey", "Main Material": "PVCU", "Model": "2part", "Shape": "Straight"},
    {"Colour": "Black", "Main Material": "PVCU", "Model": "1part", "Shape": "Straight"},
    {"Colour": "Grey", "Main Material": "PVCU", "Model": "1part", "Shape": "Straight"}
]

# Sort the data based on the desired order
sorted_data = df.sort_values(by=column_name, key=lambda x: [desired_order.index(str_to_dict(item)) for item in x[column_name]])

# Export sorted data to a new Excel file
output_file_path = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Template Output.xlsx'
sorted_data.to_excel(output_file_path, index=False)

print(f"Exported sorted data to '{output_file_path}'")
