import pandas as pd

# Define the path to your Excel file
excel_file_path = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Terrain Images\ImagesCodes.xlsx'

# Read the Excel file into a DataFrame
df = pd.read_excel(excel_file_path)

# Identify duplicate product codes
duplicate_product_codes = df[df.duplicated(subset=['Product Code'], keep=False)]['Product Code'].unique()

# Print duplicate product codes
print("Duplicate Product Codes:")
print(duplicate_product_codes)

# Add a new column to mark duplicates
df['Is Duplicate'] = df['Product Code'].isin(duplicate_product_codes)

# Display the DataFrame with the 'Is Duplicate' column highlighted
print("\nDataFrame with Duplicates Marked:")
print(df)

# Optionally, you can save the modified DataFrame back to a new Excel file
# df.to_excel('HighlightedDuplicates.xlsx', index=False)
# in this program we can clear numbers which are repeated as codes in the excel source sheet, so that means the numbers are entered twice or may be in different locations in the catalog with same products codes
#however, they are unique, but are repeated in the catalog, thats why they are in different pages under may be different headings or columns. 