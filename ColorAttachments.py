import pandas as pd

# File paths
input_file_path = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Terrain Images\Codess.xlsx'
output_file_path = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Terrain Images\Codess_Expanded.xlsx'

# Load data from Excel
df = pd.read_excel(input_file_path)

# Ensure 'Colors' column is treated as string and handle NaN
df['Colors'] = df['Colors'].astype(str).apply(lambda x: x.split(', ') if x != 'nan' else [])

# Create a list to store expanded rows
expanded_rows = []

# Iterate over each row and expand based on colors
for index, row in df.iterrows():
    if row['Colors']:  # If there are colors specified
        for color in row['Colors']:
            expanded_rows.append({
                'Product Code': f"{row['Product Code']}{color[0]}",
                'Color': color,
                'Original Product Code': row['Product Code'],
                'Order': index
            })
    else:  # If there are no colors specified, retain the original row
        expanded_rows.append({
            'Product Code': row['Product Code'],
            'Color': '',
            'Original Product Code': row['Product Code'],
            'Order': index
        })

# Create a new DataFrame with expanded rows
expanded_df = pd.DataFrame(expanded_rows)

# Sort by the original order based on the 'Order' column
expanded_df.sort_values(by='Order', inplace=True)

# Drop the 'Order' and 'Original Product Code' columns
expanded_df.drop(columns=['Order', 'Original Product Code'], inplace=True)

# Save to Excel
expanded_df.to_excel(output_file_path, index=False)

print(f"Expanded data saved to {output_file_path}")
