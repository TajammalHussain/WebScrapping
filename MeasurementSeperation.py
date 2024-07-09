import pandas as pd
import re

# Function to split measurements
def split_measurements(measurement):
    # Split by 'x' or 'X' and strip extra spaces
    parts = re.split(r'[xX]', measurement)
    parts = [part.strip() for part in parts]
    return parts

# Path to the input Excel file
input_path = r"C:\Users\Tajammal\Desktop\MRF Brands upload\MSAccessMRF\Measurements.xlsx"
# Path to the output Excel file
output_path = r"C:\Users\Tajammal\Desktop\MRF Brands upload\MSAccessMRF\Measurements_Split.xlsx"

# Read the Excel file
df = pd.read_excel(input_path)

# Print the column names to verify
print(df.columns)

# Column name with measurements
measurements_column = 'Product Specifications'
measurements = df[measurements_column]

# Split the measurements into separate columns
split_data = measurements.apply(split_measurements)

# Create a new DataFrame from the split data
max_parts = split_data.apply(len).max()  # Find the maximum number of parts
split_df = pd.DataFrame(split_data.tolist(), columns=[f'Measurement {i+1}' for i in range(max_parts)])

# Combine the original DataFrame with the split DataFrame
result_df = pd.concat([df, split_df], axis=1)

# Save the result to a new Excel file
result_df.to_excel(output_path, index=False)

print(f"Excel file created successfully at {output_path}")
