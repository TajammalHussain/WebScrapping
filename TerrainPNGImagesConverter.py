import os
import pandas as pd
import requests
from PIL import Image
from io import BytesIO

# Define file paths and output directory
excel_file_path = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Terrain Images\ImagesCodes.xlsx'
output_directory = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Terrain Images\PNGImages'

# Create output directory if it doesn't exist
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Read the Excel file
df = pd.read_excel(excel_file_path)

# Initialize counters
saved_count = 0
not_saved_count = 0

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    product_code = row['Product Code']
    image_url = row['Image URL']
    
    # Check if image_url is NaN or None
    if pd.isna(image_url) or not image_url:
        print(f"Image URL is missing for product code {product_code}. Skipping...")
        not_saved_count += 1
        continue
    
    try:
        # Send a request to the image URL
        response = requests.get(image_url)
        response.raise_for_status()
        
        # Open the image from the response content
        image = Image.open(BytesIO(response.content))
        
        # Define the output file path
        output_file_path = os.path.join(output_directory, f"{product_code}.png")
        
        # Save the image in PNG format
        image.save(output_file_path, format='PNG')
        
        print(f"Downloaded and saved {product_code}.png")
        saved_count += 1
        
        # Additional debug output
        print(f"Current saved count: {saved_count}")
    
    except requests.RequestException as e:
        print(f"Failed to download image for product code {product_code}: {e}")
        not_saved_count += 1
    except Exception as e:
        print(f"Failed to save image for product code {product_code}: {e}")
        not_saved_count += 1

# Print summary
print("\n--- Summary ---")
print(f"Total images processed: {len(df)}")
print(f"Images saved successfully: {saved_count}")
print(f"Images not saved (missing URL or error): {not_saved_count}")

# Count PNG files in the directory after the script completes
png_files = [f for f in os.listdir(output_directory) if f.endswith('.png')]
print(f"Actual number of PNG files in '{output_directory}': {len(png_files)}")

print("\nImage download and save process completed.")
