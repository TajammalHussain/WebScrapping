import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
from PIL import Image as PILImage

# Define the path to the directory containing the images
image_dir = r"C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\DownloadedFiles\pngImages"

# Get a list of image filenames
image_files = [f for f in os.listdir(image_dir) if f.endswith('.png')]

# Create a DataFrame with the filenames
df = pd.DataFrame(image_files, columns=['Image Name'])

# Create a new Excel workbook and select the active worksheet
wb = Workbook()
ws = wb.active
ws.title = "Images"

# Write the DataFrame to the worksheet
for r in dataframe_to_rows(df, index=False, header=True):
    ws.append(r)

# Add images to the worksheet
for idx, image_file in enumerate(image_files, start=2):  # Start from row 2 to account for the header
    try:
        # Open the image using PIL to ensure it's not corrupted
        pil_img = PILImage.open(os.path.join(image_dir, image_file))
        pil_img.verify()  # Verify that it's a valid image
        
        # Load the image using openpyxl
        img = Image(os.path.join(image_dir, image_file))
        img.width = 100  # Adjust the image width if needed
        img.height = 100  # Adjust the image height if needed
        cell_address = f'B{idx}'
        ws.add_image(img, cell_address)
        
    except (IOError, OSError) as e:
        print(f"Error opening image {image_file}: {e}")

# Adjust column widths
ws.column_dimensions['A'].width = 30
ws.column_dimensions['B'].width = 30

# Save the workbook
output_path = r"C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\DownloadedFiles\ImageList.xlsx"
try:
    wb.save(output_path)
    print(f"Excel file created successfully at {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")
