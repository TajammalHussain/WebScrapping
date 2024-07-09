# #import requests 
# #from bs4 import BeautifulSoup

# #page_to_scrape= requests.get("https://www.zoho.com/sites/")
# #soup= BeautifulSoup(page_to_scrape.text, "html.parser")
# #quotes= soup.findAll("h2", attrs={"class":"rtl-cent"})
# #print(quotes.text)
# #import requests
# #from bs4 import BeautifulSoup
# #import pprint

# #page_to_scrape = requests.get("https://www.zoho.com/sites/")
# #soup = BeautifulSoup(page_to_scrape.text, "html.parser")
# #quotes = soup.find_all("h2", class_="rtl-cent")

# # Extract and print the text content of each element
# #rtl_cent_values = [quote.text.strip() for quote in quotes]
# # #pprint.pprint(rtl_cent_values)
# # import numpy
# # arr = numpy.array([1, 2, 3, 4, 5])

# # print(arr)
# import fitz
# import os

# def extract_images_from_pdf(pdf_file):
#     doc = fitz.open(pdf_file)
#     images = []

#     for page_number in range(len(doc)):
#         page = doc.load_page(page_number)
#         image_list = page.get_images(full=True)

#         for img_index, img in enumerate(image_list):
#             xref = img[0]
#             base_image = doc.extract_image(xref)
#             image_bytes = base_image["image"]
#             images.append(image_bytes)

#     return images

# # Function to save images to the specified folder
# def save_images(images, folder_path):
#     for idx, image in enumerate(images):
#         file_path = os.path.join(folder_path, f"GeberitImage_{idx}.png")
#         with open(file_path, "wb") as f:
#             f.write(image)

# # Source PDF file path
# pdf_file = r"C:\Users\Tajammal\Desktop\MRF Brands upload\Geberit Drainage Price List - 01.05.2022.pdf"

# # Folder to save images
# output_folder = r"C:\Users\Tajammal\Desktop\MRF Brands upload\Geberit Images"

# # Extract images from PDF
# images = extract_images_from_pdf(pdf_file)

# # Save images to the specified folder
# save_images(images, output_folder)

import fitz
import os
from PIL import Image
import io

def extract_images_from_pdf(pdf_file):
    doc = fitz.open(pdf_file)
    images = []

    for page_number in range(len(doc)):
        page = doc.load_page(page_number)
        image_list = page.get_images(full=True)

        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            images.append(image_bytes)

    return images

# Function to optimize and save images to the specified folder
def save_images(images, folder_path, target_size=200):
    for idx, image_bytes in enumerate(images):
        # Open image using Pillow from image bytes
        img = Image.open(io.BytesIO(image_bytes))

        # Add alpha channel if the image doesn't already have one
        if img.mode != 'RGBA':
            img = img.convert('RGBA')

        # Save image with transparent background
        img.save(f'{folder_path}/TerrainImages_{idx}.png', optimize=True)

# Source PDF file path
pdf_file = r"C:\Users\Tajammal\Desktop\MRF Brands upload\Terrain_product_guide_march_2023.pdf"

# Folder to save images
output_folder = r"C:\Users\Tajammal\Desktop\MRF Brands upload\Terrain Images"

# Extract images from PDF
images = extract_images_from_pdf(pdf_file)

# Save optimized images to the specified folder
save_images(images, output_folder)
