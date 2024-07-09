import pandas as pd
import os
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
import io
import time

# Path to the input and output Excel files
input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Wavin images\Codes.xlsx'
output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Wavin images\WavinTextSecond.xlsx'
image_dir = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Wavin images\DownloadedImages'
sheet_name = 'Sheet1'

# Load the Excel file
df = pd.read_excel(input_file, sheet_name=sheet_name)

# Ensure 'Product Code' column is correctly identified (handle potential leading/trailing spaces)
df.columns = df.columns.str.strip()

# Check for 'Product Code' column presence
if 'Product Code' not in df.columns:
    raise KeyError("Column 'Product Code' not found in the Excel file.")

# Convert 'Product Code' to string
df['Product Code'] = df['Product Code'].astype(str)

# Create directory for downloaded images if it doesn't exist
if not os.path.exists(image_dir):
    os.makedirs(image_dir)

# Initialize the web driver
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

# Function to search for the product link based on the product code
def find_product_link(product_code):
    search_url = "https://www.wavin.com/en-gb/search"
    driver.get(search_url)
    
    try:
        # Wait until the search box is present and find it
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="search"]'))
        )
        search_box.send_keys(product_code)
        search_box.send_keys(Keys.RETURN)
        
        # Wait for the search results to load
        time.sleep(1)
        
        # Find the first search result link
        first_result_link = driver.find_element(By.CSS_SELECTOR, 'a[data-qa="elastic-search-item_link"]')
        product_page_url = first_result_link.get_attribute('href')
        return product_page_url
    except Exception as e:
        print(f"Error finding product link for {product_code}: {e}")
        return 'Link Not Found'

# Function to extract additional information from the product page
def extract_product_info(url):
    driver.get(url)
    
    try:
        # Wait until the product title and image are present
        product_title = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'h1[data-qa="product-detail_title"]'))
        ).text.strip()
        
        product_image_url = driver.find_element(By.CSS_SELECTOR, 'img[data-qa="product-image_image"]').get_attribute('src')
        
        # Wait until the collapsible triggers are present
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.collapsible-trigger_lKDGc'))
        )
        
        # Find all collapsible triggers and click to expand them if not already expanded
        collapsible_triggers = driver.find_elements(By.CSS_SELECTOR, '.collapsible-trigger_lKDGc')
        for trigger in collapsible_triggers:
            try:
                trigger_button = trigger.find_element(By.XPATH, './button')
                if trigger_button.getAttribute('aria-expanded') == 'false':
                    trigger.click()
                    time.sleep(2)  # Give some time for the content to expand
            except Exception as e:
                print(f"Error clicking trigger: {e}")
                continue
        
        # Extract the text within <dt> and <dd> tags with specific data-qa attribute
        dt_elements = driver.find_elements(By.CSS_SELECTOR, 'dt[data-qa="product-specifications_spec-label"]')
        dd_elements = driver.find_elements(By.CSS_SELECTOR, 'dd[data-qa="product-specifications_spec-value"]')
        
        product_info = ""
        
        for dt, dd in zip(dt_elements, dd_elements):
            dt_text = dt.text.strip()
            dd_text = dd.text.strip()
            if dt_text and dd_text:
                product_info += f"{dt_text}: {dd_text}\n"
        
        return product_title, product_image_url, product_info.strip()
    
    except Exception as e:
        print(f"Error extracting product info from {url}: {e}")
        return 'Title Not Found', 'Image Not Found', 'Info Not Found'

# Function to download an image from a URL and convert it to PNG
def download_and_convert_image(url, save_path):
    try:
        response = requests.get(url)
        response.raise_for_status()

        # Convert the image to PNG format
        image = Image.open(io.BytesIO(response.content))
        image = image.convert('RGBA')  # Ensure image has an alpha channel
        image.save(save_path, 'PNG')
        
        return save_path
    except Exception as e:
        print(f"Error downloading or converting image from {url}: {e}")
        return 'Image Download Failed'

# Iterate over each row, find the product link, extract additional info, download image, and save after each row
for index, row in df.iterrows():
    product_code = row['Product Code']
    product_link = find_product_link(product_code)
    if product_link != 'Link Not Found':
        product_title, product_image_url, additional_info = extract_product_info(product_link)
    else:
        product_title, product_image_url, additional_info = 'Title Not Found', 'Image Not Found', 'Info Not Found'
    
    df.at[index, 'Product Link'] = product_link
    df.at[index, 'Product Title'] = product_title
    df.at[index, 'Additional Info'] = additional_info
    
    # Replace "/" with "_" in the product code for saving the image
    sanitized_product_code = product_code.replace('/', '_')
    if product_image_url != 'Image Not Found':
        image_path = os.path.join(image_dir, f"{sanitized_product_code}.png")
        download_and_convert_image(product_image_url, image_path)
        df.at[index, 'Image Path'] = image_path
    else:
        df.at[index, 'Image Path'] = 'Image Not Found'
    
    # Save the DataFrame after each update
    df.to_excel(output_file, sheet_name=sheet_name, index=False)
    print(f"Updated record for product code: {product_code}")

# Close the web driver
driver.quit()

print("Product links, additional information, and images have been successfully updated in the Excel file.")
