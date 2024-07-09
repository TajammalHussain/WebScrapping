import os
import shutil
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

# Path to the input and output Excel files
input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\Codds.xlsx'
output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\PDFFF.xlsx'
sheet_name = 'Sheet1'

# Load the Excel file
df = pd.read_excel(input_file)

# Print the DataFrame to inspect column names
print("DataFrame columns:", df.columns)

# Ensure 'Product Code' column is correctly identified (handle potential leading/trailing spaces)
df.columns = df.columns.str.strip()

# Check for 'Product Code' column presence
if 'Product Code' not in df.columns:
    raise KeyError("Column 'Product Code' not found in the Excel file.")

# Create a new folder to save downloaded files
download_folder = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\DownloadPDFF'
if not os.path.exists(download_folder):
    os.makedirs(download_folder)

# Configure Chrome options to set the default download directory
chrome_options = Options()
prefs = {'download.default_directory': download_folder}
chrome_options.add_experimental_option('prefs', prefs)

# Initialize the web driver with options
driver = webdriver.Chrome(options=chrome_options)

# Function to search for the product link and description based on the product code
def find_product_info(product_code):
    search_url = "https://aalberts-ips.co.uk/"
    driver.get(search_url)
    
    try:
        search_box = driver.find_element(By.ID, 'input_3_2')
    except Exception as e:
        print(f"Search box not found for product code {product_code}: {e}")
        return 'Search Box Not Found', 'Description Not Found'
    
    search_box.send_keys(product_code)
    search_box.send_keys(Keys.RETURN)
    
    time.sleep(1)  # Wait for the page to load (adjust as needed)
    
    try:
        # Find the product link element
        product_link_element = driver.find_element(By.CSS_SELECTOR, 'div.product-card__info a')
        product_link = product_link_element.get_attribute('href')
        
        # Navigate to the product link
        driver.get(product_link)
        
        time.sleep(1)  # Wait for the page to load (adjust as needed)
        
        # Find the description element on the product page
        description_element = driver.find_element(By.CSS_SELECTOR, 'div.product-detail--info--content')
        description_text = description_element.text

        # Find and download the datasheet
        datasheet_link_element = driver.find_element(By.ID, 'pdf-button')
        datasheet_link_element.click()
        
        time.sleep(3)  # Wait for the file to download (adjust as needed)

        # Wait for the download to complete and ensure the file is moved to the target folder
        download_successful = False
        for _ in range(20):  # Adjust number of iterations and sleep time as needed
            time.sleep(1)
            for filename in os.listdir(download_folder):
                if filename.endswith('.pdf'):  # Assuming the datasheet is a PDF file
                    download_successful = True
                    break
            if download_successful:
                break
        
        if not download_successful:
            print(f"Datasheet download failed for product code {product_code}")
    except Exception as e:
        print(f"Description or datasheet not found for product code {product_code}: {e}")
        product_link = driver.current_url
        description_text = 'Description Not Found'
    
    return product_link, description_text

# Create new columns in the DataFrame for product links and descriptions
df['Product Link'] = ''
df['Description'] = ''

# Iterate over each product code and update the DataFrame and Excel file
for index, row in df.iterrows():
    product_code = row['Product Code']
    product_link, description = find_product_info(product_code)
    
    df.at[index, 'Product Link'] = product_link
    df.at[index, 'Description'] = description
    
    # Save the updated DataFrame to a new Excel file after each iteration
    df.to_excel(output_file, sheet_name=sheet_name, index=False)

# Close the web driver
driver.quit()

print("Product links, descriptions, and datasheets have been successfully updated and downloaded.")
