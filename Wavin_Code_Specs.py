import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

# Path to the input and output Excel files
input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Wavin images\Codes.xlsx'
output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Wavin images\Codes_Links.xlsx'
sheet_name = 'Sheet1'

# Load the Excel file
df = pd.read_excel(input_file, sheet_name=sheet_name)

# Print the DataFrame to inspect column names
print("DataFrame columns:", df.columns)

# Ensure 'Product Code' column is correctly identified (handle potential leading/trailing spaces)
df.columns = df.columns.str.strip()

# Check for 'Product Code' column presence
if 'Product Code' not in df.columns:
    raise KeyError("Column 'Product Code' not found in the Excel file.")

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
        time.sleep(5)
        
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
        # Wait until the specs list div is present
        specs_list = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'dl.specs-list_GuAFh'))
        )
        product_info = "\n".join([spec.text.strip() for spec in specs_list])
    except Exception as e:
        print(f"Error extracting product info from {url}: {e}")
        product_info = 'Info Not Found'
    
    return product_info

# Create new columns in the DataFrame for product links and additional info
df['Product Link'] = df['Product Code'].apply(find_product_link)
df['Additional Info'] = df['Product Link'].apply(lambda link: extract_product_info(link) if link != 'Link Not Found' else 'Info Not Found')

# Save the updated DataFrame to a new Excel file
df.to_excel(output_file, sheet_name=sheet_name, index=False)

# Close the web driver
driver.quit()

print("Product links and additional information have been successfully updated in the Excel file.")
