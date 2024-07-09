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
output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Wavin images\WavinPDFLinks2.xlsx'
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
        time.sleep(1.5)
        # time-sleep is the the time taken by code in order get the results on the page, And it will direct us to grasp the the required information and then provide that details 
        # into the columns of the excel sheet.
        
        # Find the first search result link
        first_result_link = driver.find_element(By.CSS_SELECTOR, 'a[data-qa="elastic-search-item_link"]')
        product_page_url = first_result_link.get_attribute('href')
        return product_page_url
    except Exception as e:
        print(f"Error finding product link for {product_code}: {e}")
        return 'Link Not Found'

#Functions to extract product names and corresponding href links
#the functions are used to extract URLs for the pdf files from the official Wavin Website. 
# Function to extract product names and corresponding href links
def extract_product_links(url):
    driver.get(url)
    
    try:
        # Wait until all card-wrapper elements are present
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'li.card-wrapper_2DgF8'))
        )
        
        # Find all elements with class 'card-wrapper_2DgF8'
        card_elements = driver.find_elements(By.CSS_SELECTOR, 'li.card-wrapper_2DgF8')
        
        product_names = []
        href_links = []
        
        for card in card_elements:
            # Find the <h3> tag within the current card
            h3_tag = card.find_element(By.CSS_SELECTOR, 'h3[data-qa="teaser-card_title"]')
            product_name = h3_tag.text.strip() if h3_tag else 'Product Name Not Found'
            product_names.append(product_name)
            
            # Find the <a> tag within the current card
            a_tag = card.find_element(By.CSS_SELECTOR, 'a[data-qa="teaser-card_link"]')
            href_link = a_tag.get_attribute('href') if a_tag else 'Link Not Found'
            href_links.append(href_link)
        
        return product_names, href_links
    
    except Exception as e:
        print(f"Error extracting product links from {url}: {e}")
        return [], []

# Iterate over each row, find the product link, extract additional info, and save to the DataFrame
for index, row in df.iterrows():
    product_code = row['Product Code']
    product_link = find_product_link(product_code)
    
    if product_link != 'Link Not Found':
        product_names, href_links = extract_product_links(product_link)
        
        if product_names and href_links:
            for name, link in zip(product_names, href_links):
                df.at[index, name] = link
    
    # Save the DataFrame after each update
    df.to_excel(output_file, sheet_name=sheet_name, index=False)
    print(f"Updated record for product code: {product_code}")

# Close the web driver
driver.quit()

print("Product links have been successfully updated in the Excel file.")
