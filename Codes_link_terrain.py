import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Path to the input and output Excel files
input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Terrain Images\Codess.xlsx'
output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Terrain Images\ImagesAndDescription.xlsx'
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

# Initialize the web driver (make sure you have the appropriate driver installed)
driver = webdriver.Chrome()

# Function to search for the product link, image URL, and product title
def find_product_info(product_code):
    search_url = "https://www.tglynes.co.uk/site/search"
    driver.get(search_url)
    
    # Try to find the search box using the provided class name and attributes
    try:
        search_box = driver.find_element(By.ID, 'ctl00_mainHeader_MasterTop_b419_txtName')
    except Exception as e:
        print(f"Search box not found for product code {product_code}: {e}")
        return 'Search Box Not Found', '', '', ''
    
    search_box.send_keys(product_code)
    search_box.send_keys(Keys.RETURN)
    
    time.sleep(5)  # Wait for the page to load (adjust as needed)
    
    # Locate the product link and click on it
    try:
        product_link = driver.find_element(By.CLASS_NAME, 'FlexProductListingText')
        product_title = product_link.text.strip()
        product_link.click()
        time.sleep(3)  # Wait for the page to load (adjust as needed)
        product_page_url = driver.current_url
    except Exception as e:
        print(f"Product link not found for product code {product_code}: {e}")
        return 'Product Link Not Found', '', '', ''
    
    # Locate the image element
    try:
        img_element = driver.find_element(By.CLASS_NAME, 'FlexProductImage')
        image_url = img_element.get_attribute('data-original')
    except Exception as e:
        print(f"Image not found for product code {product_code}: {e}")
        image_url = ''
    
    # Save a screenshot of the product page
    screenshot_filename = f'{product_code}_screenshot.png'
    driver.save_screenshot(screenshot_filename)
    print(f"Screenshot saved for product code {product_code}")
    
    return product_page_url, image_url, product_title, screenshot_filename

# Create new columns in the DataFrame for product links, image URLs, product titles, and screenshot filenames
df[['Product Page URL', 'Image URL', 'Product Title', 'Screenshot File']] = df['Product Code'].apply(find_product_info).apply(pd.Series)

# Save the updated DataFrame to a new Excel file
df.to_excel(output_file, sheet_name=sheet_name, index=False)

# Close the web driver
driver.quit()

print("Product links, images, titles, and screenshots have been successfully updated in the Excel file.")
