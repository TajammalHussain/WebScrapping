import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Path to the input and output Excel files
input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\Codss.xlsx'
output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\Code_Description3.xlsx'
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

# Initialize the web driver (make sure you have the appropriate driver installed)
driver = webdriver.Chrome()

# Function to search for the product link and description based on the product code
def find_product_info(product_code):
    search_url = "https://www.tglynes.co.uk/site/search"
    driver.get(search_url)
    
    try:
        search_box = driver.find_element(By.ID, 'ctl00_mainHeader_MasterTop_b419_txtName')
    except Exception as e:
        print(f"Search box not found for product code {product_code}: {e}")
        return 'Search Box Not Found', 'Description Not Found'
    
    search_box.send_keys(product_code)
    search_box.send_keys(Keys.RETURN)
    
    time.sleep(1)  # Wait for the page to load (adjust as needed)
    
    search_results_url = driver.current_url
    
    try:
        # Find the product link element
        product_link_element = driver.find_element(By.CLASS_NAME, 'FlexProductListingText.FlexProductDescLink')
        product_link = product_link_element.get_attribute('href')
        
        # Navigate to the product link
        driver.get(product_link)
        
        time.sleep(1)  # Wait for the page to load (adjust as needed)
        
        # Find the description element on the product page
        description_element = driver.find_element(By.CLASS_NAME, 'col-sm-6.col-lg-5.col-lg-offset-1.hidden-xs.pb-2')
        description_text = description_element.find_element(By.TAG_NAME, 'h1').text
    except Exception as e:
        print(f"Description not found for product code {product_code}: {e}")
        product_link = search_results_url
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

print("Product links and descriptions have been successfully updated in the Excel file.")
