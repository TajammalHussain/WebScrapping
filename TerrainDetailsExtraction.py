import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Path to the Excel file containing product codes
file_path = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Terrain Images\code.xlsx'
output_file_path = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Terrain Images\extracted_product_details.xlsx'

# Read the Excel file
df = pd.read_excel(file_path)

# Initialize the web driver (assuming Chrome; make sure the chromedriver is in your PATH or specify its path)
driver = webdriver.Chrome()

# List to store the extracted details
product_details = []

# Function to fetch product details from URL
def get_product_details(url):
    try:
        # Open the URL
        driver.get(url)
        
        # Wait for the product title link to be present
        product_title_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//h3[@class='info--title']/a"))
        )
        
        # Extract product details
        product_title = product_title_link.text.strip()
        product_category = driver.find_element(By.CLASS_NAME, 'info--category').text.strip()
        product_code = driver.find_element(By.CLASS_NAME, 'sku-code').find_element(By.TAG_NAME, 'span').text.strip()
        image_url = driver.find_element(By.CLASS_NAME, 'image--img').get_attribute('src')
        full_details_link = driver.find_element(By.CLASS_NAME, 'button__blue_light').get_attribute('href')
        
        # Try to find the Find Stockist link
        try:
            find_stockist_link = driver.find_element(By.CLASS_NAME, 'button__green').get_attribute('href')
        except NoSuchElementException:
            find_stockist_link = ''
        
        return {
            'Product Code': product_code,
            'Product Title': product_title,
            'Product Category': product_category,
            'Image URL': image_url,
            'Full Details Link': full_details_link,
            'Find Stockist Link': find_stockist_link
        }
    
    except TimeoutException:
        print(f"Timeout fetching product details from URL {url}")
        return None
    except NoSuchElementException:
        print(f"Element not found fetching product details from URL {url}")
        return None
    except Exception as e:
        print(f"Error fetching product details from URL {url}: {e}")
        return None

# Iterate over each product code and fetch details
for code in df['Product Code']:
    url = f"https://www.polypipe.com/search?combine={code}"
    print(f"Searching for product code: {code}")
    product_details_dict = get_product_details(url)
    
    if product_details_dict:
        print(f"Found product details for code: {code}")
        product_details.append(product_details_dict)
    else:
        print(f"No product details found for code: {code}")
        product_details.append({
            'Product Code': code,
            'Product Title': '',
            'Product Category': '',
            'Image URL': '',
            'Full Details Link': '',
            'Find Stockist Link': ''
        })

# Close the web driver
driver.quit()

# Create a new DataFrame with the extracted details
result_df = pd.DataFrame(product_details)

# Save the new DataFrame to an Excel file
result_df.to_excel(output_file_path, index=False)

print(f"Extracted product details saved to {output_file_path}")
