# import pandas as pd
# import requests
# from bs4 import BeautifulSoup
# import time

# # Path to the input and output Excel files
# input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\Code.xlsx'
# output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\Code_Links.xlsx'
# sheet_name = 'Sheet1'

# # Load the Excel file
# df = pd.read_excel(input_file, sheet_name=sheet_name)

# # Print the DataFrame to inspect column names
# print("DataFrame columns:", df.columns)

# # Ensure 'Product Code' column is correctly identified (handle potential leading/trailing spaces)
# df.columns = df.columns.str.strip()

# # Check for 'Product Code' column presence
# if 'Product Code' not in df.columns:
#     raise KeyError("Column 'Product Code' not found in the Excel file.")

# # Function to find product link based on product code
# def find_product_link(product_code):
#     search_url = f"https://www.tglynes.co.uk/site/search?q=200.125.40"
#     response = requests.get(search_url)
    
#     # Check if the request was successful
#     if response.status_code == 200:
#         soup = BeautifulSoup(response.content, 'html.parser')
#         # Find product link if available
#         product_link = soup.find('a', class_='item-title')
#         if product_link:
#             return product_link['href']
#         else:
#             return 'Product Link Not Found'
#     else:
#         print(f"Failed to retrieve data for product code {product_code}. Status code: {response.status_code}")
#         return 'Request Failed'

# # Create a new column in the DataFrame for product links
# df['Product Link'] = df['Product Code'].apply(find_product_link)

# # Save the updated DataFrame to a new Excel file
# df.to_excel(output_file, sheet_name=sheet_name, index=False)

# print("Product links have been successfully updated in the Excel file.")
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Path to the input and output Excel files
input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\Code.xlsx'
output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\Code_Links.xlsx'
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

# Function to search for the product link based on the product code
def find_product_link(product_code):
    search_url = "https://www.tglynes.co.uk/site/search?q=38036"
    driver.get(search_url)
    
    # Try to find the search box using the provided class name and attributes
    try:
        search_box = driver.find_element(By.ID, 'ctl00_mainHeader_MasterTop_b419_txtName')
    except Exception as e:
        print(f"Search box not found for product code {product_code}: {e}")
        return 'Search Box Not Found'
    
    search_box.send_keys(product_code)
    search_box.send_keys(Keys.RETURN)
    
    time.sleep(5)  # Wait for the page to load (adjust as needed)
    
    # Capture the current URL (search results page URL)
    search_results_url = driver.current_url
    
    return search_results_url

# Create a new column in the DataFrame for product links
df['Product Link'] = df['Product Code'].apply(find_product_link)

# Save the updated DataFrame to a new Excel file
df.to_excel(output_file, sheet_name=sheet_name, index=False)

# Close the web driver
driver.quit()

print("Product links have been successfully updated in the Excel file.")
