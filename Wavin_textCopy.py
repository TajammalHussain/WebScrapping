# # import pandas as pd
# # from selenium import webdriver
# # from selenium.webdriver.common.by import By
# # from selenium.webdriver.common.keys import Keys
# # from selenium.webdriver.chrome.service import Service as ChromeService
# # from selenium.webdriver.support.ui import WebDriverWait
# # from selenium.webdriver.support import expected_conditions as EC
# # from webdriver_manager.chrome import ChromeDriverManager
# # import time

# # # Path to the input and output Excel files
# # input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Wavin images\Codes.xlsx'
# # output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Wavin images\WavinTextSecond.xlsx'
# # sheet_name = 'Sheet1'

# # # Load the Excel file
# # df = pd.read_excel(input_file, sheet_name=sheet_name)

# # # Ensure 'Product Code' column is correctly identified (handle potential leading/trailing spaces)
# # df.columns = df.columns.str.strip()

# # # Check for 'Product Code' column presence
# # if 'Product Code' not in df.columns:
# #     raise KeyError("Column 'Product Code' not found in the Excel file.")

# # # Initialize the web driver
# # driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

# # # Function to search for the product link based on the product code
# # def find_product_link(product_code):
# #     search_url = "https://www.wavin.com/en-gb/search"
# #     driver.get(search_url)
    
# #     try:
# #         # Wait until the search box is present and find it
# #         search_box = WebDriverWait(driver, 10).until(
# #             EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="search"]'))
# #         )
# #         search_box.send_keys(product_code)
# #         search_box.send_keys(Keys.RETURN)
        
# #         # Wait for the search results to load
# #         time.sleep(1)
        
# #         # Find the first search result link
# #         first_result_link = driver.find_element(By.CSS_SELECTOR, 'a[data-qa="elastic-search-item_link"]')
# #         product_page_url = first_result_link.get_attribute('href')
# #         return product_page_url
# #     except Exception as e:
# #         print(f"Error finding product link for {product_code}: {e}")
# #         return 'Link Not Found'

# # # Function to extract additional information from the product page
# # def extract_product_info(url):
# #     driver.get(url)
    
# #     try:
# #         # Wait until the collapsible triggers are present
# #         WebDriverWait(driver, 10).until(
# #             EC.presence_of_element_located((By.CSS_SELECTOR, '.collapsible-trigger_lKDGc'))
# #         )
        
# #         # Find all collapsible triggers and click to expand them
# #         collapsible_triggers = driver.find_elements(By.CSS_SELECTOR, '.collapsible-trigger_lKDGc')
# #         for trigger in collapsible_triggers:
# #             try:
# #                 trigger.click()
# #                 time.sleep(2)  # Give some time for the content to expand
# #             except Exception as e:
# #                 print(f"Error clicking trigger: {e}")
# #                 continue
        
# #         # Extract the text within <dt> and <dd> tags with specific data-qa attribute
# #         specs_lists = driver.find_elements(By.CSS_SELECTOR, 'dl.specs-list_GuAFh')
# #         product_info = ""
        
# #         for specs_list in specs_lists:
# #             dt_elements = specs_list.find_elements(By.TAG_NAME, 'dt')
# #             dd_elements = specs_list.find_elements(By.CSS_SELECTOR, 'dd[data-qa="product-specifications_spec-value"]')
            
# #             for dt, dd in zip(dt_elements, dd_elements):
# #                 dt_text = dt.text.strip()
# #                 dd_text = dd.text.strip()
# #                 product_info += f"{dt_text}: {dd_text}\n"
        
# #         return product_info.strip()
    
# #     except Exception as e:
# #         print(f"Error extracting product info from {url}: {e}")
# #         return 'Info Not Found'

# # # Iterate over each row, find the product link, extract additional info, and save after each row
# # for index, row in df.iterrows():
# #     product_code = row['Product Code']
# #     product_link = find_product_link(product_code)
# #     additional_info = extract_product_info(product_link) if product_link != 'Link Not Found' else 'Info Not Found'
    
# #     df.at[index, 'Product Link'] = product_link
# #     df.at[index, 'Additional Info'] = additional_info
    
# #     # Save the DataFrame after each update
# #     df.to_excel(output_file, sheet_name=sheet_name, index=False)

# # # Close the web driver
# # driver.quit()

# # print("Product links and additional information have been successfully updated in the Excel file.")
# import pandas as pd
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.chrome.service import Service as ChromeService
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from webdriver_manager.chrome import ChromeDriverManager
# import time

# # Path to the input and output Excel files
# input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Wavin images\Codes.xlsx'
# output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Wavin images\WavinTextSecond.xlsx'
# sheet_name = 'Sheet1'

# # Load the Excel file
# df = pd.read_excel(input_file, sheet_name=sheet_name)

# # Ensure 'Product Code' column is correctly identified (handle potential leading/trailing spaces)
# df.columns = df.columns.str.strip()

# # Check for 'Product Code' column presence
# if 'Product Code' not in df.columns:
#     raise KeyError("Column 'Product Code' not found in the Excel file.")

# # Convert 'Product Code' to string
# df['Product Code'] = df['Product Code'].astype(str)

# # Initialize the web driver
# driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

# # Function to search for the product link based on the product code
# def find_product_link(product_code):
#     search_url = "https://www.wavin.com/en-gb/search"
#     driver.get(search_url)
    
#     try:
#         # Wait until the search box is present and find it
#         search_box = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="search"]'))
#         )
#         search_box.send_keys(product_code)
#         search_box.send_keys(Keys.RETURN)
        
#         # Wait for the search results to load
#         time.sleep(1)
        
#         # Find the first search result link
#         first_result_link = driver.find_element(By.CSS_SELECTOR, 'a[data-qa="elastic-search-item_link"]')
#         product_page_url = first_result_link.get_attribute('href')
#         return product_page_url
#     except Exception as e:
#         print(f"Error finding product link for {product_code}: {e}")
#         return 'Link Not Found'

# # Function to extract additional information from the product page
# def extract_product_info(url):
#     driver.get(url)
    
#     try:
#         # Wait until the collapsible triggers are present
#         WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.CSS_SELECTOR, '.collapsible-trigger_lKDGc'))
#         )
        
#         # Find all collapsible triggers and click to expand them if not already expanded
#         collapsible_triggers = driver.find_elements(By.CSS_SELECTOR, '.collapsible-trigger_lKDGc')
#         for trigger in collapsible_triggers:
#             try:
#                 trigger_button = trigger.find_element(By.XPATH, './button')
#                 if trigger_button.get_attribute('aria-expanded') == 'false':
#                     trigger.click()
#                     time.sleep(2)  # Give some time for the content to expand
#             except Exception as e:
#                 print(f"Error clicking trigger: {e}")
#                 continue
        
#         # Extract the text within <dt> and <dd> tags with specific data-qa attribute
#         dt_elements = driver.find_elements(By.CSS_SELECTOR, 'dt[data-qa="product-specifications_spec-label"]')
#         dd_elements = driver.find_elements(By.CSS_SELECTOR, 'dd[data-qa="product-specifications_spec-value"]')
        
#         product_info = ""
        
#         for dt, dd in zip(dt_elements, dd_elements):
#             dt_text = dt.text.strip()
#             dd_text = dd.text.strip()
#             if dt_text and dd_text:
#                 product_info += f"{dt_text}: {dd_text}\n"
        
#         return product_info.strip()
    
#     except Exception as e:
#         print(f"Error extracting product info from {url}: {e}")
#         return 'Info Not Found'

# # Iterate over each row, find the product link, extract additional info, and save after each row
# for index, row in df.iterrows():
#     product_code = row['Product Code']
#     product_link = find_product_link(product_code)
#     additional_info = extract_product_info(product_link) if product_link != 'Link Not Found' else 'Info Not Found'
    
#     df.at[index, 'Product Link'] = product_link
#     df.at[index, 'Additional Info'] = additional_info
    
#     # Save the DataFrame after each update
#     df.to_excel(output_file, sheet_name=sheet_name, index=False)

# # Close the web driver
# driver.quit()

# print("Product links and additional information have been successfully updated in the Excel file.")
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Function to extract additional information from the product page
def extract_product_info(driver, url):
    driver.get(url)
    
    try:
        # Wait until the collapsible triggers are present
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.collapsible-trigger_lKDGc'))
        )
        
        # Find all collapsible triggers and click to expand them
        collapsible_triggers = driver.find_elements(By.CSS_SELECTOR, '.collapsible-trigger_lKDGc')
        for trigger in collapsible_triggers:
            try:
                trigger.click()
                WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '.collapsible-content_iZtkC'))
                )
            except Exception as e:
                print(f"Error clicking trigger: {e}")
                continue
        
        # Extract the text within <dt> and <dd> tags with specific data-qa attribute
        specs_lists = driver.find_elements(By.CSS_SELECTOR, 'dl.specs-list_GuAFh')
        product_info = {}
        
        for specs_list in specs_lists:
            dt_elements = specs_list.find_elements(By.TAG_NAME, 'dt')
            dd_elements = specs_list.find_elements(By.TAG_NAME, 'dd')
            
            for dt, dd in zip(dt_elements, dd_elements):
                dt_text = dt.text.strip()
                dd_text = dd.text.strip()
                product_info[dt_text] = dd_text
        
        return product_info
    
    except Exception as e:
        print(f"Error extracting product info from {url}: {e}")
        return {}

# Initialize the web driver
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

# URL of the page to scrape
url = 'https://www.wavin.com/en-gb/product/f315b791-98e2-4b3c-93a8-3e172b7e076f?navType=search'

# Extract product info
product_info = extract_product_info(driver, url)

# Print the extracted data
for label, value in product_info.items():
    print(f'{label}: {value}')

# Close the web driver
driver.quit()
