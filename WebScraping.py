# # # import requests
# # # from bs4 import BeautifulSoup

# # # page_to_scrape= requests.get("http://quotes.toscrape.com")
# # # soup= BeautifulSoup(page_to_scrape.text, "html.parser")
# # # quotes= soup.findAll("span", attrs={"class":"text"})
# # # author= soup.findAll("small", attrs={"class": "author"})

# # # for quote, author in zip(quotes, author):
# # #     print(quote.text + "-" +author.text)
# # # print("I am representing some of the key quotes which are being copied from one of the website which are encompassing quotes from different perspectives")
# # # print("Alhamdulillah Web Scrapping is lovely to perform. ALHAMDULILLAH")


# # from numpy import random
# # # x= random.randint(1000)
# # # print(x)

# # # y= random.rand()
# # # print(y)
# # # z= random.rand(3,5)
# # # print(z)

# # x= random.choice([3,5,7,9], size=(3,5))
# # print(x)

# import pandas as pd
# from itertools import combinations, permutations

# # Correct the file path for Windows
# input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Material List - For TJ - Copy.xlsx'
# output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\output_with_factors.xlsx'

# # Read the Excel file
# df = pd.read_excel(input_file)

# # Function to generate combinations and permutations
# def generate_factors(words):
#     factors = []
#     for r in range(2, min(len(words), 6) + 1):  # Only up to quintuplets
#         for combo in combinations(words, r):
#             for perm in permutations(combo):
#                 factors.append(' '.join(perm))
#     return '; '.join(factors)

# # Apply the function to each row in the DataFrame
# df['Factors'] = df.apply(lambda row: generate_factors(row.astype(str).tolist()), axis=1)

# # Save the modified DataFrame to a new Excel file
# df.to_excel(output_file, index=False)

# print(f"Output saved to {output_file}")

import pandas as pd
from itertools import combinations, permutations

# Correct the file path for Windows
input_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\Material List - For TJ - Copy.xlsx'
output_file = r'C:\Users\Tajammal\Desktop\MRF Brands upload\33output_with_factors.xlsx'

# Read the Excel file
df = pd.read_excel(input_file)

# Function to generate combinations and permutations
def generate_factors(text):
    words = text.split()
    factors = []
    for r in range(2, min(len(words), 6) + 1):  # Only up to quintuplets
        for combo in combinations(words, r):
            for perm in permutations(combo):
                factors.append(' '.join(perm))
    return '; '.join(factors)

# Combine all values in each row into a single string
df['Combined'] = df.astype(str).apply(lambda row: ' '.join(row), axis=1)

# Apply the function to each combined string
df['Factors'] = df['Combined'].apply(generate_factors)

# Drop the intermediate 'Combined' column
df.drop(columns=['Combined'], inplace=True)

# Save the modified DataFrame to a new Excel file
df.to_excel(output_file, index=False)

print(f"Output saved to {output_file}")
