import pandas as pd
from rake_nltk import Rake
import nltk
import os

# Download the NLTK data (only needed the first time)
nltk.download('punkt')
nltk.download('stopwords')

# Define the file paths
input_file_path = r'C:\Users\Tajammal\Desktop\MRF Brands upload\3 Geberit.xlsx'
output_file_path = r'C:\Users\Tajammal\Desktop\MRF Brands upload\3 Geberit.xlsx'

# Load the Excel file
df = pd.read_excel(input_file_path)

# Initialize the RAKE object
rake_nltk_var = Rake()

# Function to extract keywords
def extract_keywords(text):
    if pd.isna(text):
        return ''
    text = str(text)  # Ensure the text is a string
    rake_nltk_var.extract_keywords_from_text(text)
    key_words_dict_scores = rake_nltk_var.get_word_degrees()
    # Get the top 5 keywords
    top_keywords = sorted(key_words_dict_scores, key=key_words_dict_scores.get, reverse=True)[:5]
    return ', '.join(top_keywords)

# Assuming the column with descriptions is named 'Description'
df['Keywords'] = df['Description'].apply(extract_keywords)

# Save the results to a new Excel file
df.to_excel(output_file_path, index=False)

print(f"Keywords extracted and saved to {output_file_path}")
