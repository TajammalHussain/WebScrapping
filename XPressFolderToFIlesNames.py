# import os
# import shutil
# #lets write a function, a function that is related to  to pdf files in the folder.
# # Function to find PDF files in a directory
# def find_pdfs(directory):
#     pdf_files = []
#     for filename in os.listdir(directory):
#         if filename.endswith(".pdf"):
#             pdf_files.append(filename)
#     return pdf_files
# # this is simple program, for ignoring the folders, and copy their names for their files. 
# # Main function to consolidate PDF files
# def consolidate_pdfs(source_dir, destination_dir):
#     # Create the destination directory if it doesn't exist
#     if not os.path.exists(destination_dir):
#         os.makedirs(destination_dir)

#     # Iterate through each product code folder in the source directory
#     for product_code in os.listdir(source_dir):
#         product_folder = os.path.join(source_dir, product_code)
#         if os.path.isdir(product_folder):
#             # Find PDF file in the product code folder
#             pdf_files = find_pdfs(product_folder)
#             if pdf_files:
#                 # Assuming only one PDF file per folder, handle first file found
#                 pdf_file = pdf_files[0]
#                 # Construct new file name with product code and copy to destination
#                 new_filename = f"{product_code}.pdf"
#                 source_file = os.path.join(product_folder, pdf_file)
#                 destination_file = os.path.join(destination_dir, new_filename)
#                 shutil.copyfile(source_file, destination_file)
#                 print(f"Copied '{pdf_file}' to '{new_filename}'")

# # Example usage:
# if __name__ == "__main__":
#     # Replace with your source directory containing product code folders
#     source_directory = r"C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\DownloadPDFF"
#     # Replace with the directory where you want to consolidate PDF files
#     destination_directory = r"C:\Users\Tajammal\Desktop\MRF Brands upload\Xpress images\ConsolidatedPDFs"

#     consolidate_pdfs(source_directory, destination_directory)
# #the above code works perfefctly, and works according to the requirement. the folders contains the name of the codes of pegler yorkshire products, and they
# # are copied and saved the pdf files according to the codes of the products.
# #vsh_express_copper_gas_elbow_90_2_x_press.pdf to 39783
# #
import time

start_number = 38009

for i in range(250):
    print("Pegler Yorkshire", +start_number + i)
    time.sleep(0.5)
