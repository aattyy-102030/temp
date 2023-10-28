import os
import zipfile
from datetime import datetime
import shutil
import pytesseract
from PIL import Image
import pandas as pd
import glob

# Input for search query
search_query = input("The string you want to search for: ")

# Step 1: Rename .docx to .zip and move to './Input/zip'
zip_folder = './Input/zip'
docx_folder = './Input/docx'

# Clean up './Input/zip' folder after copying
for zfile in glob.glob(f"{zip_folder}/*.zip"):
    os.remove(zfile)

# Copy and rename Word files to ZIP folder
for filename in os.listdir(docx_folder):
    if filename.endswith('.docx'):
        shutil.copy(os.path.join(docx_folder, filename), os.path.join(zip_folder, filename.replace('.docx', '.zip')))

# Step 2: Unzip files to './Output/(Datetime)_DetectText-InPic'
current_time = datetime.now().strftime('%Y%m%d-%H%M%S')
output_folder = f'./Output/{current_time}_DetectText-InPic'
os.makedirs(output_folder, exist_ok=True)

table_data = []

for filename in os.listdir(zip_folder):
    if filename.endswith('.zip'):
        word_filename = filename.replace('.zip', '.docx')
        subfolder_name = os.path.join(output_folder, word_filename)
        os.makedirs(subfolder_name, exist_ok=True)

        with zipfile.ZipFile(os.path.join(zip_folder, filename), 'r') as zip_ref:
            zip_ref.extractall(subfolder_name)

        image_folder = os.path.join(subfolder_name, 'word', 'media')
        if os.path.exists(image_folder):
            for image_filename in os.listdir(image_folder):
                if image_filename.endswith(('.png', '.jpg', '.jpeg')):
                    image_path = os.path.join(image_folder, image_filename)
                    image = Image.open(image_path)
                    extracted_text = pytesseract.image_to_string(image, lang='eng')

                    if search_query in extracted_text:
                        table_data.append({
                            'Detected Text': search_query,
                            'Word File': word_filename,
                            'Image File': image_filename
                        })

# Step 4: Write result to Excel file
df = pd.DataFrame(table_data)
df.to_excel(f'{output_folder}/Result.xlsx', index=False)
