import os
import zipfile
from datetime import datetime
import pytesseract
from PIL import Image
import pandas as pd
import shutil

# Clean up './Input/zip' folder at the beginning
zip_folder = './Input/zip'
if os.path.exists(zip_folder):
    shutil.rmtree(zip_folder)
os.makedirs(zip_folder, exist_ok=True)

# Step 1: Rename .docx to .zip and move to './Input/zip'
docx_folder = './Input/docx'

for filename in os.listdir(docx_folder):
    if filename.endswith('.docx'):
        os.rename(os.path.join(docx_folder, filename), os.path.join(zip_folder, filename.replace('.docx', '.zip')))

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

                    if 'Apple' in extracted_text:
                        table_data.append({
                            'Detected Text': 'Apple',
                            'Word File': word_filename,
                            'Image File': image_filename
                        })

# Step 4: Write result to Excel file
df = pd.DataFrame(table_data)
df.to_excel(f'{output_folder}/Result.xlsx', index=False)
