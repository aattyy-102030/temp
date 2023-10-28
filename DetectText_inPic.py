import os
import zipfile
import shutil
from datetime import datetime
from PIL import Image
import pytesseract

# 1. Wordファイルの拡張子を.zipに変更
input_docx_folder = "./Input/docx"
input_zip_folder = "./Input/zip"

if not os.path.exists(input_zip_folder):
    os.makedirs(input_zip_folder)

for filename in os.listdir(input_docx_folder):
    if filename.endswith(".docx"):
        src = os.path.join(input_docx_folder, filename)
        dst = os.path.join(input_zip_folder, filename.replace(".docx", ".zip"))
        shutil.copy(src, dst)

# 2. .zipファイルを解凍
execution_time = datetime.now().strftime("%Y%m%d-%H%M%S")
output_folder = f"./Output/{execution_time}_DetectText-InPic"

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

for filename in os.listdir(input_zip_folder):
    if filename.endswith(".zip"):
        zip_path = os.path.join(input_zip_folder, filename)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(output_folder)

# 3 & 4. OCRと特定のテキストの検索
target_word = "Apple"
output_list = []

for root, dirs, files in os.walk(output_folder):
    for filename in files:
        if filename.endswith(('.png', '.jpg', '.jpeg')):
            image_path = os.path.join(root, filename)
            image = Image.open(image_path)
            extracted_text = pytesseract.image_to_string(image, lang='eng')

            if target_word in extracted_text:
                output_list.append(filename)

# 5. 結果をテキストファイルに出力
result_file_path = os.path.join(output_folder, "Result.txt")
with open(result_file_path, "w") as f:
    for item in output_list:
        f.write(f"{item}\n")

print(f"Process completed. Check {output_folder} for Result.txt.")
