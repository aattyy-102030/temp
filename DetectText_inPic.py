import os
from PIL import Image
import pytesseract

# Tesseract-OCRのパスを指定（Windowsの場合）
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# InputフォルダとOutputフォルダのパスを指定
input_folder_path = "./Input"
output_folder_path = "./Output"

# 特定の文字列（例："apple"）を設定
target_word = "apple"

# 出力結果を保存するためのリスト
output_list = []

# Inputフォルダ内の各画像ファイルに対してOCRを実行
for image_file in os.listdir(input_folder_path):
    if image_file.endswith(('.png', '.jpg', '.jpeg')):
        image_path = os.path.join(input_folder_path, image_file)
        image = Image.open(image_path)

        # OCRで画像内のテキストを抽出
        extracted_text = pytesseract.image_to_string(image, lang='eng')

        # 特定の文字列がテキスト内に存在するか確認
        if target_word in extracted_text:
            output_list.append(image_file)

# 出力結果をOutputフォルダ内のresult.txtに保存
result_file_path = os.path.join(output_folder_path, "result.txt")
with open(result_file_path, "w") as f:
    for item in output_list:
        f.write(f"{item}\n")

print("Process completed. Check the Output folder for result.txt.")
