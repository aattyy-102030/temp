from docx import Document

def replace_text_while_keeping_styles(doc_path, replacements):
    # Word文書を読み込む
    doc = Document(doc_path)

    # 各段落に対して処理を行う
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            # 指定されたキー（置換前の文字列）に基づいて置換を行う
            replace_text_in_paragraph(paragraph, key, value)

    # 編集後の文書を保存
    doc.save('updated_document.docx')

def replace_text_in_paragraph(paragraph, key, value):
    new_runs = []  # 新しいrunを格納するためのリスト
    for run in paragraph.runs:
        if key in run.text:
            # 置換対象のテキストを含む場合、それを指定された文字列で置換
            parts = run.text.split(key)
            for i, part in enumerate(parts):
                # 新しいrunを作成して、元のテキストの一部を設定
                new_run = run._element._new_r()
                new_run.text = part
                if i < len(parts) - 1:
                    # 置換後のテキストを新しいrunに追加
                    new_run = run._element._new_r()
                    new_run.text = value
                new_runs.append(new_run)
        else:
            # 置換不要の場合は、元のrunを新しいリストに追加
            new_runs.append(run._element)

    # 元のrunを段落から削除
    for run in paragraph.runs:
        paragraph._element.remove(run._element)
    
    # 新しいrunを段落に追加
    for new_run in new_runs:
        paragraph._element.append(new_run)

# 置換する文字列の辞書
replacements = {
    '{Table}': 'Table',
    '{Figure}': 'Figure',
    '{Section}': 'Section'
}

# 編集する文書のパス
doc_path = 'path_to_your_document.docx'
# 編集処理の実行
replace_text_while_keeping_styles(doc_path, replacements)