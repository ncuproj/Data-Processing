import os
from pdf2docx import Converter

# 設定主資料夾路徑
root_folder = ".\議會速紀錄"

# 遍歷所有子資料夾
for subdir, _, files in os.walk(root_folder):
    for file in files:
        if file.lower().endswith(".pdf"):
            pdf_path = os.path.join(subdir, file)
            docx_path = os.path.join(subdir, file.replace(".pdf", ".docx"))
            
            # 轉換 PDF 到 DOCX
            try:
                cv = Converter(pdf_path)
                cv.convert(docx_path, start=0, end=None)
                cv.close()
                print(f"轉換成功: {docx_path}")
            except Exception as e:
                print(f"轉換失敗: {pdf_path}，錯誤訊息: {e}")
