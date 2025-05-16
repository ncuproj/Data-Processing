import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# 提取局處室
def extract_parentheses(text):
    if isinstance(text, str):
        match_colon = re.search(r'[\uFF1A\u3002\uFF1B][\s]*[\(（](.*?)[\)）]', text)
        if match_colon:
            content = match_colon.group(1)
            if len(content) > 1:
                return content.replace('、', ';')
            return '抓取不到局處室'

        match_last = re.findall(r'[\(（](.*?)[\)）]', text)
        if match_last:
            content = match_last[-1]
            if len(content) > 1:
                return content.replace('、', ';')
            return '抓取不到局處室'
    return '抓取不到局處室'

# 去除非法字元用於檔案名稱
def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', '_', name)

# 主處理邏輯 
def process_excel(file_path):
    try:
        df = pd.read_excel(file_path)

        output_dir = os.path.join(os.getcwd(), "112簡報表處理")
        os.makedirs(output_dir, exist_ok=True)

        record_map = {}  # 用來收集每個檔案的資料

        for _, row in df.iterrows():
            try:
                folder_name = str(row['質詢議員']).strip()
                file_base = str(row['屆、次、會議、組']).strip()
                file_name = sanitize_filename(file_base) + '.xlsx'
                full_topic = str(row['質詢議題']).strip()
                agency_extracted = extract_parentheses(full_topic)

                key = (folder_name, file_name)
                new_entry = {'質詢議題': full_topic, '局處室': agency_extracted}

                if key not in record_map:
                    record_map[key] = []
                record_map[key].append(new_entry)

            except Exception as e_inner:
                output_text.insert(tk.END, f"錯誤處理行：{e_inner}\n")

        # 寫入檔案階段
        for (folder_name, file_name), records in record_map.items():
            folder_path = os.path.join(output_dir, folder_name)
            os.makedirs(folder_path, exist_ok=True)
            file_path_out = os.path.join(folder_path, file_name)

            new_df = pd.DataFrame(records)

            if os.path.exists(file_path_out):
                existing_df = pd.read_excel(file_path_out)
                combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                combined_df.drop_duplicates(inplace=True)
            else:
                combined_df = new_df.drop_duplicates()

            combined_df.to_excel(file_path_out, index=False, engine='openpyxl')

        output_text.insert(tk.END, f"處理完成：{file_path}\n\n")
    except Exception as e:
        output_text.insert(tk.END, f"處理錯誤：{file_path} - {e}\n")

# --- GUI 控制 ---
def select_file():
    file_path = filedialog.askopenfilename(title="選擇 Excel 檔案", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_path_var.set(file_path)

def run_processing():
    path = file_path_var.get().strip()
    if not path:
        messagebox.showwarning("未選擇檔案", "請先選擇 Excel 檔案")
        return

    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, f"開始處理：{path}\n")
    process_excel(path)
    messagebox.showinfo("完成", "處理完成！")

# --- GUI 介面 ---
root = tk.Tk()
root.title("簡報表 Excel 處理工具")
root.geometry("700x500")

frame = tk.Frame(root)
frame.pack(pady=10)

tk.Button(frame, text="選擇 Excel 檔案", command=select_file).grid(row=0, column=0, padx=5)
tk.Button(frame, text="開始處理", command=run_processing).grid(row=0, column=1, padx=5)

file_path_var = tk.StringVar()
tk.Label(root, text="選擇的檔案：").pack(anchor="w", padx=10)
tk.Entry(root, textvariable=file_path_var, width=100).pack(padx=10)

output_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=20)
output_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

root.mainloop()
