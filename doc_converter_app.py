import os
import sys
import zipfile
import shutil
import threading
from tkinter import filedialog, messagebox, StringVar
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from pdf2docx import Converter
import fitz  # PyMuPDF 用來處理 PDF 轉文字或圖片

APP_NAME = "文件專業轉檔工具 (純 Python 輕量版)"

input_file_path = ""
output_folder_path = ""

def choose_file():
    global input_file_path
    file_path = filedialog.askopenfilename(
        title="選擇要轉換的文件",
        filetypes=[
            ("支援的文件", "*.pdf;*.docx;*.txt;*.pages"),
            ("PDF 檔案", "*.pdf"),
            ("Word 檔案", "*.docx"),
            ("Apple Pages 檔案", "*.pages"),
            ("文字檔", "*.txt"),
            ("所有檔案", "*.*")
        ]
    )
    if not file_path:
        return
    input_file_path = file_path
    source_label.config(text=f"來源檔案：\n{os.path.basename(file_path)}")

def choose_output_folder():
    global output_folder_path
    folder_path = filedialog.askdirectory(title="選擇輸出資料夾")
    if not folder_path:
        return
    output_folder_path = folder_path
    output_label.config(text=f"輸出資料夾：\n{output_folder_path}")

def convert_document():
    global input_file_path, output_folder_path
    
    if not input_file_path or not output_folder_path:
        messagebox.showwarning("提醒", "請先選擇來源檔案與輸出資料夾！")
        return

    target_format = format_var.get().lower()
    base_name = os.path.splitext(os.path.basename(input_file_path))[0]
    ext = os.path.splitext(input_file_path)[1][1:].lower()
    
    output_filename = f"{base_name}_converted.{target_format}"
    output_path = os.path.join(output_folder_path, output_filename)

    status_label.config(text="正在進行本地轉檔...")
    progress.start(10)
    window.update_idletasks()

    temp_pdf_path = None

    try:
        actual_input_pdf = input_file_path

        # 1. 如果是 Apple Pages 檔案，直接透過解壓縮提取內建的 PDF 預覽
        if ext == 'pages':
            temp_pdf_path = os.path.join(output_folder_path, f"temp_{base_name}.pdf")
            with zipfile.ZipFile(input_file_path, 'r') as z:
                # 尋找 Pages 內建的預覽 PDF
                pdf_internal_path = None
                for filename in z.namelist():
                    if filename.endswith('Preview.pdf') or filename.endswith('quicklook/Preview.pdf'):
                        pdf_internal_path = filename
                        break
                
                if pdf_internal_path:
                    with z.open(pdf_internal_path) as source, open(temp_pdf_path, "wb") as target:
                        shutil.copyfileobj(source, target)
                    actual_input_pdf = temp_pdf_path
                else:
                    raise Exception("無法從此 Pages 檔案中讀取預覽內容（格式不支援）")

        # 2. PDF 轉 Word (.docx)
        if (ext in ['pdf', 'pages']) and target_format == 'docx':
            cv = Converter(actual_input_pdf)
            cv.convert(output_path, start=0, end=None)
            cv.close()

        # 3. PDF 轉文字 (.txt)
        elif (ext in ['pdf', 'pages']) and target_format == 'txt':
            doc = fitz.open(actual_input_pdf)
            text_content = ""
            for page in doc:
                text_content += page.get_text() + "\n"
            doc.close()
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(text_content)

        # 4. Word 轉文字 (.docx -> .txt)
        elif ext == 'docx' and target_format == 'txt':
            import docx
            doc = docx.Document(input_file_path)
            text_content = "\n".join([p.text for p in doc.paragraphs])
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(text_content)

        else:
            raise Exception(f"不支援從 .{ext} 轉換為 .{target_format} 的組合")

        # 清理暫存檔
        if temp_pdf_path and os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)

        progress.stop()
        status_label.config(text="轉換完成！")
        messagebox.showinfo("成功", f"檔案已成功轉換並儲存至：\n{output_path}")

    except Exception as e:
        progress.stop()
        status_label.config(text="轉換失敗")
        if temp_pdf_path and os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        messagebox.showerror("錯誤", f"轉檔過程發生錯誤：\n{str(e)}")

def start_conversion_thread():
    threading.Thread(target=convert_document, daemon=True).start()

# UI 介面
window = tb.Window(title=APP_NAME, themename="cosmo", size=(700, 600))
window.resizable(False, False)

tb.Label(window, text="文件專業轉檔工具 (純 Python 離線版)", font=("Microsoft JhengHei UI", 16, "bold")).pack(pady=20)

tb.Button(window, text="選擇要轉換的檔案 (PDF/Word/Pages/TXT)", bootstyle="primary", command=choose_file, width=40).pack(pady=5)
source_label = tb.Label(window, text="尚未選擇來源檔案", font=("Microsoft JhengHei UI", 10), bootstyle="secondary")
source_label.pack(pady=5)

tb.Button(window, text="選擇輸出資料夾", bootstyle="info", command=choose_output_folder, width=40).pack(pady=5)
output_label = tb.Label(window, text="尚未選擇輸出資料夾", font=("Microsoft JhengHei UI", 10), bootstyle="secondary")
output_label.pack(pady=5)

tb.Label(window, text="目標輸出格式", font=("Microsoft JhengHei UI", 11, "bold")).pack(pady=(15, 5))
format_var = StringVar(value="docx")
tb.Combobox(window, textvariable=format_var, values=["docx", "txt"], state="readonly", width=20).pack(pady=5)

status_label = tb.Label(window, text="待命中", font=("Microsoft JhengHei UI", 11))
status_label.pack(pady=10)

progress = tb.Progressbar(window, length=500, mode="indeterminate", bootstyle="success-striped")
progress.pack(pady=10)

tb.Button(window, text="開始轉換", bootstyle="success", command=start_conversion_thread, width=20).pack(pady=15)

window.mainloop()
