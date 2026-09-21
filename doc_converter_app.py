import os
import sys
import subprocess
import threading
from tkinter import filedialog, messagebox, StringVar
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from pdf2docx import Converter

APP_NAME = "文件專業轉檔工具 (獨立引擎版)"

input_file_path = ""
output_folder_path = ""

# 自動尋找放在 .exe 旁邊的 libreoffice_portable 引擎
def get_libreoffice_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    side_path = os.path.join(base_dir, "libreoffice_portable", "program", "soffice.exe")
    if os.path.exists(side_path):
        return side_path

    return "soffice"

def choose_file():
    global input_file_path
    file_path = filedialog.askopenfilename(
        title="選擇要轉換的文件",
        filetypes=[
            ("支援的文件", "*.pdf;*.docx;*.doc;*.pptx;*.ppt;*.txt;*.pages"),
            ("PDF 檔案", "*.pdf"),
            ("Word 檔案", "*.docx;*.doc"),
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

    status_label.config(text="正在轉檔中，請稍候...")
    progress.start(10)
    window.update_idletasks()

    try:
        # A. PDF 轉 Word
        if ext == 'pdf' and target_format == 'docx':
            cv = Converter(input_file_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()

        # B. 透過本機旁邊的 LibreOffice 引擎處理 (完美支援 Word、Pages 轉 PDF/Word/TXT)
        elif target_format in ['pdf', 'docx', 'txt']:
            soffice_bin = get_libreoffice_path()
            if soffice_bin == "soffice":
                raise Exception("找不到獨立的 LibreOffice 引擎！\n請確保 'libreoffice_portable' 資料夾與 .exe 放 在同一個資料夾下。")

            cmd = [
                soffice_bin, '--headless', '--convert-to', target_format, 
                input_file_path, '--outdir', output_folder_path
            ]
            
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(cmd, capture_output=True, text=True, startupinfo=startupinfo)
            
            if result.returncode != 0:
                raise Exception(f"轉檔引擎失敗: {result.stderr}")
            
            default_out = os.path.join(output_folder_path, f"{base_name}.{target_format}")
            if os.path.exists(default_out) and os.path.normpath(default_out) != os.path.normpath(output_path):
                if os.path.exists(output_path):
                    os.remove(output_path)
                os.rename(default_out, output_path)
        else:
            raise Exception("不支援此轉換格式組合")

        progress.stop()
        status_label.config(text="轉換完成！")
        messagebox.showinfo("成功", f"檔案已成功轉換並儲存至：\n{output_path}")

    except Exception as e:
        progress.stop()
        status_label.config(text="轉換失敗")
        messagebox.showerror("錯誤", f"轉檔過程發生錯誤：\n{str(e)}")

def start_conversion_thread():
    threading.Thread(target=convert_document, daemon=True).start()

# UI 介面設定
window = tb.Window(title=APP_NAME, themename="cosmo", size=(700, 600))
window.resizable(False, False)

tb.Label(window, text="文件專業轉檔工具 (完美排版支援)", font=("Microsoft JhengHei UI", 16, "bold")).pack(pady=20)

tb.Button(window, text="選擇要轉換的檔案 (PDF/Word/Pages/TXT)", bootstyle="primary", command=choose_file, width=40).pack(pady=5)
source_label = tb.Label(window, text="尚未選擇來源檔案", font=("Microsoft JhengHei UI", 10), bootstyle="secondary")
source_label.pack(pady=5)

tb.Button(window, text="選擇輸出資料夾", bootstyle="info", command=choose_output_folder, width=40).pack(pady=5)
output_label = tb.Label(window, text="尚未選擇輸出資料夾", font=("Microsoft JhengHei UI", 10), bootstyle="secondary")
output_label.pack(pady=5)

tb.Label(window, text="目標輸出格式", font=("Microsoft JhengHei UI", 11, "bold")).pack(pady=(15, 5))
format_var = StringVar(value="pdf")
tb.Combobox(window, textvariable=format_var, values=["pdf", "docx", "txt"], state="readonly", width=20).pack(pady=5)

status_label = tb.Label(window, text="待命中", font=("Microsoft JhengHei UI", 11))
status_label.pack(pady=10)

progress = tb.Progressbar(window, length=500, mode="indeterminate", bootstyle="success-striped")
progress.pack(pady=10)

tb.Button(window, text="開始轉換", bootstyle="success", command=start_conversion_thread, width=20).pack(pady=15)

window.mainloop()
