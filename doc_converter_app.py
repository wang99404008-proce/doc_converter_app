import os
import sys
import subprocess
import threading
from tkinter import filedialog, messagebox, StringVar
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from pdf2docx import Converter

# ==================================
# 全域變數與常數
# ==================================
APP_NAME = "文件專業轉檔工具 (內建離線引擎版)"

SUPPORTED_INPUTS = (".pdf", ".docx", ".doc", ".pptx", ".ppt", ".txt", ".pages", ".odt", ".rtf")

input_file_path = ""
output_folder_path = ""

# ==================================
# 動態尋找內建或系統的 LibreOffice 路徑
# ==================================
def get_libreoffice_path():
    # 1. 檢查是否為 PyInstaller 打包後的暫存執行環境
    if hasattr(sys, '_MEIPASS'):
        bundled_path = os.path.join(sys._MEIPASS, "libreoffice_portable", "program", "soffice.exe")
        if os.path.exists(bundled_path):
            return bundled_path

    # 2. 檢查本機專案資料夾內的內建路徑 (開發測試用)
    local_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "libreoffice_portable", "program", "soffice.exe")
    if os.path.exists(local_path):
        return local_path

    # 3. 檢查 Windows 常見安裝路徑
    possible_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"soffice"
    ]
    for path in possible_paths:
        if path == "soffice":
            return "soffice"
        if os.path.exists(path):
            return path
            
    return "soffice"

# ==================================
# 選擇來源檔案
# ==================================
def choose_file():
    global input_file_path
    file_path = filedialog.askopenfilename(
        title="選擇要轉換的文件",
        filetypes=[
            ("所有支援的文件", "*.pdf;*.docx;*.doc;*.pptx;*.ppt;*.txt;*.pages;*.odt;*.rtf"),
            ("PDF 檔案", "*.pdf"),
            ("Word 檔案", "*.docx;*.doc;*.odt"),
            ("Apple Pages 檔案", "*.pages"),
            ("PowerPoint 檔案", "*.pptx;*.ppt"),
            ("文字檔", "*.txt;*.rtf"),
            ("所有檔案", "*.*")
        ]
    )
    if not file_path:
        return
    
    input_file_path = file_path
    file_name = os.path.basename(file_path)
    source_label.config(text=f"來源檔案：\n{file_name}")

# ==================================
# 選擇輸出資料夾
# ==================================
def choose_output_folder():
    global output_folder_path
    folder_path = filedialog.askdirectory(title="選擇輸出資料夾")
    if not folder_path:
        return
    
    output_folder_path = folder_path
    output_label.config(text=f"輸出資料夾：\n{output_folder_path}")

# ==================================
# 核心轉檔邏輯 (支援 Pages、Word、PDF 等互轉)
# ==================================
def convert_document():
    global input_file_path, output_folder_path
    
    if not input_file_path:
        messagebox.showwarning("提醒", "請先選擇要轉換的檔案！")
        return
    
    if not output_folder_path:
        messagebox.showwarning("提醒", "請先選擇輸出資料夾！")
        return

    target_format = format_var.get().lower()
    base_name = os.path.splitext(os.path.basename(input_file_path))[0]
    ext = os.path.splitext(input_file_path)[1][1:].lower()
    
    output_filename = f"{base_name}_converted.{target_format}"
    output_path = os.path.join(output_folder_path, output_filename)

    status_label.config(text=f"正在轉換中，請稍候...")
    progress.start(10)
    window.update_idletasks()

    try:
        # A. PDF 轉 Word (.pdf -> .docx)
        if ext == 'pdf' and target_format == 'docx':
            cv = Converter(input_file_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()

        # B. 透過內建的 LibreOffice 引擎處理 (支援 .pages、.docx、.pptx、.txt、.pdf 互轉)
        elif target_format in ['docx', 'pdf', 'txt']:
            soffice_bin = get_libreoffice_path()
            cmd = [
                soffice_bin, '--headless', '--convert-to', target_format, 
                input_file_path, '--outdir', output_folder_path
            ]
            
            # 隱藏 Windows 命令提示字元黑框
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(cmd, capture_output=True, text=True, startupinfo=startupinfo)
            
            if result.returncode != 0:
                raise Exception(f"轉檔引擎失敗: {result.stderr}")
            
            # 處理 LibreOffice 輸出檔名對應
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

# ==================================
# UI 介面設計 (明亮清爽的 cosmo 主題)
# ==================================
window = tb.Window(
    title=APP_NAME,
    themename="cosmo",  # 明亮主題
    size=(700, 600)
)
window.resizable(False, False)

# 標題
title_label = tb.Label(
    window,
    text="文件專業轉檔工具 (支援 Pages/PDF/Word)",
    font=("Microsoft JhengHei UI", 16, "bold")
)
title_label.pack(pady=20)

# 選擇檔案按鈕與標籤
file_btn = tb.Button(
    window,
    text="選擇要轉換的檔案 (Pages/PDF/Word/TXT)",
    bootstyle="primary",
    command=choose_file,
    width=40
)
file_btn.pack(pady=5)

source_label = tb.Label(
    window,
    text="尚未選擇來源檔案",
    font=("Microsoft JhengHei UI", 10),
    bootstyle="secondary"
)
source_label.pack(pady=5)

# 選擇輸出資料夾按鈕與標籤
output_btn = tb.Button(
    window,
    text="選擇輸出資料夾",
    bootstyle="info",
    command=choose_output_folder,
    width=40
)
output_btn.pack(pady=5)

output_label = tb.Label(
    window,
    text="尚未選擇輸出資料夾",
    font=("Microsoft JhengHei UI", 10),
    bootstyle="secondary"
)
output_label.pack(pady=5)

# 目標格式選擇
tb.Label(window, text="目標輸出格式", font=("Microsoft JhengHei UI", 11, "bold")).pack(pady=(15, 5))
format_var = StringVar(value="docx")
format_menu = tb.Combobox(
    window,
    textvariable=format_var,
    values=["docx", "pdf", "txt"],
    state="readonly",
    width=20
)
format_menu.pack(pady=5)

# 狀態與進度條
status_label = tb.Label(window, text="待命中", font=("Microsoft JhengHei UI", 11))
status_label.pack(pady=10)

progress = tb.Progressbar(
    window,
    length=500,
    mode="indeterminate",
    bootstyle="success-striped"
)
progress.pack(pady=10)

# 開始轉換按鈕
start_btn = tb.Button(
    window,
    text="開始轉換",
    bootstyle="success",
    command=start_conversion_thread,
    width=20
)
start_btn.pack(pady=15)

window.mainloop()
