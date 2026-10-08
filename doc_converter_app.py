import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, Text
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from pdf2docx import Converter

APP_NAME = "文件專業轉檔工具 (升級互動版)"

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

def set_input_file(file_path):
    global input_file_path
    if not file_path:
        return
    # 清理路徑字號（有時候拖放會帶引號）
    file_path = file_path.strip('"{ }')
    input_file_path = file_path
    source_label.config(text=f"已選來源檔案：\n{os.path.basename(file_path)}", bootstyle="success")
    path_display_box.config(state="normal")
    path_display_box.delete("1.0", tk.END)
    path_display_box.insert(tk.END, file_path)
    path_display_box.config(state="disabled")

def choose_file():
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
    if file_path:
        set_input_file(file_path)

def choose_output_folder():
    global output_folder_path
    folder_path = filedialog.askdirectory(title="選擇輸出資料夾")
    if not folder_path:
        return
    output_folder_path = folder_path
    output_label.config(text=f"輸出資料夾：\n{output_folder_path}", bootstyle="success")

def convert_document():
    global input_file_path, output_folder_path
    
    if not input_file_path or not output_folder_path:
        messagebox.showwarning("提醒", "請先完整選擇『來源檔案』與『輸出資料夾』！")
        return

    target_format = format_var.get().lower()
    base_name = os.path.splitext(os.path.basename(input_file_path))[0]
    ext = os.path.splitext(input_file_path)[1][1:].lower()
    
    output_filename = f"{base_name}_converted.{target_format}"
    output_path = os.path.join(output_folder_path, output_filename)

    status_label.config(text="正在轉檔中，請稍候...", bootstyle="warning")
    progress.start(10)
    window.update_idletasks()

    try:
        # A. PDF 轉 Word
        if ext == 'pdf' and target_format == 'docx':
            cv = Converter(input_file_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()

        # B. 透過本機旁邊的 LibreOffice 引擎處理
        elif target_format in ['pdf', 'docx', 'txt']:
            soffice_bin = get_libreoffice_path()
            if soffice_bin == "soffice":
                raise Exception("找不到獨立的 LibreOffice 引擎！\n請確保 'libreoffice_portable' 資料夾與 .exe 放在同一個資料夾下。")

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
        status_label.config(text="轉換完成！", bootstyle="success")
        
        # 顯示完整輸出路徑在介面上
        final_output_box.config(state="normal")
        final_output_box.delete("1.0", tk.END)
        final_output_box.insert(tk.END, output_path)
        final_output_box.config(state="disabled")
        
        messagebox.showinfo("成功", f"檔案已成功轉換並儲存至：\n{output_path}")

    except Exception as e:
        progress.stop()
        status_label.config(text="轉換失敗", bootstyle="danger")
        messagebox.showerror("錯誤", f"轉檔過程發生錯誤：\n{str(e)}")

def start_conversion_thread():
    threading.Thread(target=convert_document, daemon=True).start()

# --- UI 介面設計 ---
window = tb.Window(title=APP_NAME, themename="cosmo", size=(750, 720))
window.resizable(False, False)

# 標題
tb.Label(window, text="📄 文件專業轉檔工具 (支援拖拉與路徑顯示)", font=("Microsoft JhengHei UI", 15, "bold")).pack(pady=10)

# 上半部：檔案選擇與拖拉區
drop_frame = tb.Labelframe(window, text=" 檔案來源 (可點擊按鈕或直接將檔案拖曳至下方框內) ", padding=10, bootstyle="primary")
drop_frame.pack(fill=X, padx=20, pady=5)

btn_file = tb.Button(drop_frame, text="📁 點擊選擇檔案 (PDF/Word/Pages/TXT)", bootstyle="primary-outline", command=choose_file, width=45)
btn_file.pack(pady=5)

source_label = tb.Label(drop_frame, text="尚未選擇來源檔案", font=("Microsoft JhengHei UI", 9), bootstyle="secondary")
source_label.pack(pady=2)

# 顯示完整來源路徑框
path_display_box = Text(drop_frame, height=2, width=80, font=("Consolas", 9), state="disabled", bg="#f8f9fa")
path_display_box.pack(pady=5)

# 中間區：輸出資料夾與預覽/圖片顯示區
output_frame = tb.Labelframe(window, text=" 輸出設定與預覽區 ", padding=10, bootstyle="info")
output_frame.pack(fill=X, padx=20, pady=5)

btn_folder = tb.Button(output_frame, text="📂 選擇輸出資料夾", bootstyle="info-outline", command=choose_output_folder, width=30)
btn_folder.pack(pady=5)

output_label = tb.Label(output_frame, text="尚未選擇輸出資料夾", font=("Microsoft JhengHei UI", 9), bootstyle="secondary")
output_label.pack(pady=2)

# 保留的圖片/狀態預覽顯示小框（可用於未來顯示圖示或縮圖）
preview_canvas = tk.Canvas(output_frame, width=120, height=80, bg="#e9ecef", highlightthickness=1, highlightbackground="#ced4da")
preview_canvas.pack(pady=5)
preview_canvas.create_text(60, 40, text="[ 檔案圖示預覽區 ]", fill="#6c757d", font=("Microsoft JhengHei UI", 8))

# 轉檔選項與按鈕區
settings_frame = tb.Frame(window, padding=10)
settings_frame.pack(fill=X, padx=20)

tb.Label(settings_frame, text="目標輸出格式:", font=("Microsoft JhengHei UI", 10, "bold")).pack(side=LEFT, padx=5)
format_var = StringVar(value="pdf")
tb.Combobox(settings_frame, textvariable=format_var, values=["pdf", "docx", "txt"], state="readonly", width=12).pack(side=LEFT, padx=5)

status_label = tb.Label(settings_frame, text="待命中", font=("Microsoft JhengHei UI", 10, "bold"), bootstyle="info")
status_label.pack(side=RIGHT, padx=10)

# 進度條
progress = tb.Progressbar(window, length=700, mode="indeterminate", bootstyle="success-striped")
progress.pack(pady=5)

# 執行按鈕與最終輸出路徑顯示
action_frame = tb.Frame(window, padding=10)
action_frame.pack(fill=X, padx=20)

btn_start = tb.Button(action_frame, text="🚀 開始轉換文件", bootstyle="success", command=start_conversion_thread, width=30)
btn_start.pack(pady=5)

tb.Label(window, text="最終輸出檔案完整路徑：", font=("Microsoft JhengHei UI", 9)).pack(anchor="w", padx=25)
final_output_box = Text(window, height=2, width=90, font=("Consolas", 9), state="disabled", bg="#e2f0d9")
final_output_box.pack(padx=20, pady=5)

window.mainloop()
