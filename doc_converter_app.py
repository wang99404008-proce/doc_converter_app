import os
import sys
import zipfile
import shutil
import threading
from tkinter import filedialog, messagebox, StringVar
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from pdf2docx import Converter
import fitz  # PyMuPDF
import docx
from fpdf import FPDF

APP_NAME = "文件專業轉檔工具 (表格優化版)"

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
        extracted_pdf_path = None
        if ext == 'pages':
            temp_pdf_path = os.path.join(output_folder_path, f"temp_{base_name}.pdf")
            with zipfile.ZipFile(input_file_path, 'r') as z:
                pdf_internal_path = None
                for filename in z.namelist():
                    if filename.endswith('Preview.pdf') or filename.endswith('quicklook/Preview.pdf'):
                        pdf_internal_path = filename
                        break
                if pdf_internal_path:
                    with z.open(pdf_internal_path) as source, open(temp_pdf_path, "wb") as target:
                        shutil.copyfileobj(source, target)
                    extracted_pdf_path = temp_pdf_path
                else:
                    raise Exception("無法從此 Pages 檔案中讀取預覽內容")

        # 1. 轉成 docx
        if target_format == 'docx':
            source_pdf = extracted_pdf_path if ext == 'pages' else input_file_path
            if ext not in ['pdf', 'pages']:
                raise Exception("目前僅支援從 PDF 或 Pages 轉換為 Word (.docx)")
            
            cv = Converter(source_pdf)
            cv.convert(output_path, start=0, end=None)
            cv.close()

        # 2. 轉成 txt
        elif target_format == 'txt':
            text_content = ""
            if ext in ['pdf', 'pages']:
                source_pdf = extracted_pdf_path if ext == 'pages' else input_file_path
                doc = fitz.open(source_pdf)
                for page in doc:
                    text_content += page.get_text() + "\n"
                doc.close()
            elif ext == 'docx':
                doc = docx.Document(input_file_path)
                for element in doc.element.body:
                    if element.tag.endswith('p'):
                        text_content += element.text + "\n"
                    elif element.tag.endswith('tbl'):
                        for row in element.xpath('.//w:tr'):
                            row_text = " | ".join([cell.text.strip() for cell in row.xpath('.//w:tc')])
                            text_content += row_text + "\n"
            elif ext == 'txt':
                with open(input_file_path, "r", encoding="utf-8", errors="ignore") as f:
                    text_content = f.read()
            
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(text_content)

        # 3. 轉成 pdf (具備優化表格與網格繪製功能)
        elif target_format == 'pdf':
            if ext == 'pages':
                shutil.copy(extracted_pdf_path, output_path)
            elif ext in ['txt', 'docx']:
                pdf = FPDF()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()
                pdf.set_left_margin(10)
                pdf.set_right_margin(10)

                font_path = "C:/Windows/Fonts/msjh.ttc"
                if os.path.exists(font_path):
                    pdf.add_font("ChineseFont", "", font_path)
                    pdf.set_font("ChineseFont", size=10)
                else:
                    pdf.set_font("Arial", size=10)
                
                printable_width = 210 - 20  # A4 寬度扣除左右邊界

                if ext == 'txt':
                    with open(input_file_path, "r", encoding="utf-8", errors="ignore") as f:
                        text_content = f.read()
                    for line in text_content.split('\n'):
                        if line.strip():
                            pdf.multi_cell(printable_width, 7, txt=line)
                        else:
                            pdf.ln(4)
                else:
                    doc = docx.Document(input_file_path)
                    # 依序走訪 Word 中的段落與表格，確保順序正確
                    for element in doc.element.body:
                        if element.tag.endswith('p'):
                            p = docx.text.paragraph.Paragraph(element, doc)
                            if p.text.strip():
                                pdf.multi_cell(printable_width, 7, txt=p.text)
                            else:
                                pdf.ln(4)
                        elif element.tag.endswith('tbl'):
                            table = docx.table.Table(element, doc)
                            pdf.ln(3)
                            for row in table.rows:
                                cells_text = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                                num_cols = len(cells_text)
                                if num_cols == 0:
                                    continue
                                col_width = printable_width / num_cols
                                
                                # 繪製表格每一列的儲存格與框線
                                for text in cells_text:
                                    pdf.cell(col_width, 8, txt=text[:35], border=1, align="C")
                                pdf.ln()
                            pdf.ln(3)

                pdf.output(output_path)
            else:
                raise Exception(f"不支援從 .{ext} 轉換為 PDF")

        else:
            raise Exception("不支援此轉換格式組合")

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

# UI 介面設定
window = tb.Window(title=APP_NAME, themename="cosmo", size=(700, 600))
window.resizable(False, False)

tb.Label(window, text="文件專業轉檔工具 (表格優化版)", font=("Microsoft JhengHei UI", 16, "bold")).pack(pady=20)

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
