"""
簡易檔案格式轉換工具 (Windows)
支援：
  - JPG/PNG/Word(docx)/Excel(xlsx) -> PDF
  - PDF -> JPG / Word(docx)
  - PDF -> 拆成單頁 PDF
  - 多個 PDF -> 合併成一個 PDF

需求套件 (requirements.txt)：
  pillow
  docx2pdf
  pywin32
  PyMuPDF
  pdf2docx
  pypdf

安裝方式：
  pip install -r requirements.txt

備註：
  - Word/Excel 轉 PDF 需要電腦有安裝 Microsoft Word / Excel (透過 COM 呼叫)。
  - 其餘轉換為純 Python 套件，不需要額外安裝其他軟體。
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ---------------------------------------------------------------------------
# 轉換功能
# ---------------------------------------------------------------------------

def convert_image_to_pdf(image_path, output_path):
    from PIL import Image
    img = Image.open(image_path)
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.save(output_path, "PDF")


def convert_word_to_pdf(docx_path, output_path):
    from docx2pdf import convert
    convert(docx_path, output_path)


def convert_excel_to_pdf(xlsx_path, output_path):
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
        # 0 = xlTypePDF
        wb.ExportAsFixedFormat(0, os.path.abspath(output_path))
        wb.Close(False)
    finally:
        excel.Quit()


def convert_pdf_to_images(pdf_path, output_dir):
    import fitz  # PyMuPDF

    base = os.path.splitext(os.path.basename(pdf_path))[0]
    doc = fitz.open(pdf_path)
    saved = []
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=200)
        out_path = os.path.join(output_dir, f"{base}_p{i}.jpg")
        pix.save(out_path)
        saved.append(out_path)
    doc.close()
    return saved


def convert_pdf_to_word(pdf_path, output_path):
    from pdf2docx import Converter

    cv = Converter(pdf_path)
    cv.convert(output_path)
    cv.close()


def split_pdf_to_single_pages(pdf_path, output_dir):
    from pypdf import PdfReader, PdfWriter

    base = os.path.splitext(os.path.basename(pdf_path))[0]
    reader = PdfReader(pdf_path)
    saved = []
    for i, page in enumerate(reader.pages, start=1):
        writer = PdfWriter()
        writer.add_page(page)
        out_path = os.path.join(output_dir, f"{base}_p{i}.pdf")
        with open(out_path, "wb") as f:
            writer.write(f)
        saved.append(out_path)
    return saved


def merge_pdfs(pdf_paths, output_path):
    from pypdf import PdfWriter

    writer = PdfWriter()
    for p in pdf_paths:
        writer.append(p)
    with open(output_path, "wb") as f:
        writer.write(f)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

IMAGE_EXT = (".jpg", ".jpeg", ".png")
WORD_EXT = (".docx",)
EXCEL_EXT = (".xlsx", ".xls")
PDF_EXT = (".pdf",)


class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("檔案格式轉換工具")
        self.geometry("520x380")
        self.resizable(False, False)

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_convert = ttk.Frame(notebook)
        self.tab_pdf_tools = ttk.Frame(notebook)
        notebook.add(self.tab_convert, text="格式轉換")
        notebook.add(self.tab_pdf_tools, text="PDF 合併 / 拆分")

        self._build_convert_tab()
        self._build_pdf_tools_tab()

        self.selected_file = None

    # -- Tab 1: 單檔轉換 ----------------------------------------------------
    def _build_convert_tab(self):
        frame = self.tab_convert

        ttk.Label(frame, text="步驟 1：選擇要轉換的檔案").pack(anchor="w", padx=10, pady=(15, 5))

        row = ttk.Frame(frame)
        row.pack(fill="x", padx=10)
        self.file_entry = ttk.Entry(row, state="readonly")
        self.file_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="選擇檔案...", command=self._choose_file).pack(side="left", padx=5)

        ttk.Label(frame, text="步驟 2：選擇要轉換成的格式").pack(anchor="w", padx=10, pady=(20, 5))
        self.target_var = tk.StringVar()
        self.target_combo = ttk.Combobox(frame, textvariable=self.target_var, state="readonly")
        self.target_combo.pack(fill="x", padx=10)

        ttk.Button(frame, text="開始轉換", command=self._start_convert).pack(pady=20)

        self.status_label = ttk.Label(frame, text="", foreground="blue")
        self.status_label.pack(padx=10, anchor="w")

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="選擇檔案",
            filetypes=[
                ("支援的檔案", "*.jpg *.jpeg *.png *.docx *.xlsx *.xls *.pdf"),
                ("所有檔案", "*.*"),
            ],
        )
        if not path:
            return

        self.selected_file = path
        self.file_entry.config(state="normal")
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, path)
        self.file_entry.config(state="readonly")

        ext = os.path.splitext(path)[1].lower()
        if ext in IMAGE_EXT:
            options = ["PDF"]
        elif ext in WORD_EXT:
            options = ["PDF"]
        elif ext in EXCEL_EXT:
            options = ["PDF"]
        elif ext in PDF_EXT:
            options = ["JPG 圖片", "Word (docx)"]
        else:
            options = []
            messagebox.showwarning("不支援的格式", f"不支援的副檔名：{ext}")

        self.target_combo["values"] = options
        self.target_var.set(options[0] if options else "")

    def _start_convert(self):
        if not self.selected_file:
            messagebox.showwarning("尚未選擇檔案", "請先選擇要轉換的檔案")
            return
        target = self.target_var.get()
        if not target:
            messagebox.showwarning("尚未選擇格式", "請選擇要轉換成的格式")
            return

        threading.Thread(target=self._run_convert, args=(self.selected_file, target), daemon=True).start()

    def _run_convert(self, path, target):
        self._set_status("轉換中，請稍候...")
        try:
            ext = os.path.splitext(path)[1].lower()

            if target == "PDF":
                default_name = os.path.splitext(os.path.basename(path))[0] + ".pdf"
                out_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf", initialfile=default_name,
                    filetypes=[("PDF 檔案", "*.pdf")],
                )
                if not out_path:
                    self._set_status("")
                    return

                if ext in IMAGE_EXT:
                    convert_image_to_pdf(path, out_path)
                elif ext in WORD_EXT:
                    convert_word_to_pdf(path, out_path)
                elif ext in EXCEL_EXT:
                    convert_excel_to_pdf(path, out_path)

                self._set_status(f"完成！已儲存至：{out_path}")
                messagebox.showinfo("完成", f"轉換完成：\n{out_path}")

            elif target == "JPG 圖片":
                out_dir = filedialog.askdirectory(title="選擇圖片輸出資料夾")
                if not out_dir:
                    self._set_status("")
                    return
                saved = convert_pdf_to_images(path, out_dir)
                self._set_status(f"完成！共輸出 {len(saved)} 張圖片")
                messagebox.showinfo("完成", f"已輸出 {len(saved)} 張圖片至：\n{out_dir}")

            elif target == "Word (docx)":
                default_name = os.path.splitext(os.path.basename(path))[0] + ".docx"
                out_path = filedialog.asksaveasfilename(
                    defaultextension=".docx", initialfile=default_name,
                    filetypes=[("Word 檔案", "*.docx")],
                )
                if not out_path:
                    self._set_status("")
                    return
                convert_pdf_to_word(path, out_path)
                self._set_status(f"完成！已儲存至：{out_path}")
                messagebox.showinfo("完成", f"轉換完成：\n{out_path}")

        except Exception as e:
            self._set_status("轉換失敗")
            messagebox.showerror("轉換失敗", str(e))

    def _set_status(self, text):
        self.status_label.config(text=text)

    # -- Tab 2: PDF 合併 / 拆分 ----------------------------------------------
    def _build_pdf_tools_tab(self):
        frame = self.tab_pdf_tools

        # 拆分
        split_box = ttk.LabelFrame(frame, text="拆成單頁 PDF")
        split_box.pack(fill="x", padx=10, pady=(15, 10))
        ttk.Label(split_box, text="選擇一個 PDF，將每一頁拆成獨立的 PDF 檔案").pack(
            anchor="w", padx=10, pady=(5, 5)
        )
        ttk.Button(split_box, text="選擇 PDF 並拆分...", command=self._do_split).pack(
            anchor="w", padx=10, pady=(0, 10)
        )

        # 合併
        merge_box = ttk.LabelFrame(frame, text="合併多個 PDF")
        merge_box.pack(fill="both", expand=True, padx=10, pady=10)
        ttk.Label(merge_box, text="選擇多個 PDF 檔案（依選取順序合併）").pack(
            anchor="w", padx=10, pady=(5, 5)
        )

        list_row = ttk.Frame(merge_box)
        list_row.pack(fill="both", expand=True, padx=10)
        self.merge_listbox = tk.Listbox(list_row, height=6)
        self.merge_listbox.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_row, command=self.merge_listbox.yview)
        scrollbar.pack(side="left", fill="y")
        self.merge_listbox.config(yscrollcommand=scrollbar.set)

        btn_row = ttk.Frame(merge_box)
        btn_row.pack(fill="x", padx=10, pady=10)
        ttk.Button(btn_row, text="加入 PDF...", command=self._add_merge_files).pack(side="left")
        ttk.Button(btn_row, text="清空清單", command=self._clear_merge_files).pack(side="left", padx=5)
        ttk.Button(btn_row, text="開始合併...", command=self._do_merge).pack(side="right")

        self.merge_files = []

    def _do_split(self):
        path = filedialog.askopenfilename(title="選擇要拆分的 PDF", filetypes=[("PDF 檔案", "*.pdf")])
        if not path:
            return
        out_dir = filedialog.askdirectory(title="選擇輸出資料夾")
        if not out_dir:
            return
        try:
            saved = split_pdf_to_single_pages(path, out_dir)
            messagebox.showinfo("完成", f"已拆分為 {len(saved)} 個檔案，輸出至：\n{out_dir}")
        except Exception as e:
            messagebox.showerror("拆分失敗", str(e))

    def _add_merge_files(self):
        paths = filedialog.askopenfilenames(title="選擇 PDF 檔案（可多選）", filetypes=[("PDF 檔案", "*.pdf")])
        for p in paths:
            self.merge_files.append(p)
            self.merge_listbox.insert(tk.END, os.path.basename(p))

    def _clear_merge_files(self):
        self.merge_files = []
        self.merge_listbox.delete(0, tk.END)

    def _do_merge(self):
        if len(self.merge_files) < 2:
            messagebox.showwarning("檔案不足", "請至少選擇兩個 PDF 檔案進行合併")
            return
        out_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", initialfile="merged.pdf",
            filetypes=[("PDF 檔案", "*.pdf")],
        )
        if not out_path:
            return
        try:
            merge_pdfs(self.merge_files, out_path)
            messagebox.showinfo("完成", f"合併完成：\n{out_path}")
        except Exception as e:
            messagebox.showerror("合併失敗", str(e))


if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()
