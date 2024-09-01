import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class ExcelExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("表格提取器 ExcelDataMatcher v 1.0.1  - QwejayHhuang")
        self.root.geometry("600x600")  # 设置窗口大小

        self.file_path = None
        self.sheet_names = []
        self.selected_sheet = tk.StringVar()
        self.search_all_sheets = tk.BooleanVar()

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(padx=10, pady=10)

        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)

        self.file_label = ttk.Label(file_frame, text="未选择文件")
        self.file_label.pack(side=tk.LEFT, padx=5)

        self.select_file_button = ttk.Button(file_frame, text="选择表格文件", command=self.select_file)
        self.select_file_button.pack(side=tk.RIGHT, padx=5)

        sheet_frame = ttk.Frame(main_frame)
        sheet_frame.pack(fill=tk.X, pady=5)

        self.sheet_label = ttk.Label(sheet_frame, text="选择工作表:")
        self.sheet_label.pack(side=tk.LEFT, padx=5)

        self.sheet_combobox = ttk.Combobox(sheet_frame, textvariable=self.selected_sheet, state='readonly')
        self.sheet_combobox.pack(side=tk.LEFT, padx=5)

        self.search_all_checkbox = ttk.Checkbutton(sheet_frame, text="搜索所有工作表", variable=self.search_all_sheets, command=self.toggle_search_all)
        self.search_all_checkbox.pack(side=tk.LEFT, padx=5)

        column_frame = ttk.Frame(main_frame)
        column_frame.pack(fill=tk.X, pady=5)

        self.column_label = ttk.Label(column_frame, text="请输入单元格数据或关键词，每行一项，支持批量输入：")
        self.column_label.pack(side=tk.LEFT, padx=5)

        self.column_text = tk.Text(main_frame, height=10)
        self.column_text.pack(fill=tk.X, pady=5)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)

        self.extract_button = ttk.Button(button_frame, text="提取并保存", command=self.extract_and_save, state=tk.DISABLED)
        self.extract_button.pack(side=tk.RIGHT, padx=5)

        log_frame = ttk.Frame(main_frame)
        log_frame.pack(fill=tk.X, pady=5)

        self.log_text = tk.Text(log_frame, height=10, bg="gray90")
        self.log_text.pack(fill=tk.X, pady=5)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file_path:
            self.file_label.config(text=f"已选择文件: {self.file_path}")
            self.log(f"已选择文件: {self.file_path}")
            try:
                self.sheet_names = pd.ExcelFile(self.file_path).sheet_names
                self.sheet_combobox['values'] = self.sheet_names
                self.selected_sheet.set(self.sheet_names[0])
                self.extract_button.config(state=tk.NORMAL)
            except Exception as e:
                self.log(f"错误: 读取Excel文件失败: {e}")
                self.extract_button.config(state=tk.DISABLED)
        else:
            self.log("未选择文件")
            self.extract_button.config(state=tk.DISABLED)

    def toggle_search_all(self):
        if self.search_all_sheets.get():
            self.sheet_combobox.config(state=tk.DISABLED)
        else:
            self.sheet_combobox.config(state=tk.NORMAL)

    def extract_and_save(self):
        if not self.file_path:
            self.log("错误: 未选择文件")
            return

        column_names = self.column_text.get("1.0", "end-1c").splitlines()
        column_names = [col.strip() for col in column_names if col.strip()]

        if not column_names:
            self.log("错误: 列名不能为空")
            return

        try:
            if self.search_all_sheets.get():
                self.log("开始搜索所有工作表")
                extracted_rows = []
                for sheet_name in self.sheet_names:
                    df = pd.read_excel(self.file_path, sheet_name=sheet_name)
                    self.log(f"读取工作表: {sheet_name}")
                    for column_name in column_names:
                        matching_rows = df[df.apply(lambda row: row.astype(str).str.contains(column_name, case=False)).any(axis=1)]
                        if not matching_rows.empty:
                            extracted_rows.append(matching_rows)
                            self.log(f"在 {sheet_name} 中提取 '{column_name}' 的行")
                        else:
                            self.log(f"警告: 在 {sheet_name} 中不存在 '{column_name}' ")
            else:
                sheet_name = self.selected_sheet.get()
                df = pd.read_excel(self.file_path, sheet_name=sheet_name)
                self.log(f"读取工作表: {sheet_name}")
                extracted_rows = []
                for column_name in column_names:
                    matching_rows = df[df.apply(lambda row: row.astype(str).str.contains(column_name, case=False)).any(axis=1)]
                    if not matching_rows.empty:
                        extracted_rows.append(matching_rows)
                        self.log(f"在 {sheet_name} 中提取 '{column_name}' 的行")
                    else:
                        self.log(f"警告: 在 {sheet_name} 中不存在 '{column_name}' ")

            if not extracted_rows:
                self.log("信息: 未找到任何匹配的行")
                return

            extracted_df = pd.concat(extracted_rows).drop_duplicates()

            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])
            if save_path:
                try:
                    extracted_df.to_excel(save_path, index=False)
                    self.log(f"成功: 数据已保存到 {save_path}")
                except Exception as e:
                    self.log(f"错误: 保存文件失败: {e}")

        except Exception as e:
            self.log(f"错误: 读取Excel文件失败: {e}")

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

if __name__ == "__main__":
    root = ttk.Window(themename="superhero")  # 使用 "cosmo" 主题
    app = ExcelExtractorApp(root)
    root.mainloop()