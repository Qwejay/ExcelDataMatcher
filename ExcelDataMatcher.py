import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class ExcelExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("表格提取器 ExcelDataMatcher v 1.0.2  - QwejayHuang")
        self.root.geometry("600x600")

        self.file_path = None
        self.sheet_names = []
        self.selected_sheet = tk.StringVar()
        self.search_all_sheets = tk.BooleanVar()
        self.header_row = tk.StringVar(value="1")  # 默认第一行是表头，显示为1
        self.no_header = tk.BooleanVar(value=True)  # 默认勾选不需要表头

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(padx=10, pady=10)

        self.create_file_selection_widgets(main_frame)
        self.create_sheet_selection_widgets(main_frame)
        self.create_header_selection_widgets(main_frame)
        self.create_column_input_widgets(main_frame)
        self.create_button_widgets(main_frame)
        self.create_log_widgets(main_frame)

    def create_file_selection_widgets(self, parent):
        file_frame = ttk.Frame(parent)
        file_frame.pack(fill=tk.X, pady=5)

        self.file_label = ttk.Label(file_frame, text="未选择文件")
        self.file_label.pack(side=tk.LEFT, padx=5)

        self.select_file_button = ttk.Button(file_frame, text="选择表格文件", command=self.select_file)
        self.select_file_button.pack(side=tk.RIGHT, padx=5)

    def create_sheet_selection_widgets(self, parent):
        sheet_frame = ttk.Frame(parent)
        sheet_frame.pack(fill=tk.X, pady=5)

        self.sheet_label = ttk.Label(sheet_frame, text="选择工作表:")
        self.sheet_label.pack(side=tk.LEFT, padx=5)

        self.sheet_combobox = ttk.Combobox(sheet_frame, textvariable=self.selected_sheet, state='readonly')
        self.sheet_combobox.pack(side=tk.LEFT, padx=5)

        self.search_all_checkbox = ttk.Checkbutton(sheet_frame, text="搜索所有工作表", variable=self.search_all_sheets, command=self.toggle_search_all)
        self.search_all_checkbox.pack(side=tk.LEFT, padx=5)

    def create_header_selection_widgets(self, parent):
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=5)

        self.header_label = ttk.Label(header_frame, text="表头行:")
        self.header_label.pack(side=tk.LEFT, padx=5)

        self.header_entry = ttk.Entry(header_frame, textvariable=self.header_row, width=5)
        self.header_entry.pack(side=tk.LEFT, padx=5)

        self.no_header_checkbox = ttk.Checkbutton(header_frame, text="不需要表头", variable=self.no_header, command=self.toggle_no_header)
        self.no_header_checkbox.pack(side=tk.LEFT, padx=5)

        # 初始化时根据 no_header 的值设置 header_entry 的状态
        self.toggle_no_header()

    def create_column_input_widgets(self, parent):
        column_frame = ttk.Frame(parent)
        column_frame.pack(fill=tk.X, pady=5)

        self.column_label = ttk.Label(column_frame, text="输入列名（每个列名占一行）:")
        self.column_label.pack(side=tk.LEFT, padx=5)

        self.column_text = tk.Text(parent, height=10)
        self.column_text.pack(fill=tk.X, pady=5)

    def create_button_widgets(self, parent):
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=5)

        self.extract_button = ttk.Button(button_frame, text="提取并保存", command=self.extract_and_save, state=tk.DISABLED)
        self.extract_button.pack(side=tk.RIGHT, padx=5)

    def create_log_widgets(self, parent):
        log_frame = ttk.Frame(parent)
        log_frame.pack(fill=tk.X, pady=5)

        self.log_text = tk.Text(log_frame, height=10, bg="gray90")
        self.log_text.pack(fill=tk.X, pady=5)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file_path:
            file_name = self.file_path.split("/")[-1]
            truncated_file_name = self.truncate_filename(file_name)
            self.file_label.config(text=f"已选择文件: {truncated_file_name}")
            self.log(f"已选择文件: {file_name}")
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

    def truncate_filename(self, filename, max_length=30):
        if len(filename) > max_length:
            return filename[:max_length - 3] + "..."
        return filename

    def toggle_search_all(self):
        self.sheet_combobox.config(state='disabled' if self.search_all_sheets.get() else 'readonly')

    def toggle_no_header(self):
        if self.no_header.get():
            self.header_entry.config(state=tk.DISABLED)
            self.header_row.set("None")
        else:
            self.header_entry.config(state=tk.NORMAL)
            self.header_row.set("1")  # 显示为1

    def extract_and_save(self):
        if not self.file_path:
            self.log("错误: 未选择文件")
            return

        column_names = self.column_text.get("1.0", "end-1c").splitlines()
        column_names = [col.strip() for col in column_names if col.strip()]

        if not column_names:
            self.log("错误: 列名不能为空")
            return

        header_row = self.header_row.get()
        if header_row == "None" or self.no_header.get():
            header_row = None
        else:
            try:
                header_row = int(header_row) - 1  # 转换为0索引
            except ValueError:
                self.log("错误: 表头行必须是整数或选择不需要表头")
                return

        try:
            extracted_rows = self.extract_matching_rows(column_names, header_row)

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

    def extract_matching_rows(self, column_names, header_row):
        extracted_rows = []
        if self.search_all_sheets.get():
            self.log("开始搜索所有工作表")
            for sheet_name in self.sheet_names:
                df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=header_row)
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
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=header_row)
            self.log(f"读取工作表: {sheet_name}")
            for column_name in column_names:
                matching_rows = df[df.apply(lambda row: row.astype(str).str.contains(column_name, case=False)).any(axis=1)]
                if not matching_rows.empty:
                    extracted_rows.append(matching_rows)
                    self.log(f"在 {sheet_name} 中提取 '{column_name}' 的行")
                else:
                    self.log(f"警告: 在 {sheet_name} 中不存在 '{column_name}' ")
        return extracted_rows

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

if __name__ == "__main__":
    root = ttk.Window(themename="superhero")
    app = ExcelExtractorApp(root)
    root.mainloop()
