import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook, Workbook
import xlrd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class ExcelExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("表格提取器 ExcelDataMatcher v 1.1  - QwejayHuang")
        self.root.geometry("640x480")

        self.file_path = None
        self.sheet_names = []
        self.selected_sheet = tk.StringVar()
        self.search_all_sheets = tk.BooleanVar()
        self.header_row = tk.StringVar(value="1")  # 默认第一行是表头，显示为1
        self.no_header = tk.BooleanVar(value=True)  # 默认勾选不需要表头

        self.create_widgets()
        self.create_context_menu()  # 创建右键菜单

        # 初始化时设置输入框状态
        self.toggle_no_header()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件选择
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        self.file_label = ttk.Label(file_frame, text="未选择文件", width=40, anchor=tk.W)
        self.file_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="选择表格文件", command=self.select_file).pack(side=tk.RIGHT, padx=5)

        # 工作表选择
        sheet_frame = ttk.Frame(main_frame)
        sheet_frame.pack(fill=tk.X, pady=5)
        ttk.Label(sheet_frame, text="选择工作表:", width=12, anchor=tk.E).pack(side=tk.LEFT, padx=5)
        self.sheet_combobox = ttk.Combobox(sheet_frame, textvariable=self.selected_sheet, state='readonly', width=20)
        self.sheet_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Checkbutton(sheet_frame, text="搜索所有", variable=self.search_all_sheets, command=self.toggle_search_all).pack(side=tk.RIGHT, padx=5)

        # 表头选择
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=5)
        ttk.Label(header_frame, text="表头行:", width=12, anchor=tk.E).pack(side=tk.LEFT, padx=5)
        self.header_entry = ttk.Entry(header_frame, textvariable=self.header_row, width=5)
        self.header_entry.pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(header_frame, text="不需要表头", variable=self.no_header, command=self.toggle_no_header).pack(side=tk.LEFT, padx=5)

        # 列名输入标签
        ttk.Label(main_frame, text="输入列名（每列占一行）:", anchor=tk.W).pack(fill=tk.X, pady=5)

        # 列名输入
        self.column_text = tk.Text(main_frame, height=6, width=50)
        self.column_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 提取按钮
        ttk.Button(main_frame, text="提取并保存", command=self.extract_and_save, bootstyle=SUCCESS).pack(pady=10)

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=5)

    def create_context_menu(self):
        """创建右键菜单"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="粘贴", command=self.paste_text)
        self.context_menu.add_command(label="复制", command=self.copy_text)
        self.context_menu.add_command(label="清除", command=self.clear_text)

        # 绑定右键点击事件
        self.column_text.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        """显示右键菜单"""
        self.context_menu.post(event.x_root, event.y_root)

    def paste_text(self):
        """粘贴文本"""
        self.column_text.event_generate("<<Paste>>")
        self.status_var.set("已粘贴内容")

    def copy_text(self):
        """复制文本"""
        self.column_text.event_generate("<<Copy>>")
        self.status_var.set("已复制选中文本")

    def clear_text(self):
        """清除文本"""
        self.column_text.delete("1.0", tk.END)
        self.status_var.set("已清除文本框内容")

    def reset_file_selection(self):
        self.file_path = None
        self.file_label.config(text="未选择文件")
        self.sheet_names = []
        self.sheet_combobox['values'] = []
        self.selected_sheet.set("")
        self.status_var.set("操作取消: 未选择文件")

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path = file_path
            file_name = self.file_path.split("/")[-1]
            self.file_label.config(text=f"已选择文件: {file_name}")
            try:
                if self.file_path.endswith(".xlsx"):
                    workbook = load_workbook(self.file_path, read_only=True)
                    self.sheet_names = workbook.sheetnames
                elif self.file_path.endswith(".xls"):
                    workbook = xlrd.open_workbook(self.file_path)
                    self.sheet_names = workbook.sheet_names()
                self.sheet_combobox['values'] = self.sheet_names
                self.selected_sheet.set(self.sheet_names[0] if self.sheet_names else "")
                self.status_var.set(f"已加载: {file_name}")
            except Exception as e:
                self.status_var.set(f"错误: 读取Excel文件失败: {e}")
                self.reset_file_selection()
        else:
            if self.file_path is not None:
                self.status_var.set("操作取消: 请先选择表格文件")
            else:
                self.status_var.set("操作取消: 请重新选择文件")

    def toggle_search_all(self):
        self.sheet_combobox.config(state='disabled' if self.search_all_sheets.get() else 'readonly')
        self.status_var.set("已切换搜索模式")

    def toggle_no_header(self):
        if self.no_header.get():
            self.header_entry.config(state=tk.DISABLED)  # 禁用输入框
            self.header_row.set("None")  # 设置表头行为 None
        else:
            self.header_entry.config(state=tk.NORMAL)  # 启用输入框
            self.header_row.set("1")  # 设置表头行为 1

    def extract_and_save(self):
        if not self.file_path:
            self.status_var.set("错误: 请先选择表格文件")
            return

        column_names = self.column_text.get("1.0", "end-1c").splitlines()
        column_names = [col.strip() for col in column_names if col.strip()]

        if not column_names:
            self.status_var.set("错误: 列名不能为空")
            return

        header_row = self.header_row.get()
        if header_row == "None" or self.no_header.get():
            header_row = None
        else:
            try:
                header_row = int(header_row) - 1  # 转换为0索引
            except ValueError:
                self.status_var.set("错误: 表头行必须是整数或选择不需要表头")
                return

        try:
            extracted_data = self.extract_matching_rows(column_names, header_row)
            if not extracted_data:
                self.status_var.set("信息: 未找到任何匹配的行")
                return

            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])
            if save_path:
                wb = Workbook()
                ws = wb.active
                for row in extracted_data:
                    ws.append(row)
                wb.save(save_path)
                self.status_var.set(f"成功: 数据已保存到 {save_path}")
            else:
                self.status_var.set("操作取消: 未选择保存路径")
        except Exception as e:
            self.status_var.set(f"错误: 提取或保存失败: {e}")

    def extract_matching_rows(self, column_names, header_row):
        extracted_data = []
        if self.file_path.endswith(".xlsx"):
            workbook = load_workbook(self.file_path, read_only=True)
            if self.search_all_sheets.get():
                for sheet_name in self.sheet_names:
                    sheet = workbook[sheet_name]
                    data = list(sheet.iter_rows(values_only=True))
                    if header_row is not None:
                        headers = data[header_row]
                        data = data[header_row + 1:]
                    else:
                        headers = None
                    for row in data:
                        if any(str(cell).lower() in [col.lower() for col in column_names] for cell in row):
                            extracted_data.append(row)
            else:
                sheet = workbook[self.selected_sheet.get()]
                data = list(sheet.iter_rows(values_only=True))
                if header_row is not None:
                    headers = data[header_row]
                    data = data[header_row + 1:]
                else:
                    headers = None
                for row in data:
                    if any(str(cell).lower() in [col.lower() for col in column_names] for cell in row):
                        extracted_data.append(row)
        elif self.file_path.endswith(".xls"):
            workbook = xlrd.open_workbook(self.file_path)
            if self.search_all_sheets.get():
                for sheet_name in self.sheet_names:
                    sheet = workbook.sheet_by_name(sheet_name)
                    data = [sheet.row_values(row) for row in range(sheet.nrows)]
                    if header_row is not None:
                        headers = data[header_row]
                        data = data[header_row + 1:]
                    else:
                        headers = None
                    for row in data:
                        if any(str(cell).lower() in [col.lower() for col in column_names] for cell in row):
                            extracted_data.append(row)
            else:
                sheet = workbook.sheet_by_name(self.selected_sheet.get())
                data = [sheet.row_values(row) for row in range(sheet.nrows)]
                if header_row is not None:
                    headers = data[header_row]
                    data = data[header_row + 1:]
                else:
                    headers = None
                for row in data:
                    if any(str(cell).lower() in [col.lower() for col in column_names] for cell in row):
                        extracted_data.append(row)
        return extracted_data

if __name__ == "__main__":
    root = ttk.Window(themename="superhero")
    app = ExcelExtractorApp(root)
    root.mainloop()