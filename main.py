import os
import re
import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.tooltip import ToolTip
from openpyxl import load_workbook, Workbook
import xlrd

class ExcelExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("表格数据提取器 ExcelDataMatcher 1.2")
        self.root.geometry("880x720")
        self.root.minsize(800, 650)

        self.file_path = None
        self.sheet_names = []
        self.extracted_data = []
        self.is_processing = False

        self.selected_sheet = tk.StringVar()
        self.search_all_sheets = tk.BooleanVar(value=False)
        self.header_row = tk.StringVar(value="1")
        self.no_header = tk.BooleanVar(value=False)
        
        self.remove_duplicates_var = tk.BooleanVar(value=False)
        self.remove_inner_empty_var = tk.BooleanVar(value=True)
        self.remove_empty_rows_var = tk.BooleanVar(value=True)
        self.include_source_sheet_var = tk.BooleanVar(value=True)

        self.match_mode = tk.StringVar(value="contains")
        self.match_logic = tk.StringVar(value="OR")
        self.case_sensitive = tk.BooleanVar(value=False)
        self.target_columns = tk.StringVar(value="")

        self.msg_queue = queue.Queue()

        self.create_widgets()
        self.create_context_menu()
        self.toggle_no_header()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.root.after(100, self.process_queue)

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main_frame, text=" 1. 文件与工作表选择 ", padding=10)
        file_frame.pack(fill=tk.X, pady=4)

        f_top = ttk.Frame(file_frame)
        f_top.pack(fill=tk.X, pady=(0, 6))
        self.file_label = ttk.Label(f_top, text="未选择文件", font=("Helvetica", 9, "italic"))
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(f_top, text="选择表格文件", command=self.select_file, bootstyle=PRIMARY).pack(side=tk.RIGHT, padx=5)

        f_grid = ttk.Frame(file_frame)
        f_grid.pack(fill=tk.X, pady=2)

        ttk.Label(f_grid, text="目标工作表:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=3)
        self.sheet_combobox = ttk.Combobox(f_grid, textvariable=self.selected_sheet, state='readonly', width=22)
        self.sheet_combobox.grid(row=0, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Checkbutton(f_grid, text="搜索所有Sheet", variable=self.search_all_sheets, command=self.toggle_search_all).grid(row=0, column=2, sticky=tk.W, padx=15, pady=3)

        ttk.Label(f_grid, text="表头所在行:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=3)
        self.header_entry = ttk.Entry(f_grid, textvariable=self.header_row, width=8)
        self.header_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=3)
        ttk.Checkbutton(f_grid, text="无表头数据", variable=self.no_header, command=self.toggle_no_header).grid(row=1, column=2, sticky=tk.W, padx=15, pady=3)

        rule_frame = ttk.LabelFrame(main_frame, text=" 2. 提取与匹配规则 ", padding=10)
        rule_frame.pack(fill=tk.X, pady=4)

        r_grid = ttk.Frame(rule_frame)
        r_grid.pack(fill=tk.X, pady=2)

        ttk.Label(r_grid, text="匹配模式:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=3)
        m_frame = ttk.Frame(r_grid)
        m_frame.grid(row=0, column=1, columnspan=3, sticky=tk.W)
        ttk.Radiobutton(m_frame, text="包含(模糊)", value="contains", variable=self.match_mode).pack(side=tk.LEFT, padx=(5, 10))
        ttk.Radiobutton(m_frame, text="精确匹配", value="exact", variable=self.match_mode).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(m_frame, text="正则匹配", value="regex", variable=self.match_mode).pack(side=tk.LEFT, padx=10)

        ttk.Label(r_grid, text="逻辑关系:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=3)
        l_frame = ttk.Frame(r_grid)
        l_frame.grid(row=1, column=1, columnspan=3, sticky=tk.W)
        ttk.Radiobutton(l_frame, text="满足任一关键词(OR)", value="OR", variable=self.match_logic).pack(side=tk.LEFT, padx=(5, 10))
        ttk.Radiobutton(l_frame, text="满足所有关键词(AND)", value="AND", variable=self.match_logic).pack(side=tk.LEFT, padx=10)

        ttk.Checkbutton(r_grid, text="区分大小写", variable=self.case_sensitive).grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5, pady=3)
        
        ttk.Label(r_grid, text="指定查找列(可选):").grid(row=2, column=2, sticky=tk.E, padx=(20, 5), pady=3)
        self.col_entry = ttk.Entry(r_grid, textvariable=self.target_columns, width=15)
        self.col_entry.grid(row=2, column=3, sticky=tk.W, padx=5, pady=3)
        ToolTip(self.col_entry, text="为空则查找全部列。可输入字母或数字，如：A, C 或 1, 3")

        middle_frame = ttk.Frame(main_frame)
        middle_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        kw_frame = ttk.LabelFrame(middle_frame, text=" 关键词/正则列表（每行一个） ", padding=5)
        kw_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        self.column_text = tk.Text(kw_frame, height=5, width=35)
        self.column_text.pack(fill=tk.BOTH, expand=True)

        clean_frame = ttk.LabelFrame(middle_frame, text=" 数据清洗与输出设置 ", padding=8)
        clean_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))

        cb1 = ttk.Checkbutton(clean_frame, text="去除单元格内换行符/空行", variable=self.remove_inner_empty_var)
        cb1.pack(anchor=tk.W, pady=3)

        cb2 = ttk.Checkbutton(clean_frame, text="移除整行空数据", variable=self.remove_empty_rows_var)
        cb2.pack(anchor=tk.W, pady=3)

        cb3 = ttk.Checkbutton(clean_frame, text="去除重复匹配行", variable=self.remove_duplicates_var)
        cb3.pack(anchor=tk.W, pady=3)

        cb4 = ttk.Checkbutton(clean_frame, text="输出结果添加来源Sheet列", variable=self.include_source_sheet_var)
        cb4.pack(anchor=tk.W, pady=3)

        self.btn_run = ttk.Button(clean_frame, text=" 开始提取数据 ", command=self.start_extract_thread, bootstyle=SUCCESS, width=20)
        self.btn_run.pack(pady=(15, 5))

        preview_frame = ttk.LabelFrame(main_frame, text=" 数据提取预览 (最多显示前100条) ", padding=5)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=4)

        self.tree = ttk.Treeview(preview_frame, show="headings", height=6)
        tree_scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=4)

        self.progress_bar = ttk.Progressbar(bottom_frame, mode='indeterminate', bootstyle=STRIPED)
        self.progress_bar.pack(fill=tk.X, pady=2)

        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(bottom_frame, textvariable=self.status_var, anchor=tk.W, font=("Helvetica", 9))
        status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.save_btn = ttk.Button(bottom_frame, text="导出结果到Excel", command=self.save_to_excel, bootstyle=INFO, state=tk.DISABLED)
        self.save_btn.pack(side=tk.RIGHT, padx=5)

    def create_context_menu(self):
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="粘贴", command=lambda: self.column_text.event_generate("<<Paste>>"))
        self.context_menu.add_command(label="复制", command=lambda: self.column_text.event_generate("<<Copy>>"))
        self.context_menu.add_command(label="清空", command=lambda: self.column_text.delete("1.0", tk.END))
        self.column_text.bind("<Button-3>", lambda e: self.context_menu.post(e.x_root, e.y_root))

    def select_file(self):
        if self.is_processing:
            messagebox.showwarning("警告", "后台正在提取数据，请勿切换文件！")
            return

        file_path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path), font=("Helvetica", 9, "bold"))
            try:
                if file_path.endswith(".xlsx"):
                    wb = load_workbook(file_path, read_only=True)
                    self.sheet_names = wb.sheetnames
                    wb.close()
                else:
                    wb = xlrd.open_workbook(file_path)
                    self.sheet_names = wb.sheet_names()
                
                if not self.sheet_names:
                    raise ValueError("该 Excel 文件中没有包含任何有效的工作表！")

                self.sheet_combobox['values'] = self.sheet_names
                self.selected_sheet.set(self.sheet_names[0])
                self.status_var.set("文件加载成功！")
            except Exception as e:
                self.status_var.set(f"文件读取错误: {str(e)}")
                messagebox.showerror("文件读取失败", f"无法解析该文件，请检查是否损坏或被占用:\n{str(e)}")
                self.file_path = None
                self.sheet_combobox['values'] = []
                self.selected_sheet.set("")

    def toggle_search_all(self):
        self.sheet_combobox.config(state='disabled' if self.search_all_sheets.get() else 'readonly')

    def toggle_no_header(self):
        if self.no_header.get():
            self.header_entry.config(state=tk.DISABLED)
        else:
            self.header_entry.config(state=tk.NORMAL)

    def parse_target_columns(self):
        raw = self.target_columns.get().strip()
        if not raw:
            return None
        cols = []
        for item in re.split(r'[,，\s]+', raw):
            if not item:
                continue
            if item.isdigit():
                val = int(item) - 1
                if val >= 0:
                    cols.append(val)
            elif item.isalpha():
                idx = 0
                for char in item.upper():
                    idx = idx * 26 + (ord(char) - ord('A')) + 1
                cols.append(idx - 1)
        return list(set(cols)) if cols else None

    def parse_header_row(self):
        if self.no_header.get():
            return None
        try:
            val = int(self.header_row.get().strip())
            return max(0, val - 1)
        except ValueError:
            return 0 

    def start_extract_thread(self):
        if self.is_processing:
            return

        if not self.file_path or not os.path.exists(self.file_path):
            messagebox.showwarning("提示", "请先选择有效的 Excel 文件！")
            return

        keywords = [col.strip() for col in self.column_text.get("1.0", "end-1c").splitlines() if col.strip()]
        if not keywords:
            messagebox.showwarning("提示", "请输入至少一个关键词或正则表达规则！")
            return

        if self.match_mode.get() == "regex":
            for kw in keywords:
                try:
                    re.compile(kw)
                except re.error as e:
                    messagebox.showerror("正则表达式语法错误", f"规则 [{kw}] 不合法:\n{str(e)}")
                    return

        self.is_processing = True
        self.btn_run.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self.progress_bar.start(10)
        self.status_var.set("正在提取数据中，请稍候...")

        threading.Thread(target=self.run_extraction, args=(keywords,), daemon=True).start()

    def run_extraction(self, keywords):
        try:
            header_idx = self.parse_header_row()
            sheets_to_read = self.sheet_names if self.search_all_sheets.get() else [self.selected_sheet.get()]
            target_cols = self.parse_target_columns()
            
            extracted_results = []
            
            for sheet_name in sheets_to_read:
                rows = self.read_sheet_data(sheet_name)
                if not rows:
                    continue

                if header_idx is not None:
                    if header_idx >= len(rows):
                        continue
                    data_rows = rows[header_idx + 1:]
                else:
                    data_rows = rows

                for row in data_rows:
                    if self.is_row_matched(row, keywords, target_cols):
                        row_list = list(row)
                        if self.include_source_sheet_var.get():
                            row_list.insert(0, sheet_name)
                        extracted_results.append(row_list)

            if self.remove_inner_empty_var.get():
                extracted_results = self.clean_inner_empty(extracted_results)
            
            if self.remove_empty_rows_var.get():
                extracted_results = [r for r in extracted_results if any(c is not None and str(c).strip() != "" for c in r)]

            if self.remove_duplicates_var.get():
                extracted_results = self.remove_duplicates(extracted_results)

            self.msg_queue.put(("SUCCESS", extracted_results))
        except Exception as e:
            self.msg_queue.put(("ERROR", str(e)))

    def read_sheet_data(self, sheet_name):
        data = []
        try:
            if self.file_path.endswith(".xlsx"):
                wb = load_workbook(self.file_path, read_only=True, data_only=True)
                if sheet_name in wb.sheetnames:
                    sheet = wb[sheet_name]
                    data = list(sheet.iter_rows(values_only=True))
                wb.close()
            else:
                wb = xlrd.open_workbook(self.file_path)
                if sheet_name in wb.sheet_names():
                    sheet = wb.sheet_by_name(sheet_name)
                    data = [sheet.row_values(r) for r in range(sheet.nrows)]
        except Exception as e:
            print(f"读取 Sheet [{sheet_name}] 警告: {e}")
        return data

    def is_row_matched(self, row, keywords, target_cols):
        mode = self.match_mode.get()
        logic = self.match_logic.get()
        case_sen = self.case_sensitive.get()

        cells_to_check = []
        row_len = len(row)

        if target_cols is not None:
            for idx in target_cols:
                if idx < row_len and row[idx] is not None:
                    cells_to_check.append(str(row[idx]))
        else:
            for cell in row:
                if cell is not None:
                    cells_to_check.append(str(cell))

        if not cells_to_check:
            return False

        kw_matches = []
        for kw in keywords:
            kw_hit = False
            for cell_str in cells_to_check:
                try:
                    if mode == "exact":
                        matched = (cell_str == kw) if case_sen else (cell_str.lower() == kw.lower())
                    elif mode == "contains":
                        matched = (kw in cell_str) if case_sen else (kw.lower() in cell_str.lower())
                    elif mode == "regex":
                        flags = 0 if case_sen else re.IGNORECASE
                        matched = bool(re.search(kw, cell_str, flags))
                    else:
                        matched = False

                    if matched:
                        kw_hit = True
                        break
                except Exception:
                    continue

            kw_matches.append(kw_hit)

        return any(kw_matches) if logic == "OR" else all(kw_matches)

    def clean_inner_empty(self, data):
        cleaned = []
        for row in data:
            cleaned_row = []
            for cell in row:
                if isinstance(cell, str):
                    lines = [line.strip() for line in cell.splitlines() if line.strip()]
                    cleaned_row.append("\n".join(lines))
                else:
                    cleaned_row.append(cell)
            cleaned.append(cleaned_row)
        return cleaned

    def remove_duplicates(self, data):
        seen = set()
        unique_data = []
        for row in data:
            row_tuple = tuple(str(c) if c is not None else "" for c in row)
            if row_tuple not in seen:
                seen.add(row_tuple)
                unique_data.append(row)
        return unique_data

    def process_queue(self):
        try:
            msg_type, result = self.msg_queue.get_nowait()
            self.progress_bar.stop()
            self.is_processing = False
            self.btn_run.config(state=tk.NORMAL)
            
            if msg_type == "SUCCESS":
                self.extracted_data = result
                count = len(result)
                self.status_var.set(f"完成！共匹配到 {count} 条符合条件的数据。")
                self.update_treeview_preview(result)
                if count > 0:
                    self.save_btn.config(state=tk.NORMAL)
                else:
                    messagebox.showinfo("结果", "未能找到任何匹配的数据行。")
            elif msg_type == "ERROR":
                self.status_var.set(f"提取失败: {result}")
                messagebox.showerror("提取失败", f"处理过程中发生异常:\n{result}")

        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_queue)

    def update_treeview_preview(self, data):
        self.tree.delete(*self.tree.get_children())
        if not data:
            return

        max_cols_in_data = max(len(row) for row in data[:100])
        display_col_count = min(max_cols_in_data, 30) 

        cols = [f"col_{i}" for i in range(display_col_count)]
        self.tree["columns"] = cols

        for i in range(display_col_count):
            col_name = "来源Sheet" if (i == 0 and self.include_source_sheet_var.get()) else f"列 {i if not self.include_source_sheet_var.get() else i}"
            self.tree.heading(f"col_{i}", text=col_name)
            self.tree.column(f"col_{i}", width=110, anchor=tk.W)

        for row in data[:100]:
            display_row = [str(c) if c is not None else "" for c in row[:display_col_count]]
            self.tree.insert("", tk.END, values=display_row)

    def save_to_excel(self):
        if not self.extracted_data:
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if save_path:
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "提取结果"

                for row in self.extracted_data:
                    ws.append([str(cell) if cell is not None else "" for cell in row])

                wb.save(save_path)
                self.status_var.set(f"数据已成功保存至: {save_path}")
                messagebox.showinfo("导出成功", f"文件保存成功！\n共保存 {len(self.extracted_data)} 行数据。")
            except PermissionError:
                messagebox.showerror("保存失败", f"无法写入文件！\n文件可能已在 Excel 中打开，请先关闭该文件后再试：\n{save_path}")
            except Exception as e:
                messagebox.showerror("保存失败", f"导出过程中发生未知错误:\n{str(e)}")

    def on_closing(self):
        if self.is_processing:
            if messagebox.askokcancel("退出确认", "程序正在提取数据，确定要强行退出吗？"):
                self.root.destroy()
        else:
            self.root.destroy()

if __name__ == "__main__":
    root = ttk.Window(themename="superhero")
    app = ExcelExtractorApp(root)
    root.mainloop()