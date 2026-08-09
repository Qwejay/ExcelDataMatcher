# ExcelDataMatcher

ExcelDataMatcher 是一个用于从 Excel 文件中提取特定行数据的应用程序。它使用 Tkinter 和 ttkbootstrap 构建用户界面，并使用 pandas 处理 Excel 文件。

## 功能特点

- **文件兼容**：支持 `.xlsx` 与 `.xls`，可检索单个 Sheet 或跨全表检索，支持指定表头行。
- **高级匹配**：支持`模糊包含`、`精确匹配`与`正则表达式`；支持 `AND/OR` 多条件组合及`指定列`查找。
- **数据清洗**：支持清理单元格内换行符、剔除整行空数据及结果一键去重。
- **流畅体验**：采用多线程后台架构，大文件提取不卡死，自带结果 Treeview 实时预览。
- **结果导出**：提取数据可一键导出为全新 `.xlsx` 文件，并可选择保留来源 Sheet 追踪信息。

---

### 📝 CHANGELOG

## v1.2 - 2026-8-9

### 新增功能 (Features)
- **多模式匹配**：支持 `包含(模糊)`、`精确匹配`、`正则表达式` 3种匹配模式。
- **逻辑运算符**：支持 `满足任一(OR)` 与 `满足所有(AND)` 逻辑控制。
- **指定列查找**：可限制仅在特定列（如 `A, B` 或 `1, 3`）中查找。
- **数据实时预览**：界面新增 Treeview 表格，提取后直接预览前 100 条数据。
- **溯源标记**：新增“保留来源 Sheet 名称”选项，方便多表合并后追溯数据。

### 性能与体验 (Improvements)
- **多线程架构**：采用后台异步提取，解决大文件处理时界面“未响应”卡死问题。
- **UI 布局重构**：全面改用 Grid 网格布局，彻底解决表头行挤压、遮挡显示不全的问题。

### 鲁棒性与修复 (Fixes)
- **正则安全预检**：非法正则表达式输入时抛出友好的错误提示，防止程序崩溃。
- **文件占用捕获**：导出时如遇到 Excel 文件被占用，自动提示用户关闭后重试。
- **防越界与卡顿**：限制预览表格最大列数，修复部分数据列数不足导致的索引越界。

---

## v1.1

### 功能
- 优化程序大小，使用xlrd和openpyxl替代pandas
- 列名输入框增加右键菜单，支持粘贴、复制和清除操作
- 增加状态栏，实时显示操作状态
- 优化界面布局

## 使用方法

```bash
克隆仓库
git clone https://github.com/yourusername/ExcelDataMatcher.git
cd ExcelDataMatcher

创建虚拟环境（可选）：
python -m venv venv
source venv/bin/activate

安装依赖
pip install -r requirements.txt

运行
python ExcelDataMatcher.py

打包代码
pyinstaller --onefile --noconsole --icon=icon.ico --name=ExcelDataMatcher --add-data "icon.ico;." --hidden-import=tkinter --exclude-module=pytest --exclude-module=unittest --clean --strip ExcelDataMatcher.py