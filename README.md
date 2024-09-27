# ExcelDataMatcher

ExcelDataMatcher 是一个用于从 Excel 文件中提取特定行数据的 Python 应用程序。它使用 Tkinter 和 ttkbootstrap 构建用户界面，并使用 pandas 处理 Excel 文件。

## 功能特点

- 选择 Excel 文件并提取特定列名的行数据。
- 支持选择单个工作表或搜索所有工作表。
- 支持指定表头行或不需要表头。
- 提取的数据可以保存为新的 Excel 文件。

## 使用方法

克隆仓库
```bash
git clone https://github.com/yourusername/ExcelDataMatcher.git
cd ExcelDataMatcher

创建虚拟环境（可选）：
python -m venv venv
source venv/bin/activate  # 在 Windows 上使用 `venv\Scripts\activate`

安装依赖
pip install -r requirements.txt

运行
python ExcelDataMatcher.py
