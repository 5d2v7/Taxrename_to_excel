# Taxrename_to_excel

## 项目介绍
Taxrename_to_excel是一个自动化脚本，用于从中国电子税务局导出的纳税申报表PDF文件中提取表格数据并将其保存到Excel文件中，并根据提取的信息重命名文件。

## 功能特点
- **自动提取**: 从PDF文档中自动提取表格数据（最多提取到第二页）。
- **数据清洗**: 清除数据中的换行符，确保数据整洁。
- **智能重命名**: 根据提取的公司名称、税种和受理日期重命名Excel和PDF文件。
- **日志记录**: 记录处理过程中的所有重要信息和警告。

## 使用前提
在使用本脚本之前，请确保已安装以下Python库：
- `pdfplumber`
- `openpyxl`
- `re`
- `os`
- `glob`
- `configparser`
- `logging`
- `shutil`

## 配置文件
脚本使用`config.ini`文件来读取必要的路径和参数。请确保在脚本同一目录下提供此配置文件，并正确设置以下参数：
- `PDFDirectoryPath`: PDF文件存放的目录路径。
- `OutputDirectoryPath`: 输出Excel文件和重命名后的PDF文件存放的目录路径。
- `TaxTypes`: 税种列表，用于提取和重命名文件。

## 使用指南
1. 准备`config.ini`配置文件，并设置好相关路径和参数。
2. 在脚本所在目录下运行脚本：
   ```shell
   python Taxrename_to_excel.py
