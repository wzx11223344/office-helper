---
name: office-helper-zx
displayName: 办公效率工具箱
version: 1.0.1
summary: 10个办公工具：批量重命名/PDF合并拆分/Excel-CSV互转/会议纪要/排班表/文件整理/文本替换/密码生成
tags: [office, productivity, pdf, excel, automation]
license: MIT
---

# 办公效率工具箱 (office-helper)

## 简介

office-helper 是一套包含10个常用办公自动化工具的 Python 技能包，旨在提升日常办公效率。涵盖文件管理、PDF处理、表格转换、文档生成、文件整理和文本处理等场景。

## 功能列表

| # | 函数名 | 功能描述 |
|---|--------|----------|
| 1 | `batch_rename_files` | 批量重命名文件（正则匹配） |
| 2 | `merge_pdf_files` | 合并多个PDF文件为一个 |
| 3 | `split_pdf_file` | 按页码范围拆分PDF文件 |
| 4 | `excel_to_csv` | Excel文件转换为CSV（支持多工作表） |
| 5 | `csv_to_excel` | CSV文件转换为Excel |
| 6 | `generate_meeting_minutes` | 生成Markdown格式会议纪要 |
| 7 | `create_work_schedule` | 生成一周排班表 |
| 8 | `file_organizer` | 按扩展名自动分类整理文件 |
| 9 | `text_replacer` | 批量文本替换（正则支持） |
| 10 | `password_generator` | 随机密码生成器（强度评估） |

## 安装

```bash
pip install openpyxl PyPDF2 python-docx
```

## 使用示例

```python
from main import password_generator, generate_meeting_minutes

# 生成密码
result = password_generator(length=16)
print(result["password"])

# 生成会议纪要
minutes = generate_meeting_minutes(
    title="项目周会",
    attendees=["张三", "李四"],
    agenda=["项目进度汇报", "下周计划"],
    decisions["项目按计划推进"],
    action_items=[{"task": "完成模块A", "owner": "张三", "deadline": "2025-01-15"}]
)
print(minutes)
```

## 依赖

- `openpyxl`: Excel文件读写
- `PyPDF2`: PDF文件操作
- `python-docx`: Word文档操作

## License

MIT
