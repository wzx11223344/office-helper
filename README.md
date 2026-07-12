# office-helper

办公效率工具箱 - 10个常用办公自动化工具集

## 功能概览

- **批量重命名文件** - 使用正则表达式批量重命名文件
- **PDF合并** - 将多个PDF文件合并为一个
- **PDF拆分** - 按页码范围拆分PDF
- **Excel转CSV** - Excel文件转换为CSV（支持多工作表）
- **CSV转Excel** - CSV文件转换为Excel
- **会议纪要生成** - 生成Markdown格式会议纪要
- **排班表生成** - 自动生成一周排班表
- **文件自动分类** - 按扩展名整理文件到分类目录
- **批量文本替换** - 正则支持的批量文本替换
- **密码生成器** - 随机密码生成（含强度评估）

## 安装

```bash
pip install -r requirements.txt
```

## 快速开始

```python
from main import password_generator

result = password_generator(length=16, use_symbols=True)
print(f"密码: {result['password']}")
print(f"强度: {result['strength']}")
```

## License

MIT
