---
name: office-helper-zx
displayName: 办公效率工具箱
version: 2.0.0
summary: 10个高级算法办公工具：TF-IDF文件分类/PDF结构分析/三重文档相似度/智能批量重命名/Excel多表合并/会议纪要解析/区间树排班冲突/Myers文本差异/MD5文件查重/TextRank文档摘要
tags: [office, productivity, tfidf, textrank, myers-diff, interval-tree, automation]
license: MIT
---

# 办公效率工具箱 (office-helper)

## 简介

office-helper 是一套包含10个高级算法驱动的办公自动化工具的 Python 技能包。所有功能仅使用Python标准库实现，无需安装任何外部依赖。涵盖文件分类、PDF分析、文档对比、批量重命名、Excel合并、会议解析、排班检测、文本差异、文件查重、文档摘要等场景。

## 功能列表

| # | 函数名 | 算法原理 | 复杂度 |
|---|--------|----------|--------|
| 1 | `smart_file_classifier` | TF-IDF向量空间模型 + 余弦相似度 + K-Means++聚类 | O(n^2 * m) |
| 2 | `pdf_structure_analyzer` | PDF二进制解析 + zlib解压 + 正则提取 | O(n) |
| 3 | `document_similarity_compare` | Jaccard + TF-IDF余弦 + Levenshtein编辑距离 | O(n*m + \|s1\|*\|s2\|) |
| 4 | `batch_rename_with_pattern` | 正则匹配 + 序号填充 + 日期格式化 + 变量替换 | O(n) |
| 5 | `excel_data_merger` | XLSX解析 + LEFT/RIGHT/OUTER/INNER JOIN | O(n+m) |
| 6 | `meeting_minutes_parser` | 正则引擎 + 规则引擎 + NLP提取 | O(n) |
| 7 | `schedule_conflict_detector` | AVL平衡区间树 + 区间重叠检测 | O(n log n + k) |
| 8 | `text_diff_analyzer` | Myers差异算法（LCS动态规划） | O(n*m) |
| 9 | `file_duplicate_finder` | MD5哈希 + 文件大小过滤 + 文件名相似度 | O(n) + O(n^2) |
| 10 | `document_summary_generator` | TextRank图排序（PageRank迭代） | O(n^2 * T) |

## 安装

无需安装额外依赖，仅使用Python标准库（hashlib, re, os, json, collections, math, difflib, zipfile, zlib等）。

## 使用示例

```python
from main import smart_file_classifier, document_similarity_compare, text_diff_analyzer

# TF-IDF智能文件分类
result = smart_file_classifier("/path/to/files")
print(result["clusters"])

# 三重算法文档相似度对比
sim = document_similarity_compare("doc1.txt", "doc2.txt")
print(f"Jaccard: {sim['jaccard']:.3f}, Cosine: {sim['cosine']:.3f}")

# Myers文本差异
diff = text_diff_analyzer("text1.txt", "text2.txt")
for line in diff["diff"]:
    print(f"{line['type']}: {line['content']}")
```

## 依赖

无外部依赖（仅使用Python标准库）

## License

MIT
