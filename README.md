# office-helper
[![CI](https://github.com/wzx11223344/office-helper/actions/workflows/ci.yml/badge.svg)](https://github.com/wzx11223344/office-helper/actions/workflows/ci.yml)


办公效率工具箱 - 10个高级算法驱动的办公自动化工具集

## 功能概览

| # | 函数名 | 算法原理 | 复杂度 |
|---|--------|----------|--------|
| 1 | `smart_file_classifier` | TF-IDF向量空间模型 + 余弦相似度 | O(n^2 * m)，n=文件数，m=特征词数 |
| 2 | `pdf_structure_analyzer` | PDF二进制解析 + zlib解压 + 正则提取 | O(n)，n=页数 |
| 3 | `document_similarity_compare` | Jaccard相似度 + TF-IDF余弦相似度 + Levenshtein编辑距离 | O(n*m + |s1|*|s2|) |
| 4 | `batch_rename_with_pattern` | 正则匹配 + 序号填充 + 日期格式化 + 变量替换 | O(n)，n=文件数 |
| 5 | `excel_data_merger` | XLSX解析 + 多表连接(LEFT/RIGHT/OUTER/INNER) | O(n*m)，n,m=表行数 |
| 6 | `meeting_minutes_parser` | 正则引擎 + 规则引擎 + 自然语言提取 | O(n)，n=文本长度 |
| 7 | `schedule_conflict_detector` | AVL平衡区间树 + 区间重叠检测 | O(n log n + k)，k=冲突数 |
| 8 | `text_diff_analyzer` | Myers差异算法（最长公共子序列） | O(n*m)，n,m=行数 |
| 9 | `file_duplicate_finder` | MD5哈希 + 文件大小双重过滤 + 文件名相似度 | O(n)哈希 + O(n^2)模糊 |
| 10 | `document_summary_generator` | TextRank图排序算法（PageRank迭代） | O(n^2 * k + k*T)，n=句子数，k=迭代次数 |

## 算法详解

### 1. TF-IDF智能文件分类器 (`smart_file_classifier`)
- **原理**: 将每个文件内容转化为TF-IDF向量，通过余弦相似度计算文件间相似度，使用K-Means++聚类自动分组
- **TF-IDF公式**: `tfidf(t,d) = tf(t,d) * log(N / df(t))`
- **余弦相似度**: `sim(a,b) = (a·b) / (|a| * |b|)`
- **复杂度**: 构建TF-IDF矩阵 O(n*m)，相似度计算 O(n^2*m)，聚类 O(n^2*k)

### 2. PDF文档结构分析器 (`pdf_structure_analyzer`)
- **原理**: 直接解析PDF二进制格式，使用zlib解压FlateDecode流，通过正则提取页面层次结构
- **解析流程**: 读取PDF → 提取xref表 → 解析Page对象 → 解压内容流 → 正则匹配文本/图片/表格区域
- **复杂度**: O(n)，n=页数

### 3. 三重算法文档相似度对比 (`document_similarity_compare`)
- **Jaccard相似度**: `J(A,B) = |A∩B| / |A∪B|`，基于词集合
- **TF-IDF余弦相似度**: 基于加权向量的夹角余弦
- **Levenshtein编辑距离**: 动态规划 `dp[i][j] = min(dp[i-1][j]+1, dp[i][j-1]+1, dp[i-1][j-1]+cost)`
- **复杂度**: 编辑距离 O(|s1|*|s2|)，使用滚动数组优化空间至 O(min(|s1|,|s2|))

### 4. 智能批量重命名 (`batch_rename_with_pattern`)
- **正则匹配**: 使用 `re` 模块进行模式匹配
- **序号填充**: 支持零填充 `{:0{n}d}`
- **日期格式化**: 支持 `%Y-%m-%d` 等strftime格式
- **变量替换**: 支持 `{name}`, `{ext}`, `{index}`, `{date}` 等变量

### 5. 多Excel文件智能合并 (`excel_data_merger`)
- **XLSX解析**: ZIP解包 → 解析XML → 提取单元格数据
- **连接算法**: 
  - LEFT JOIN: 保留左表所有行
  - RIGHT JOIN: 保留右表所有行
  - OUTER JOIN: 保留两表所有行
  - INNER JOIN: 仅保留匹配行
- **类型检测**: 自动推断数值/文本/日期类型
- **复杂度**: 哈希连接 O(n+m)

### 6. 会议纪要结构化解析 (`meeting_minutes_parser`)
- **正则引擎**: 多模式正则匹配议题、决议、待办、责任人、截止日期
- **规则引擎**: 基于关键词的上下文推断
- **自然语言提取**: "由XX负责" → 责任人提取，"X月X日前完成" → 截止日期提取

### 7. 排班冲突检测器 (`schedule_conflict_detector`)
- **区间树算法**: AVL平衡二叉搜索树，每个节点存储区间[low, high]
- **重叠检测**: 对每个区间查询区间树，找到所有重叠区间
- **AVL平衡**: 插入后通过LL/RR/LR/RL旋转保持平衡
- **复杂度**: 构建O(n log n)，查询O(k + log n)

### 8. Myers文本差异算法 (`text_diff_analyzer`)
- **原理**: 求解两个序列的最短编辑脚本(SES)，等价于最长公共子序列(LCS)
- **算法**: 动态规划 + 回溯，输出增/删/改的行级diff
- **复杂度**: O(n*m) 时间，O(n*m) 空间

### 9. 文件查重器 (`file_duplicate_finder`)
- **双重过滤**: 文件大小 + MD5哈希，快速排除不同文件
- **模糊匹配**: 基于文件名Levenshtein距离的相似度计算
- **流程**: 大小过滤 → MD5精确匹配 → 文件名相似度模糊匹配
- **复杂度**: 精确匹配O(n)，模糊匹配O(n^2)

### 10. TextRank文档摘要生成器 (`document_summary_generator`)
- **句子分割**: 基于标点符号的句子边界检测
- **TF-IDF相似度**: 计算句子间相似度构建邻接矩阵
- **PageRank迭代**: `S(Vi) = (1-d) + d * Σ(S(Vj)/C(Vj))`，d=0.85阻尼系数
- **Top-K提取**: 按PageRank得分排序，选Top-K句子作为摘要
- **复杂度**: 相似度矩阵O(n^2)，PageRank迭代O(n^2*T)

## 安装

无需安装额外依赖，仅使用Python标准库（hashlib, re, os, json, collections, math, difflib, zipfile, zlib等）。

## 使用示例

```python
from main import smart_file_classifier, document_similarity_compare, text_diff_analyzer

# TF-IDF智能文件分类
result = smart_file_classifier("/path/to/files")
print(result["clusters"])  # 自动分类结果

# 三重算法文档相似度对比
sim = document_similarity_compare("doc1.txt", "doc2.txt")
print(f"Jaccard: {sim['jaccard']:.3f}, Cosine: {sim['cosine']:.3f}, Edit: {sim['edit_distance']}")

# Myers文本差异
diff = text_diff_analyzer("text1.txt", "text2.txt")
for line in diff["diff"]:
    print(f"{line['type']}: {line['content']}")
```

## License

MIT
