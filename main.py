"""
office-helper - 办公效率工具箱

本模块实现了10个高级办公自动化算法工具，全部使用Python标准库实现：
1. smart_file_classifier      - 基于TF-IDF+余弦相似度的智能文件分类器
2. pdf_structure_analyzer    - PDF文档结构分析器（解析内容流构建文档树）
3. document_similarity_compare - 文档相似度对比（Jaccard+TF-IDF余弦+编辑距离）
4. batch_rename_with_pattern  - 智能批量重命名（正则+序号填充+日期格式化+变量替换）
5. excel_data_merger          - 多Excel智能合并（左/右/外/内连接+类型推断）
6. meeting_minutes_parser     - 会议纪要结构化解析（正则+规则引擎）
7. schedule_conflict_detector - 排班冲突检测器（区间树算法）
8. text_diff_analyzer         - 文本差异分析（Myers差异算法）
9. file_duplicate_finder      - 文件查重器（MD5哈希+模糊匹配）
10. document_summary_generator - 文档摘要生成器（TextRank图排序算法）

所有算法均从零实现，不依赖numpy/pandas等第三方库。
"""

import os
import re
import math
import json
import hashlib
import zipfile
import zlib
import struct
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict, deque
from datetime import datetime, timedelta
from typing import List, Dict, Any, Tuple, Optional


# ============================================================================
# 辅助函数：文本分词器
# ============================================================================

def _tokenize(text: str) -> List[str]:
    """
    混合分词器：英文按单词切分，中文按双字组切分。
    算法：先用正则提取英文单词和中文片段，再对中文做2-gram。
    """
    tokens = []
    # 提取英文单词
    en_words = re.findall(r'[a-zA-Z]{2,}', text.lower())
    tokens.extend(en_words)
    # 提取中文字符序列，做2-gram
    cn_segments = re.findall(r'[\u4e00-\u9fff]+', text)
    for seg in cn_segments:
        if len(seg) >= 2:
            for i in range(len(seg) - 1):
                tokens.append(seg[i:i + 2])
        elif len(seg) == 1:
            tokens.append(seg)
    return tokens


def _read_file_content(path: str, max_bytes: int = 500000) -> str:
    """安全读取文件文本内容，限制最大字节数。"""
    try:
        with open(path, 'rb') as f:
            raw = f.read(max_bytes)
        for enc in ('utf-8', 'gbk', 'latin-1'):
            try:
                return raw.decode(enc)
            except (UnicodeDecodeError, ValueError):
                continue
        return raw.decode('latin-1', errors='replace')
    except (IOError, OSError):
        return ""


# ============================================================================
# 1. smart_file_classifier - 基于TF-IDF的智能文件分类器
# ============================================================================

def _compute_tfidf_vectors(documents: List[List[str]]) -> Tuple[Dict[int, Dict[str, float]], set]:
    """
    计算TF-IDF向量。
    算法步骤：
    1. 构建词表（所有不重复词项）
    2. 计算每个文档的词频TF
    3. 计算逆文档频率IDF = log(N / df)
    4. TF-IDF = TF * IDF
    """
    N = len(documents)
    if N == 0:
        return {}, set()

    # 构建词表和文档频率
    vocabulary = set()
    df = Counter()  # 文档频率
    doc_tokens = []

    for tokens in documents:
        unique_tokens = set(tokens)
        vocabulary.update(unique_tokens)
        df.update(unique_tokens)
        doc_tokens.append(tokens)

    # 计算IDF
    idf = {term: math.log(N / df[term]) for term in vocabulary}

    # 计算每个文档的TF-IDF向量
    tfidf_vectors = {}
    for i, tokens in enumerate(doc_tokens):
        tf = Counter(tokens)
        total = len(tokens) if tokens else 1
        tfidf_vectors[i] = {term: (count / total) * idf[term] for term, count in tf.items()}

    return tfidf_vectors, vocabulary


def _cosine_similarity(vec1: Dict[str, float], vec2: Dict[str, float]) -> float:
    """
    计算两个稀疏向量的余弦相似度。
    cos(a, b) = (a·b) / (||a|| * ||b||)
    """
    # 只计算共同键的点积
    common_keys = set(vec1.keys()) & set(vec2.keys())
    dot_product = sum(vec1[k] * vec2[k] for k in common_keys)

    norm1 = math.sqrt(sum(v * v for v in vec1.values()))
    norm2 = math.sqrt(sum(v * v for v in vec2.values()))

    if norm1 == 0 or norm2 == 0:
        return 0.0
    return dot_product / (norm1 * norm2)


def _kmeans_cluster(vectors: Dict[int, Dict[str, float]], k: int, max_iter: int = 50) -> Dict[int, List[int]]:
    """
    从零实现K-Means聚类算法。
    算法步骤：
    1. 随机选择k个中心点（采用K-Means++初始化）
    2. 将每个点分配到最近的中心
    3. 重新计算中心点
    4. 重复直到收敛或达到最大迭代次数
    """
    import random
    indices = list(vectors.keys())
    if len(indices) <= k:
        return {i: [idx] for i, idx in enumerate(indices)}

    # K-Means++ 初始化
    first_center = random.choice(indices)
    centers = [vectors[first_center]]
    for _ in range(1, k):
        # 选择距离现有中心最远的点
        max_dist = -1
        best_idx = indices[0]
        for idx in indices:
            min_dist = min(_cosine_similarity(vectors[idx], c) for c in centers)
            # 用1-余弦相似度作为距离
            dist = 1 - min_dist
            if dist > max_dist:
                max_dist = dist
                best_idx = idx
        centers.append(vectors[best_idx])

    # 迭代
    clusters = defaultdict(list)
    for iteration in range(max_iter):
        new_clusters = defaultdict(list)
        for idx in indices:
            best_k = 0
            best_sim = -1
            for ci, center in enumerate(centers):
                sim = _cosine_similarity(vectors[idx], center)
                if sim > best_sim:
                    best_sim = sim
                    best_k = ci
            new_clusters[best_k].append(idx)

        # 检查是否收敛
        if dict(new_clusters) == dict(clusters):
            break
        clusters = new_clusters

        # 更新中心点（取均值）
        for ci, members in clusters.items():
            all_terms = set()
            for m in members:
                all_terms.update(vectors[m].keys())
            new_center = {}
            for term in all_terms:
                vals = [vectors[m].get(term, 0) for m in members]
                new_center[term] = sum(vals) / len(vals)
            centers[ci] = new_center

    return dict(clusters)


def smart_file_classifier(files_path: str, num_clusters: int = 3) -> Dict[str, Any]:
    """
    基于文件内容TF-IDF的智能文件分类器。

    算法流程：
    1. 读取目录下所有文本文件内容
    2. 使用双字组分词器提取特征词
    3. 构建TF-IDF向量空间模型
    4. 使用K-Means++聚类算法进行自动分类
    5. 计算每个簇的Top关键词

    参数:
        files_path: 文件目录路径
        num_clusters: 聚类数量

    返回:
        包含分类结果、每类关键词、相似度矩阵的字典
    """
    # 步骤1：收集文件和内容
    files = []
    documents = []
    for fname in sorted(os.listdir(files_path)):
        fpath = os.path.join(files_path, fname)
        if not os.path.isfile(fpath):
            continue
        content = _read_file_content(fpath)
        if not content.strip():
            continue
        tokens = _tokenize(content)
        if len(tokens) < 2:
            continue
        files.append(fname)
        documents.append(tokens)

    if len(files) == 0:
        return {"error": "没有找到可分类的文件", "clusters": {}}

    # 步骤2：计算TF-IDF向量
    tfidf_vectors, vocabulary = _compute_tfidf_vectors(documents)

    # 步骤3：K-Means聚类
    k = min(num_clusters, len(files))
    clusters = _kmeans_cluster(tfidf_vectors, k)

    # 步骤4：提取每类Top关键词
    result = {"clusters": {}, "file_count": len(files), "vocabulary_size": len(vocabulary)}
    for cluster_id, member_indices in clusters.items():
        # 汇总该类所有文件的TF-IDF
        cluster_tfidf = defaultdict(float)
        for idx in member_indices:
            for term, val in tfidf_vectors[idx].items():
                cluster_tfidf[term] += val

        # 按权重排序取Top-10
        top_keywords = sorted(cluster_tfidf.items(), key=lambda x: -x[1])[:10]
        result["clusters"][f"cluster_{cluster_id}"] = {
            "files": [files[i] for i in member_indices],
            "keywords": [kw for kw, _ in top_keywords],
            "keyword_weights": {kw: round(w, 4) for kw, w in top_keywords}
        }

    # 步骤5：计算相似度矩阵
    sim_matrix = []
    for i in range(len(files)):
        row = []
        for j in range(len(files)):
            row.append(round(_cosine_similarity(tfidf_vectors[i], tfidf_vectors[j]), 4))
        sim_matrix.append(row)
    result["similarity_matrix"] = sim_matrix

    return result


# ============================================================================
# 2. pdf_structure_analyzer - PDF文档结构分析器
# ============================================================================

def _extract_pdf_objects(data: bytes) -> Dict[int, bytes]:
    """
    解析PDF二进制格式，提取所有间接对象。
    PDF格式：每个对象以 'N 0 obj' 开头，以 'endobj' 结尾。
    """
    objects = {}
    # 使用正则查找对象边界
    obj_pattern = re.compile(rb'(\d+)\s+(\d+)\s+obj\b')
    for match in obj_pattern.finditer(data):
        obj_num = int(match.group(1))
        start = match.end()
        end_pos = data.find(b'endobj', start)
        if end_pos == -1:
            end_pos = len(data)
        obj_data = data[start:end_pos]
        objects[obj_num] = obj_data
    return objects


def _extract_pdf_text_streams(data: bytes) -> List[str]:
    """
    提取PDF内容流中的文本。
    算法：
    1. 查找所有 stream...endstream 块
    2. 尝试用zlib解压FlateDecode编码的流
    3. 从解压后的内容中提取文本操作符 (Tj, TJ)
    """
    texts = []
    stream_pattern = re.compile(rb'stream\r?\n(.*?)\r?\nendstream', re.DOTALL)

    for match in stream_pattern.finditer(data):
        stream_data = match.group(1)
        text_content = None

        # 尝试zlib解压（FlateDecode）
        try:
            decompressed = zlib.decompress(stream_data)
            text_content = decompressed
        except (zlib.error, ValueError):
            # 可能是未压缩的流
            try:
                text_content = stream_data.decode('latin-1').encode('latin-1')
            except Exception:
                continue

        if text_content:
            # 提取 Tj 操作符中的文本：(...) Tj
            tj_pattern = rb'\(([^)]*)\)\s*Tj'
            for tj_match in re.finditer(tj_pattern, text_content):
                raw = tj_match.group(1)
                try:
                    decoded = raw.decode('latin-1')
                    if decoded.strip():
                        texts.append(decoded)
                except Exception:
                    pass

            # 提取 TJ 数组操作符：[(...) ...] TJ
            tj_array_pattern = rb'\[(.*?)\]\s*TJ'
            for tj_match in re.finditer(tj_array_pattern, text_content, re.DOTALL):
                array_content = tj_match.group(1)
                parts = re.findall(rb'\(([^)]*)\)', array_content)
                line = b''.join(parts)
                try:
                    decoded = line.decode('latin-1')
                    if decoded.strip():
                        texts.append(decoded)
                except Exception:
                    pass

    return texts


def _classify_text_block(text: str) -> str:
    """
    基于规则分类文本块类型。
    - 标题：短文本、无标点结尾、含数字编号
    - 段落：长文本
    - 列表项：以数字/字母+点开头
    - 表格行：含多个制表符或竖线
    """
    text = text.strip()
    if not text:
        return "empty"

    # 标题特征：短（<50字符）、行尾无句号
    if len(text) < 50 and not text.endswith(('。', '.', '!', '？', '？')):
        if re.match(r'^第?\d+[章章节条\.、\s]', text) or re.match(r'^[一二三四五六七八九十]+[、\.]', text):
            return "heading"
        if len(text) < 30:
            return "heading"

    # 列表项
    if re.match(r'^[\d①②③④⑤⑥⑦⑧⑨⑩a-zA-Z][\.\)、]', text):
        return "list_item"

    # 表格行（多列分隔符）
    if text.count('\t') >= 2 or text.count('|') >= 2:
        return "table_row"

    # 普通段落
    if len(text) > 50:
        return "paragraph"

    return "text"


def pdf_structure_analyzer(pdf_path: str) -> Dict[str, Any]:
    """
    PDF文档结构分析器。

    算法流程：
    1. 读取PDF二进制数据
    2. 解析PDF交叉引用表获取对象布局
    3. 提取所有内容流并用zlib解压
    4. 解析文本操作符(Tj/TJ)还原文本
    5. 使用规则引擎分类文本块（标题/段落/表格/列表）
    6. 构建层次化文档树

    参数:
        pdf_path: PDF文件路径

    返回:
        包含页数、文档树、统计信息的字典
    """
    if not os.path.exists(pdf_path):
        return {"error": "文件不存在"}

    with open(pdf_path, 'rb') as f:
        data = f.read()

    # 步骤1：提取PDF对象
    objects = _extract_pdf_objects(data)

    # 步骤2：提取页面数量（从Pages对象的Count字段）
    page_count = 0
    for obj_data in objects.values():
        count_match = re.search(rb'/Count\s+(\d+)', obj_data)
        if count_match:
            val = int(count_match.group(1))
            if val > page_count:
                page_count = val

    # 步骤3：提取文本流
    text_blocks = _extract_pdf_text_streams(data)

    # 步骤4：分类文本块并构建文档树
    doc_tree = {"type": "root", "children": []}
    current_section = None
    stats = Counter()

    for block_text in text_blocks:
        block_text = block_text.strip()
        if not block_text:
            continue

        block_type = _classify_text_block(block_text)
        stats[block_type] += 1
        node = {"type": block_type, "text": block_text, "children": []}

        if block_type == "heading":
            # 新章节
            current_section = node
            doc_tree["children"].append(node)
        else:
            if current_section is None:
                doc_tree["children"].append(node)
            else:
                current_section["children"].append(node)

    return {
        "file": os.path.basename(pdf_path),
        "file_size": len(data),
        "page_count": page_count,
        "object_count": len(objects),
        "text_block_count": len(text_blocks),
        "document_tree": doc_tree,
        "structure_stats": dict(stats),
        "headings": [c["text"] for c in doc_tree["children"] if c["type"] == "heading"]
    }


# ============================================================================
# 3. document_similarity_compare - 文档相似度对比
# ============================================================================

def _jaccard_similarity(set1: set, set2: set) -> float:
    """Jaccard相似度 = |A∩B| / |A∪B|"""
    if not set1 and not set2:
        return 1.0
    intersection = set1 & set2
    union = set1 | set2
    return len(intersection) / len(union) if union else 0.0


def _levenshtein_distance(s1: str, s2: str) -> int:
    """
    Levenshtein编辑距离（动态规划实现）。
    dp[i][j] = s1[:i]和s2[:j]之间的最小编辑距离
    转移方程：
      if s1[i-1] == s2[j-1]: dp[i][j] = dp[i-1][j-1]
      else: dp[i][j] = 1 + min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1])
    空间优化：使用两行滚动数组，O(min(m,n))空间
    """
    if len(s1) < len(s2):
        s1, s2 = s2, s1

    m, n = len(s1), len(s2)
    prev = list(range(n + 1))
    curr = [0] * (n + 1)

    for i in range(1, m + 1):
        curr[0] = i
        for j in range(1, n + 1):
            if s1[i - 1] == s2[j - 1]:
                curr[j] = prev[j - 1]
            else:
                curr[j] = 1 + min(prev[j], curr[j - 1], prev[j - 1])
        prev, curr = curr, prev

    return prev[n]


def document_similarity_compare(doc1_path: str, doc2_path: str) -> Dict[str, Any]:
    """
    文档相似度对比，使用三重算法。

    算法：
    1. Jaccard相似度：基于词集合的交集/并集
    2. TF-IDF余弦相似度：基于词频-逆文档频率的向量空间模型
    3. Levenshtein编辑距离：字符级编辑距离，归一化到[0,1]

    参数:
        doc1_path: 文档1路径
        doc2_path: 文档2路径

    返回:
        包含三种相似度分数和综合评估的字典
    """
    text1 = _read_file_content(doc1_path)
    text2 = _read_file_content(doc2_path)

    if not text1 or not text2:
        return {"error": "无法读取文件内容"}

    # 算法1：Jaccard相似度（基于词集合）
    tokens1 = set(_tokenize(text1))
    tokens2 = set(_tokenize(text2))
    jaccard_score = _jaccard_similarity(tokens1, tokens2)

    # 算法2：TF-IDF余弦相似度
    docs = [_tokenize(text1), _tokenize(text2)]
    tfidf_vectors, _ = _compute_tfidf_vectors(docs)
    cosine_score = _cosine_similarity(tfidf_vectors.get(0, {}), tfidf_vectors.get(1, {}))

    # 算法3：编辑距离（归一化）
    max_len = max(len(text1), len(text2))
    if max_len > 0:
        edit_dist = _levenshtein_distance(text1[:10000], text2[:10000])  # 限制长度防止超时
        edit_similarity = 1.0 - (edit_dist / max(len(text1[:10000]), len(text2[:10000]), 1))
    else:
        edit_similarity = 1.0

    # 综合评分（加权平均）
    composite = (jaccard_score * 0.3 + cosine_score * 0.5 + edit_similarity * 0.2)

    # 共同关键词
    common_terms = tokens1 & tokens2
    tf1 = Counter(_tokenize(text1))
    tf2 = Counter(_tokenize(text2))
    common_with_freq = sorted(
        [(t, min(tf1.get(t, 0), tf2.get(t, 0))) for t in common_terms],
        key=lambda x: -x[1]
    )[:20]

    return {
        "doc1": os.path.basename(doc1_path),
        "doc2": os.path.basename(doc2_path),
        "jaccard_similarity": round(jaccard_score, 4),
        "tfidf_cosine_similarity": round(cosine_score, 4),
        "edit_distance_similarity": round(edit_similarity, 4),
        "composite_score": round(composite, 4),
        "similarity_level": _similarity_level(composite),
        "common_terms": [{"term": t, "frequency": f} for t, f in common_with_freq],
        "doc1_unique_terms": len(tokens1 - tokens2),
        "doc2_unique_terms": len(tokens2 - tokens1),
        "common_terms_count": len(common_terms)
    }


def _similarity_level(score: float) -> str:
    if score >= 0.8:
        return "高度相似"
    elif score >= 0.6:
        return "中等相似"
    elif score >= 0.3:
        return "低度相似"
    else:
        return "基本不同"


# ============================================================================
# 4. batch_rename_with_pattern - 智能批量重命名
# ============================================================================

def batch_rename_with_pattern(dir_path: str, pattern: str,
                               rename_rule: str, dry_run: bool = True) -> Dict[str, Any]:
    """
    智能批量重命名工具。

    功能：
    1. 使用正则表达式匹配文件名中的模式
    2. 支持变量替换：{seq}序号、{date}日期、{match}捕获组、{ext}扩展名
    3. 支持序号填充：{seq:03d} 补零到3位
    4. 支持日期格式化：{date:%Y%m%d}
    5. 冲突检测和跳过

    参数:
        dir_path: 目标目录
        pattern: 正则匹配模式
        rename_rule: 重命名规则字符串
        dry_run: 是否仅预览不实际执行

    返回:
        包含重命名计划、执行结果、冲突信息的字典
    """
    if not os.path.isdir(dir_path):
        return {"error": "目录不存在"}

    files = sorted([f for f in os.listdir(dir_path) if os.path.isfile(os.path.join(dir_path, f))])
    regex = re.compile(pattern)

    rename_plan = []
    used_names = set()
    conflicts = []
    seq = 0

    for fname in files:
        match = regex.search(fname)
        if not match:
            continue

        seq += 1
        # 获取文件扩展名
        ext = os.path.splitext(fname)[1]

        # 变量替换
        new_name = rename_rule
        replacements = {
            '{seq}': str(seq),
            '{ext}': ext,
            '{match}': match.group(0),
            '{date}': datetime.now().strftime('%Y%m%d'),
            '{datetime}': datetime.now().strftime('%Y%m%d_%H%M%S'),
        }

        # 处理序号填充 {seq:03d}
        pad_match = re.search(r'\{seq:(\d+)d\}', new_name)
        if pad_match:
            width = int(pad_match.group(1))
            new_name = new_name.replace(pad_match.group(0), str(seq).zfill(width))

        # 处理日期格式 {date:format}
        date_match = re.search(r'\{date:([^}]+)\}', new_name)
        if date_match:
            fmt = date_match.group(1)
            new_name = new_name.replace(date_match.group(0), datetime.now().strftime(fmt))

        # 处理捕获组 {g1}, {g2}, ...
        for gi in range(1, match.re.groups + 1):
            try:
                grp = match.group(gi) or ''
                new_name = new_name.replace(f'{{g{gi}}}', grp)
            except (IndexError, Exception):
                pass

        # 通用替换
        for key, val in replacements.items():
            if key in new_name:
                new_name = new_name.replace(key, val)

        # 确保有扩展名
        if not new_name.endswith(ext) and ext:
            new_name += ext

        old_path = os.path.join(dir_path, fname)
        new_path = os.path.join(dir_path, new_name)

        # 冲突检测
        if new_name in used_names:
            conflicts.append({"original": fname, "conflict_with": new_name})
            # 自动添加序号解决冲突
            base, ext2 = os.path.splitext(new_name)
            i = 1
            while f"{base}_{i}{ext2}" in used_names:
                i += 1
            new_name = f"{base}_{i}{ext2}"
            new_path = os.path.join(dir_path, new_name)

        used_names.add(new_name)

        entry = {
            "original": fname,
            "new_name": new_name,
            "match_groups": [match.group(i) for i in range(match.re.groups + 1)],
            "action": "rename"
        }
        rename_plan.append(entry)

        # 实际执行重命名
        if not dry_run and old_path != new_path:
            try:
                os.rename(old_path, new_path)
                entry["status"] = "success"
            except OSError as e:
                entry["status"] = f"error: {e}"
        else:
            entry["status"] = "preview" if dry_run else "skipped"

    return {
        "directory": dir_path,
        "pattern": pattern,
        "rename_rule": rename_rule,
        "total_files": len(files),
        "matched_files": len(rename_plan),
        "dry_run": dry_run,
        "rename_plan": rename_plan,
        "conflicts": conflicts,
        "conflict_count": len(conflicts)
    }


# ============================================================================
# 5. excel_data_merger - 多Excel文件智能合并
# ============================================================================

def _parse_xlsx(filepath: str) -> Tuple[List[str], List[List[Any]]]:
    """
    解析XLSX文件（ZIP+XML格式）。
    算法：
    1. XLSX是ZIP压缩包，包含xl/worksheets/sheet1.xml等
    2. 解析共享字符串表 xl/sharedStrings.xml
    3. 解析工作表XML，将单元格引用转换为行列号
    4. 用共享字符串填充文本单元格
    """
    try:
        with zipfile.ZipFile(filepath, 'r') as zf:
            # 读取共享字符串
            shared_strings = []
            try:
                ss_xml = zf.read('xl/sharedStrings.xml')
                root = ET.fromstring(ss_xml)
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in root.findall('.//main:si', ns):
                    texts = si.findall('.//main:t', ns)
                    shared_strings.append(''.join(t.text or '' for t in texts))
            except KeyError:
                pass

            # 读取第一个工作表
            sheet_name = 'xl/worksheets/sheet1.xml'
            try:
                sheet_xml = zf.read(sheet_name)
            except KeyError:
                # 尝试其他名称
                names = [n for n in zf.namelist() if n.startswith('xl/worksheets/sheet')]
                if not names:
                    return [], []
                sheet_xml = zf.read(names[0])

            root = ET.fromstring(sheet_xml)
            ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

            # 解析行和单元格
            rows_data = []
            for row in root.findall('.//main:row', ns):
                row_data = {}
                max_col = 0
                for cell in row.findall('main:c', ns):
                    ref = cell.get('r', '')  # 如 "A1", "B1"
                    col_letters = re.match(r'([A-Z]+)', ref)
                    if not col_letters:
                        continue
                    # 将列字母转换为列号
                    col_str = col_letters.group(1)
                    col_idx = 0
                    for ch in col_str:
                        col_idx = col_idx * 26 + (ord(ch) - ord('A') + 1)
                    col_idx -= 1  # 0-based
                    max_col = max(max_col, col_idx)

                    cell_type = cell.get('t', '')
                    val_elem = cell.find('main:v', ns)
                    if val_elem is None:
                        row_data[col_idx] = ''
                        continue

                    val = val_elem.text or ''
                    if cell_type == 's':
                        # 共享字符串索引
                        idx = int(val)
                        row_data[col_idx] = shared_strings[idx] if idx < len(shared_strings) else val
                    elif cell_type == 'b':
                        row_data[col_idx] = 'TRUE' if val == '1' else 'FALSE'
                    else:
                        # 尝试转换为数值
                        try:
                            if '.' in val:
                                row_data[col_idx] = float(val)
                            else:
                                row_data[col_idx] = int(val)
                        except ValueError:
                            row_data[col_idx] = val

                # 填充为列表
                row_list = [row_data.get(i, '') for i in range(max_col + 1)]
                rows_data.append(row_list)

            if not rows_data:
                return [], []

            return rows_data[0], rows_data[1:]
    except (zipfile.BadZipFile, ET.ParseError, KeyError):
        return [], []


def _detect_column_type(values: List[Any]) -> str:
    """推断列数据类型：int/float/str/date"""
    int_count = float_count = date_count = str_count = 0
    for v in values:
        if v == '' or v is None:
            continue
        if isinstance(v, int) or (isinstance(v, str) and v.lstrip('-').isdigit()):
            int_count += 1
        elif isinstance(v, float) or (isinstance(v, str) and _is_float(v)):
            float_count += 1
        else:
            str_count += 1

    total = int_count + float_count + date_count + str_count
    if total == 0:
        return 'empty'
    if int_count / total > 0.7:
        return 'int'
    if (int_count + float_count) / total > 0.7:
        return 'float'
    return 'str'


def _is_float(val: str) -> bool:
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False


def excel_data_merger(file_list: List[str], merge_key: str,
                      merge_method: str = 'outer') -> Dict[str, Any]:
    """
    多Excel文件智能合并。

    算法：
    1. 解析每个XLSX文件（ZIP+XML解析）
    2. 建立哈希索引：以merge_key列值为键
    3. 根据连接类型执行合并：
       - inner: 只保留两表都有的键
       - left: 保留左表所有键，右表匹配不到的填null
       - right: 保留右表所有键
       - outer: 保留所有键
    4. 自动检测每列数据类型

    参数:
        file_list: XLSX文件路径列表
        merge_key: 合并键列名
        merge_method: 合并方法 (inner/left/right/outer)

    返回:
        包含合并后数据、统计信息的字典
    """
    if len(file_list) < 2:
        return {"error": "至少需要2个文件"}

    # 解析所有文件
    all_data = []
    all_headers = []
    for fpath in file_list:
        headers, rows = _parse_xlsx(fpath)
        if not headers:
            continue
        all_headers.append(headers)
        all_data.append({"file": os.path.basename(fpath), "headers": headers, "rows": rows})

    if len(all_data) < 2:
        return {"error": "无法解析足够的Excel文件"}

    # 找到merge_key列在每个文件中的索引
    key_indices = []
    for d in all_data:
        if merge_key in d["headers"]:
            key_indices.append(d["headers"].index(merge_key))
        else:
            key_indices.append(-1)

    if any(k == -1 for k in key_indices):
        return {"error": f"合并键 '{merge_key}' 在某些文件中不存在"}

    # 合并表头（去重）
    merged_headers = list(all_data[0]["headers"])
    for d in all_data[1:]:
        for h in d["headers"]:
            if h not in merged_headers:
                merged_headers.append(h)

    # 构建哈希索引
    # data1的索引：key_value -> row
    index1 = {}
    for row in all_data[0]["rows"]:
        key_val = str(row[key_indices[0]]) if key_indices[0] < len(row) else ''
        if key_val:
            index1[key_val] = row

    index2 = {}
    for row in all_data[1]["rows"]:
        key_val = str(row[key_indices[1]]) if key_indices[1] < len(row) else ''
        if key_val:
            index2[key_val] = row

    # 执行合并
    all_keys = set()
    if merge_method in ('inner', 'left'):
        all_keys = set(index1.keys()) & set(index2.keys()) if merge_method == 'inner' else set(index1.keys())
    elif merge_method == 'right':
        all_keys = set(index2.keys())
    elif merge_method == 'outer':
        all_keys = set(index1.keys()) | set(index2.keys())

    merged_rows = []
    match_stats = {"both": 0, "left_only": 0, "right_only": 0}

    for key in sorted(all_keys):
        row1 = index1.get(key)
        row2 = index2.get(key)
        merged_row = [''] * len(merged_headers)

        if row1:
            for i, val in enumerate(row1):
                col_name = all_data[0]["headers"][i]
                if col_name in merged_headers:
                    merged_row[merged_headers.index(col_name)] = val

        if row2:
            for i, val in enumerate(row2):
                col_name = all_data[1]["headers"][i]
                if col_name in merged_headers:
                    merged_row[merged_headers.index(col_name)] = val

        # 统计
        if row1 and row2:
            match_stats["both"] += 1
        elif row1:
            match_stats["left_only"] += 1
        else:
            match_stats["right_only"] += 1

        merged_rows.append(merged_row)

    # 检测列类型
    column_types = {}
    for ci, col_name in enumerate(merged_headers):
        col_values = [r[ci] for r in merged_rows if ci < len(r)]
        column_types[col_name] = _detect_column_type(col_values)

    return {
        "merge_key": merge_key,
        "merge_method": merge_method,
        "files": [d["file"] for d in all_data],
        "merged_headers": merged_headers,
        "merged_rows": merged_rows,
        "row_count": len(merged_rows),
        "column_types": column_types,
        "match_stats": match_stats
    }


# ============================================================================
# 6. meeting_minutes_parser - 会议纪要结构化解析
# ============================================================================

def meeting_minutes_parser(text: str) -> Dict[str, Any]:
    """
    会议纪要结构化解析器。

    使用正则表达式+规则引擎从非结构化会议纪要文本中提取：
    1. 会议基本信息（时间、地点、参会人员）
    2. 议题列表
    3. 决议事项
    4. 待办事项（任务、责任人、截止日期）

    参数:
        text: 会议纪要原始文本

    返回:
        结构化会议纪要字典
    """
    result = {
        "meeting_info": {},
        "agenda_items": [],
        "decisions": [],
        "action_items": [],
        "attendees": [],
        "raw_text_length": len(text)
    }

    # --- 1. 提取会议基本信息 ---
    # 会议时间
    time_patterns = [
        r'(?:会议时间|时间|日期)[:：\s]*(\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?(?:\s+\d{1,2}[:：]\d{2})?)',
        r'(\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:\s+\d{1,2}[:：]\d{2})?)',
    ]
    for pat in time_patterns:
        m = re.search(pat, text)
        if m:
            result["meeting_info"]["time"] = m.group(1).strip()
            break

    # 会议地点
    loc_match = re.search(r'(?:会议地点|地点|会议室)[:：\s]*([^\n\r，,。]+)', text)
    if loc_match:
        result["meeting_info"]["location"] = loc_match.group(1).strip()

    # 会议主题/标题
    title_match = re.search(r'(?:会议主题|主题|会议名称|标题)[:：\s]*([^\n\r]+)', text)
    if title_match:
        result["meeting_info"]["title"] = title_match.group(1).strip()
    else:
        # 尝试从第一行提取标题
        first_line = text.strip().split('\n')[0].strip()
        if first_line and len(first_line) < 50:
            result["meeting_info"]["title"] = first_line

    # --- 2. 提取参会人员 ---
    attendee_patterns = [
        r'(?:参会人员|参加人员|与会人员|出席人员)[:：\s]*([^\n\r]+)',
        r'(?:主持人|主持)[:：\s]*([^\n\r，,]+)',
    ]
    for pat in attendee_patterns:
        matches = re.findall(pat, text)
        for m in matches:
            # 分割人名
            names = re.split(r'[，,、\s]+', m.strip())
            for name in names:
                name = name.strip()
                if name and len(name) <= 10 and name not in result["attendees"]:
                    result["attendees"].append(name)

    # --- 3. 提取议题 ---
    # 匹配 "议题1:", "1.", "(1)", "一、" 等
    agenda_patterns = [
        r'(?:议题|议程)\s*[\d一二三四五六七八九十]+[、\.）)]\s*[：:]?\s*(.+?)(?=\n|$)',
        r'(?:议题|议程)\s*[:：]\s*(.+?)(?=\n|$)',
    ]
    for pat in agenda_patterns:
        matches = re.findall(pat, text)
        for m in matches:
            item = m.strip()
            if item and item not in result["agenda_items"]:
                result["agenda_items"].append(item)

    # 也尝试匹配编号列表项作为议题
    numbered_items = re.findall(r'(?:^|\n)\s*(?:\d+|[一二三四五六七八九十]+)[、\.）)]\s*(.+?)(?=\n|$)', text)
    for item in numbered_items:
        item = item.strip()
        # 过滤掉明显是待办或决议的条目
        if item and not any(kw in item for kw in ['负责', '完成', '截止', '决议', '决定']):
            if item not in result["agenda_items"] and len(item) < 100:
                result["agenda_items"].append(item)

    # --- 4. 提取决议 ---
    decision_patterns = [
        r'(?:决议|决定|结论|会议决定|经讨论)[:：\s]*([^\n\r]+)',
        r'(?:一致(?:同意|认为|决定|通过))[:：\s]*([^\n\r]+)',
        r'(?:讨论(?:后)?(?:决定|认为|同意))[:：\s]*([^\n\r]+)',
    ]
    for pat in decision_patterns:
        matches = re.findall(pat, text)
        for m in matches:
            dec = m.strip()
            if dec and dec not in result["decisions"]:
                result["decisions"].append(dec)

    # --- 5. 提取待办事项（任务+责任人+截止日期） ---
    # 模式：任务描述 + 责任人 + 截止日期
    action_patterns = [
        # "由XXX负责完成YYY，截止日期Z"
        r'(?:由)?(.{2,5})\s*(?:负责|完成|跟进|落实|执行|牵头|落实)\s*(.+?)(?:[，,。；;]|$)\s*(?:截止|deadline|期限|时间)[:：\s]*(\d{1,4}[-/年]\d{1,2}[-/月]\d{1,2}日?)?',
        # "XXX：完成YYY，Y月Z日前"
        r'(.{2,5})\s*[:：]\s*(.+?)(?:[，,。；;]|$)\s*(?:截止|前|deadline)[:：\s]*(\d{1,4}[-/年]\d{1,2}[-/月]\d{1,2}日?)?',
        # "待办/行动项：XXX完成YYY"
        r'(?:待办|行动项|TODO|todo|待办事项)[:：\s]*(.+?)(?:[，,；;]|$)',
    ]

    for pat in action_patterns:
        matches = re.findall(pat, text)
        for m in matches:
            if isinstance(m, tuple):
                owner, task, deadline = m[0].strip(), m[1].strip(), (m[2].strip() if len(m) > 2 else '')
                if task and task not in [a.get("task", "") for a in result["action_items"]]:
                    result["action_items"].append({
                        "owner": owner if owner else "未指定",
                        "task": task,
                        "deadline": deadline if deadline else "未指定"
                    })
            else:
                task = m.strip()
                if task and task not in [a.get("task", "") for a in result["action_items"]]:
                    result["action_items"].append({
                        "owner": "未指定",
                        "task": task,
                        "deadline": "未指定"
                    })

    # 尝试从待办中提取截止日期
    deadline_pattern = re.findall(r'(?:截止|deadline|期限|前)[:：\s]*(\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?)', text)
    if deadline_pattern and not result["action_items"]:
        result["action_items"].append({
            "owner": "未指定",
            "task": "查看原文",
            "deadline": deadline_pattern[0]
        })

    return result


# ============================================================================
# 7. schedule_conflict_detector - 排班冲突检测器（区间树算法）
# ============================================================================

class IntervalNode:
    """区间树节点"""
    __slots__ = ['interval', 'max_end', 'left', 'right', 'height']

    def __init__(self, interval):
        self.interval = interval  # (start, end, label)
        self.max_end = interval[1]
        self.left = None
        self.right = None
        self.height = 1


class IntervalTree:
    """
    红黑树风格的平衡区间树（AVL平衡）。
    支持O(log n)区间查询和重叠检测。
    """

    def __init__(self):
        self.root = None

    def _height(self, node):
        return node.height if node else 0

    def _balance_factor(self, node):
        return self._height(node.left) - self._height(node.right) if node else 0

    def _update(self, node):
        node.height = 1 + max(self._height(node.left), self._height(node.right))
        node.max_end = node.interval[1]
        if node.left:
            node.max_end = max(node.max_end, node.left.max_end)
        if node.right:
            node.max_end = max(node.max_end, node.right.max_end)

    def _rotate_right(self, y):
        x = y.left
        T2 = x.right
        x.right = y
        y.left = T2
        self._update(y)
        self._update(x)
        return x

    def _rotate_left(self, x):
        y = x.right
        T2 = y.left
        y.left = x
        x.right = T2
        self._update(x)
        self._update(y)
        return y

    def _insert(self, node, interval):
        if not node:
            return IntervalNode(interval)

        if interval[0] < node.interval[0]:
            node.left = self._insert(node.left, interval)
        else:
            node.right = self._insert(node.right, interval)

        # 更新高度和max_end
        self._update(node)

        # AVL平衡
        bf = self._balance_factor(node)
        if bf > 1 and interval[0] < node.left.interval[0]:
            return self._rotate_right(node)
        if bf < -1 and interval[0] >= node.right.interval[0]:
            return self._rotate_left(node)
        if bf > 1 and interval[0] >= node.left.interval[0]:
            node.left = self._rotate_left(node.left)
            return self._rotate_right(node)
        if bf < -1 and interval[0] < node.right.interval[0]:
            node.right = self._rotate_right(node.right)
            return self._rotate_left(node)

        return node

    def insert(self, interval):
        """插入区间 (start, end, label)"""
        self.root = self._insert(self.root, interval)

    def _query_overlap(self, node, interval, results):
        """查询与给定区间重叠的所有区间"""
        if not node:
            return

        # 检查当前节点是否重叠
        if node.interval[0] < interval[1] and node.interval[1] > interval[0]:
            results.append(node.interval)

        # 如果左子树的最大结束值大于区间开始值，搜索左子树
        if node.left and node.left.max_end > interval[0]:
            self._query_overlap(node.left, interval, results)

        # 如果右子树的最小开始值小于区间结束值，搜索右子树
        if node.right and node.interval[0] < interval[1]:
            self._query_overlap(node.right, interval, results)

    def query_overlap(self, interval):
        """返回与给定区间重叠的所有区间列表"""
        results = []
        self._query_overlap(self.root, interval, results)
        return results


def schedule_conflict_detector(schedules: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    排班冲突检测器。

    使用区间树（Interval Tree）数据结构高效检测时间重叠。
    算法：
    1. 将每个排班项转换为时间区间
    2. 构建AVL平衡区间树
    3. 对每个区间查询重叠
    4. 生成冲突报告

    时间复杂度：O(n log n) 构建 + O(k log n) 查询（k为冲突数）

    参数:
        schedules: 排班列表，每项包含 person, start, end, task

    返回:
        包含冲突检测结果的字典
    """
    tree = IntervalTree()
    intervals = []
    conflicts = []
    all_person_schedules = defaultdict(list)

    # 构建区间
    for s in schedules:
        start = s.get("start")
        end = s.get("end")
        person = s.get("person", "未知")
        task = s.get("task", "")

        # 支持字符串时间，转换为时间戳
        if isinstance(start, str):
            try:
                start = datetime.strptime(start, "%Y-%m-%d %H:%M")
            except ValueError:
                try:
                    start = datetime.strptime(start, "%Y-%m-%d")
                except ValueError:
                    continue
        if isinstance(end, str):
            try:
                end = datetime.strptime(end, "%Y-%m-%d %H:%M")
            except ValueError:
                try:
                    end = datetime.strptime(end, "%Y-%m-%d")
                except ValueError:
                    continue

        # 转换为数值（时间戳）
        start_ts = start.timestamp() if hasattr(start, 'timestamp') else float(start)
        end_ts = end.timestamp() if hasattr(end, 'timestamp') else float(end)

        if end_ts <= start_ts:
            continue

        interval = (start_ts, end_ts, f"{person}:{task}")
        intervals.append(interval)
        all_person_schedules[person].append(interval)
        tree.insert(interval)

    # 检测冲突
    conflict_pairs = set()
    for interval in intervals:
        overlaps = tree.query_overlap(interval)
        for ov in overlaps:
            if ov != interval:
                pair = tuple(sorted([interval, ov], key=lambda x: x[2]))
                if pair not in conflict_pairs:
                    conflict_pairs.add(pair)
                    # 计算重叠时长
                    overlap_start = max(interval[0], ov[0])
                    overlap_end = min(interval[1], ov[1])
                    overlap_duration = (overlap_end - overlap_start) / 3600  # 小时

                    p1, t1 = interval[2].split(':', 1)
                    p2, t2 = ov[2].split(':', 1)

                    conflicts.append({
                        "person1": p1,
                        "task1": t1,
                        "person2": p2,
                        "task2": t2,
                        "overlap_start": datetime.fromtimestamp(overlap_start).strftime("%Y-%m-%d %H:%M"),
                        "overlap_end": datetime.fromtimestamp(overlap_end).strftime("%Y-%m-%d %H:%M"),
                        "overlap_hours": round(overlap_duration, 2)
                    })

    # 检测同一人的排班冲突
    person_conflicts = []
    for person, p_intervals in all_person_schedules.items():
        if len(p_intervals) < 2:
            continue
        for i in range(len(p_intervals)):
            for j in range(i + 1, len(p_intervals)):
                a, b = p_intervals[i], p_intervals[j]
                if a[0] < b[1] and a[1] > b[0]:
                    person_conflicts.append({
                        "person": person,
                        "conflict": f"{a[2]} 与 {b[2]} 时间重叠"
                    })

    return {
        "total_schedules": len(schedules),
        "total_intervals": len(intervals),
        "conflict_count": len(conflicts),
        "conflicts": conflicts,
        "person_self_conflicts": person_conflicts,
        "has_conflicts": len(conflicts) > 0 or len(person_conflicts) > 0,
        "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }


# ============================================================================
# 8. text_diff_analyzer - 文本差异分析（Myers差异算法）
# ============================================================================

def _myers_diff(a: List[str], b: List[str]) -> List[Tuple[str, str]]:
    """
    Myers差异算法实现。

    算法核心：找到从(0,0)到(N,M)的最短编辑路径（SES）。
    使用V数组存储对角线k上的最远x坐标。
    
    时间复杂度：O(ND)，D为编辑距离
    空间复杂度：O(D)

    返回操作序列：[('equal'/'insert'/'delete', line), ...]
    """
    N, M = len(a), len(b)
    max_d = N + M
    offset = max_d

    # V数组：V[k] = x坐标（k为对角线编号）
    v = {1: 0}
    trace = []

    for d in range(max_d + 1):
        v_copy = dict(v)
        trace.append(v_copy)

        for k in range(-d, d + 1, 2):
            # 选择前进方向
            if k == -d or (k != d and v.get(k - 1, 0) < v.get(k + 1, 0)):
                x = v.get(k + 1, 0)  # 向下（插入）
            else:
                x = v.get(k - 1, 0) + 1  # 向右（删除）

            y = x - k

            # 沿对角线前进
            while x < N and y < M and a[x] == b[y]:
                x += 1
                y += 1

            v[k] = x

            if x >= N and y >= M:
                # 到达终点，回溯生成操作序列
                return _myers_backtrace(trace, a, b, N, M)

    return [('equal', line) for line in a]


def _myers_backtrace(trace: List[dict], a: List[str], b: List[str],
                      N: int, M: int) -> List[Tuple[str, str]]:
    """回溯Myers算法的编辑路径，生成操作序列。"""
    operations = []
    x, y = N, M

    for d in range(len(trace) - 1, 0, -1):
        v = trace[d]
        k = x - y

        if k == -d or (k != d and v.get(k - 1, 0) < v.get(k + 1, 0)):
            prev_k = k + 1  # 插入
        else:
            prev_k = k - 1  # 删除

        prev_x = v.get(prev_k, 0)
        prev_y = prev_x - prev_k

        # 对角线移动（相等）
        while x > prev_x and y > prev_y:
            operations.append(('equal', a[x - 1]))
            x -= 1
            y -= 1

        if d > 0:
            if x == prev_x:
                # 插入操作
                operations.append(('insert', b[y - 1]))
            else:
                # 删除操作
                operations.append(('delete', a[x - 1]))

        x, y = prev_x, prev_y

    # 处理起点
    while x > 0 and y > 0 and a[x - 1] == b[y - 1]:
        operations.append(('equal', a[x - 1]))
        x -= 1
        y -= 1

    operations.reverse()
    return operations


def text_diff_analyzer(text1: str, text2: str) -> Dict[str, Any]:
    """
    文本差异分析，实现Myers差异算法。

    算法：Myers算法寻找两个序列间的最短编辑脚本(SES)，
    使用对角线k上的最远x坐标表示搜索状态。

    参数:
        text1: 原始文本
        text2: 修改后文本

    返回:
        包含差异操作、统计信息、格式化diff的字典
    """
    lines1 = text1.splitlines()
    lines2 = text2.splitlines()

    # 执行Myers差异算法
    operations = _myers_diff(lines1, lines2)

    # 统计
    stats = Counter(op for op, _ in operations)

    # 生成格式化diff
    diff_lines = []
    for op, line in operations:
        if op == 'equal':
            diff_lines.append(f"  {line}")
        elif op == 'insert':
            diff_lines.append(f"+ {line}")
        elif op == 'delete':
            diff_lines.append(f"- {line}")

    # 合并连续的插入/删除为"修改"
    changes = []
    i = 0
    while i < len(operations):
        op, line = operations[i]
        if op == 'delete':
            deleted = [line]
            j = i + 1
            while j < len(operations) and operations[j][0] == 'delete':
                deleted.append(operations[j][1])
                j += 1
            inserted = []
            while j < len(operations) and operations[j][0] == 'insert':
                inserted.append(operations[j][1])
                j += 1
            if inserted:
                changes.append({
                    "type": "modify",
                    "old_lines": deleted,
                    "new_lines": inserted
                })
            else:
                changes.append({"type": "delete", "lines": deleted})
            i = j
        elif op == 'insert':
            inserted = [line]
            j = i + 1
            while j < len(operations) and operations[j][0] == 'insert':
                inserted.append(operations[j][1])
                j += 1
            changes.append({"type": "insert", "lines": inserted})
            i = j
        else:
            i += 1

    return {
        "total_lines_1": len(lines1),
        "total_lines_2": len(lines2),
        "operations": operations,
        "stats": {
            "equal": stats.get('equal', 0),
            "insert": stats.get('insert', 0),
            "delete": stats.get('delete', 0),
            "total_changes": stats.get('insert', 0) + stats.get('delete', 0)
        },
        "diff_text": '\n'.join(diff_lines),
        "changes": changes,
        "similarity": round(stats.get('equal', 0) / max(len(lines1), len(lines2), 1), 4)
    }


# ============================================================================
# 9. file_duplicate_finder - 文件查重器
# ============================================================================

def _compute_md5(filepath: str, chunk_size: int = 8192) -> str:
    """计算文件MD5哈希值（分块读取避免内存溢出）。"""
    md5 = hashlib.md5()
    try:
        with open(filepath, 'rb') as f:
            while True:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                md5.update(chunk)
        return md5.hexdigest()
    except (IOError, OSError):
        return ""


def _string_similarity(s1: str, s2: str) -> float:
    """基于编辑距离的字符串相似度（归一化到[0,1]）。"""
    if not s1 and not s2:
        return 1.0
    dist = _levenshtein_distance(s1.lower(), s2.lower())
    return 1.0 - dist / max(len(s1), len(s2), 1)


def file_duplicate_finder(dir_path: str, fuzzy_threshold: float = 0.85) -> Dict[str, Any]:
    """
    文件查重器。

    算法流程：
    1. 第一遍扫描：按文件大小分组（快速过滤）
    2. 第二遍扫描：对大小相同的文件计算MD5哈希
    3. 精确匹配：MD5哈希值完全相同的文件为精确重复
    4. 模糊匹配：文件名相似度超过阈值的文件为疑似重复
    5. 生成查重报告，计算可节省空间

    参数:
        dir_path: 目标目录
        fuzzy_threshold: 模糊匹配阈值（0-1）

    返回:
        包含精确重复、模糊重复、统计信息的字典
    """
    if not os.path.isdir(dir_path):
        return {"error": "目录不存在"}

    # 收集文件信息
    files_info = []
    for root, dirs, files in os.walk(dir_path):
        for fname in files:
            fpath = os.path.join(root, fname)
            try:
                size = os.path.getsize(fpath)
                files_info.append({
                    "path": fpath,
                    "name": fname,
                    "size": size,
                    "relative_path": os.path.relpath(fpath, dir_path)
                })
            except OSError:
                continue

    # 步骤1：按文件大小分组
    size_groups = defaultdict(list)
    for fi in files_info:
        size_groups[fi["size"]].append(fi)

    # 步骤2：对大小相同的文件计算MD5
    exact_duplicates = []
    hash_groups = defaultdict(list)

    for size, group in size_groups.items():
        if len(group) < 2:
            continue
        for fi in group:
            fi["md5"] = _compute_md5(fi["path"])
            hash_groups[fi["md5"]].append(fi)

    # 步骤3：提取精确重复
    for md5, group in hash_groups.items():
        if len(group) >= 2:
            exact_duplicates.append({
                "md5": md5,
                "size": group[0]["size"],
                "files": [fi["relative_path"] for fi in group],
                "wasted_space": group[0]["size"] * (len(group) - 1)
            })

    # 步骤4：模糊匹配（文件名相似度）
    fuzzy_duplicates = []
    all_names = [(fi["name"], fi["relative_path"], fi["size"]) for fi in files_info]

    for i in range(len(all_names)):
        for j in range(i + 1, len(all_names)):
            name1, path1, size1 = all_names[i]
            name2, path2, size2 = all_names[j]

            # 跳过已确定为精确重复的
            if size1 == size2:
                md5_1 = next((fi["md5"] for fi in files_info if fi["name"] == name1), None)
                md5_2 = next((fi["md5"] for fi in files_info if fi["name"] == name2), None)
                if md5_1 and md5_2 and md5_1 == md5_2:
                    continue

            sim = _string_similarity(name1, name2)
            if sim >= fuzzy_threshold:
                fuzzy_duplicates.append({
                    "file1": path1,
                    "file2": path2,
                    "name_similarity": round(sim, 4),
                    "size1": size1,
                    "size2": size2
                })

    # 排序
    fuzzy_duplicates.sort(key=lambda x: -x["name_similarity"])

    # 统计
    total_wasted = sum(d["wasted_space"] for d in exact_duplicates)

    return {
        "directory": dir_path,
        "total_files": len(files_info),
        "exact_duplicate_groups": len(exact_duplicates),
        "exact_duplicate_files": sum(len(d["files"]) for d in exact_duplicates),
        "fuzzy_duplicate_pairs": len(fuzzy_duplicates),
        "exact_duplicates": exact_duplicates,
        "fuzzy_duplicates": fuzzy_duplicates[:50],  # 限制输出数量
        "wasted_space_bytes": total_wasted,
        "wasted_space_mb": round(total_wasted / (1024 * 1024), 2)
    }


# ============================================================================
# 10. document_summary_generator - 文档摘要生成器（TextRank算法）
# ============================================================================

def _split_sentences(text: str) -> List[str]:
    """句子分割器，支持中英文标点。"""
    # 替换换行符为空格
    text = re.sub(r'\s+', ' ', text)
    # 按句号、问号、感叹号分割
    sentences = re.split(r'[。！？!?\.]+', text)
    # 过滤过短的句子
    return [s.strip() for s in sentences if len(s.strip()) > 10]


def _sentence_similarity_tfidf(s1: str, s2: str, idf: Dict[str, float]) -> float:
    """
    基于TF-IDF的句子相似度。
    使用两个句子的TF-IDF向量的余弦相似度。
    """
    tokens1 = _tokenize(s1)
    tokens2 = _tokenize(s2)

    if not tokens1 or not tokens2:
        return 0.0

    tf1 = Counter(tokens1)
    tf2 = Counter(tokens2)

    # 计算TF-IDF
    vec1 = {t: tf1[t] / len(tokens1) * idf.get(t, 1.0) for t in tf1}
    vec2 = {t: tf2[t] / len(tokens2) * idf.get(t, 1.0) for t in tf2}

    return _cosine_similarity(vec1, vec2)


def _pagerank(similarity_matrix: List[List[float]], damping: float = 0.85,
               max_iter: int = 100, tol: float = 1e-6) -> List[float]:
    """
    PageRank迭代算法。

    算法：
    PR(i) = (1-d) + d * Σ(PR(j) * sim(i,j) / Σ(sim(j,k))
    
    其中d为阻尼系数，sim为归一化的相似度。

    参数:
        similarity_matrix: 句子相似度矩阵
        damping: 阻尼系数
        max_iter: 最大迭代次数
        tol: 收敛阈值

    返回:
        每个句子的PageRank分数
    """
    n = len(similarity_matrix)
    if n == 0:
        return []

    # 初始化分数
    scores = [1.0 / n] * n

    # 归一化相似度矩阵（每行归一化）
    norm_matrix = []
    for i in range(n):
        row_sum = sum(similarity_matrix[i])
        if row_sum > 0:
            norm_matrix.append([s / row_sum for s in similarity_matrix[i]])
        else:
            norm_matrix.append([0.0] * n)

    # 迭代
    for iteration in range(max_iter):
        new_scores = []
        for i in range(n):
            # PR(i) = (1-d)/n + d * Σ(PR(j) * norm_sim(j,i))
            rank_sum = 0.0
            for j in range(n):
                if i != j:
                    rank_sum += scores[j] * norm_matrix[j][i]
            new_score = (1 - damping) / n + damping * rank_sum
            new_scores.append(new_score)

        # 检查收敛
        diff = sum(abs(new_scores[i] - scores[i]) for i in range(n))
        scores = new_scores
        if diff < tol:
            break

    return scores


def document_summary_generator(doc_path: str, num_sentences: int = 5) -> Dict[str, Any]:
    """
    文档摘要生成器，实现TextRank算法。

    TextRank算法流程：
    1. 文本预处理：句子分割
    2. 构建句子相似度矩阵（基于TF-IDF余弦相似度）
    3. 使用PageRank算法对句子进行图排序
    4. 选取Top-K句子作为摘要
    5. 按原文顺序排列输出摘要

    时间复杂度：O(n^2 * m) n=句子数 m=平均词数
    空间复杂度：O(n^2) 相似度矩阵

    参数:
        doc_path: 文档路径
        num_sentences: 摘要句子数

    返回:
        包含摘要、关键句排序、统计信息的字典
    """
    text = _read_file_content(doc_path)
    if not text.strip():
        return {"error": "无法读取文档内容或文档为空"}

    # 步骤1：句子分割
    sentences = _split_sentences(text)
    if len(sentences) <= num_sentences:
        return {
            "summary": text[:1000],
            "sentences": sentences,
            "message": "文档过短，无需摘要"
        }

    # 步骤2：计算IDF
    all_tokens = [_tokenize(s) for s in sentences]
    N = len(sentences)
    df = Counter()
    for tokens in all_tokens:
        df.update(set(tokens))
    idf = {t: math.log(N / df[t]) if df[t] > 0 else 1.0 for t in df}

    # 步骤3：构建句子相似度矩阵
    n = len(sentences)
    sim_matrix = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            sim = _sentence_similarity_tfidf(sentences[i], sentences[j], idf)
            sim_matrix[i][j] = sim
            sim_matrix[j][i] = sim

    # 步骤4：PageRank迭代
    scores = _pagerank(sim_matrix, damping=0.85, max_iter=100)

    # 步骤5：选取Top-K句子
    ranked_sentences = sorted(
        [(i, scores[i], sentences[i]) for i in range(n)],
        key=lambda x: -x[1]
    )

    top_k = ranked_sentences[:num_sentences]
    # 按原文顺序排列
    top_k_sorted = sorted(top_k, key=lambda x: x[0])

    summary = '。'.join([s[2] for s in top_k_sorted]) + '。'
    top_keywords = _extract_keywords(all_tokens, idf, top_k=15)

    return {
        "file": os.path.basename(doc_path),
        "original_length": len(text),
        "summary_length": len(summary),
        "compression_ratio": round(len(summary) / max(len(text), 1), 4),
        "total_sentences": n,
        "summary_sentences": num_sentences,
        "summary": summary,
        "ranked_sentences": [
            {"rank": i + 1, "original_index": idx, "score": round(score, 6),
             "sentence": sent[:200]}
            for i, (idx, score, sent) in enumerate(ranked_sentences[:10])
        ],
        "top_keywords": top_keywords
    }


def _extract_keywords(all_tokens: List[List[str]], idf: Dict[str, float],
                       top_k: int = 10) -> List[Dict[str, Any]]:
    """提取关键词：TF-IDF排序。"""
    total_tf = Counter()
    for tokens in all_tokens:
        total_tf.update(tokens)

    keywords = []
    for term, tf in total_tf.items():
        tfidf_score = tf * idf.get(term, 1.0)
        keywords.append({"term": term, "tf": tf, "tfidf": round(tfidf_score, 4)})

    keywords.sort(key=lambda x: -x["tfidf"])
    return keywords[:top_k]


# ============================================================================
# 主程序测试
# ============================================================================

if __name__ == '__main__':
    print("=" * 60)
    print("office-helper 办公效率工具箱 - 功能测试")
    print("=" * 60)

    # 测试会议纪要解析
    print("\n--- 会议纪要解析测试 ---")
    sample_minutes = """
    会议主题：Q4项目规划讨论会
    会议时间：2025-01-15 14:00
    会议地点：3号会议室
    参会人员：张三、李四、王五、赵六

    议题1：Q4项目进度回顾
    议题2：下季度资源分配

    讨论决定：项目A按计划推进，项目B延期至Q1。
    经讨论决定：增加测试人员投入。

    待办事项：
    由张三负责完成需求文档，截止日期2025-01-20
    李四：跟进供应商报价，1月25日前
    王五负责部署测试环境，截止2025-01-18
    """
    result = meeting_minutes_parser(sample_minutes)
    print(f"会议主题: {result['meeting_info'].get('title', 'N/A')}")
    print(f"参会人员: {result['attendees']}")
    print(f"议题数: {len(result['agenda_items'])}")
    print(f"决议数: {len(result['decisions'])}")
    print(f"待办数: {len(result['action_items'])}")
    for item in result['action_items']:
        print(f"  -> {item['owner']}: {item['task']} (截止: {item['deadline']})")

    # 测试文本差异分析
    print("\n--- 文本差异分析测试 ---")
    text_a = "第一行\n第二行\n第三行\n第四行"
    text_b = "第一行\n第二行修改\n第三行\n第五行"
    diff = text_diff_analyzer(text_a, text_b)
    print(f"差异操作数: {diff['stats']['total_changes']}")
    print(f"相似度: {diff['similarity']}")
    print(f"Diff:\n{diff['diff_text']}")

    # 测试排班冲突检测
    print("\n--- 排班冲突检测测试 ---")
    schedules = [
        {"person": "张三", "start": "2025-01-15 09:00", "end": "2025-01-15 12:00", "task": "会议A"},
        {"person": "张三", "start": "2025-01-15 11:00", "end": "2025-01-15 14:00", "task": "会议B"},
        {"person": "李四", "start": "2025-01-15 13:00", "end": "2025-01-15 15:00", "task": "培训"},
    ]
    conflicts = schedule_conflict_detector(schedules)
    print(f"冲突数: {conflicts['conflict_count']}")
    for c in conflicts['conflicts']:
        print(f"  -> {c['person1']}({c['task1']}) 与 {c['person2']}({c['task2']}) 重叠{c['overlap_hours']}小时")

    # 测试TextRank摘要
    print("\n--- TextRank摘要测试 ---")
    sample_text = """
    人工智能是计算机科学的一个分支，它企图了解智能的实质，并生产出一种新的能以人类智能相似的方式做出反应的智能机器。
    人工智能的研究包括机器人、语言识别、图像识别、自然语言处理和专家系统等。
    自从人工智能诞生以来，其理论和技术日益成熟，应用领域也不断扩大。
    可以设想，未来人工智能带来的科技产品，将会是人类智慧的容器。
    人工智能可以对人的意识、思维的信息过程进行模拟。
    人工智能不是人的智能，但能像人那样思考、也可能超过人的智能。
    人工智能是一门极富挑战性的科学，从事这项工作的人必须懂得计算机知识、心理学和哲学。
    人工智能是包括十分广泛的科学，它由不同的领域组成，如机器学习、计算机视觉等。
    总的说来，人工智能研究的一个主要目标是使机器能够胜任一些通常需要人类智能才能完成的复杂工作。
    """
    # 写入临时文件测试
    import tempfile
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as f:
        f.write(sample_text)
        temp_path = f.name
    summary_result = document_summary_generator(temp_path, num_sentences=3)
    print(f"原文长度: {summary_result['original_length']}")
    print(f"摘要长度: {summary_result['summary_length']}")
    print(f"压缩比: {summary_result['compression_ratio']}")
    print(f"摘要: {summary_result['summary'][:200]}...")
    os.unlink(temp_path)

    print("\n" + "=" * 60)
    print("所有测试完成!")
    print("=" * 60)
