#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""textutil —— 降重流程用到的文本算法与启发式判定。

三块内容：
1. **连续重合检测**：知网连续 13 字相同即标红，所以「改写稿与原文的最长公共
   连续子串」是判断改没改干净的核心指标。``maximal_matches`` 找出所有
   ≥ k 字的极大公共子串。
2. **重合片段分类**：区分「合规保留」（引文原文、书名、专名全称）与
   「需继续处理」。不做这个分类会把无法规避的引用误报成问题，
   导致无意义的反复改写。
3. **段落角色识别**：把段落分成 标题/一级标题/摘要/关键词/正文/参考文献/空段，
   作为「该改哪些段、哪些段一律不动」的默认判据。

这里所有判定都是**启发式**，可以也应该被用户覆盖。它们是起点不是结论。
"""
from __future__ import annotations

import re

#: 知网标红阈值：连续相同字数达到此值即计入重复
DEFAULT_THRESHOLD = 13

# 文献类型标识（GB/T 7714），出现即认为该段是参考文献条目
CITATION_MARK = re.compile(r'[\[［](J|D|M|C|N|R|S|P|A|Z|G|EB/OL|DB/OL|J/OL)[\]］]')
REF_ITEM = re.compile(r'^\s*[\[［]\d+[\]］]')
HEADING_NUM = re.compile(r'^\s*(?:\d+(?:\.\d+)*[\.、]|[一二三四五六七八九十]+[、.])\s*\S')
ABSTRACT = re.compile(r'^\s*摘\s*要\s*[：:]')
KEYWORDS = re.compile(r'^\s*关\s*键\s*词\s*[：:]')
REF_TITLE = re.compile(r'^\s*(参考文献|References|REFERENCES)\s*$')


# ──────────────────────────────────────────────────────────────────────────
# 1. 最长公共连续子串
# ──────────────────────────────────────────────────────────────────────────
def lcs_length(a: str, b: str) -> tuple[int, str]:
    """返回 ``(长度, 子串)`` —— a 与 b 的最长公共连续子串。

    滚动数组实现，空间 O(len(b))。用于逐对比较单段文本，
    不要在整篇文档上直接调用。
    """
    if not a or not b:
        return 0, ''
    prev = [0] * (len(b) + 1)
    best, best_i = 0, 0
    for i in range(1, len(a) + 1):
        cur = [0] * (len(b) + 1)
        ai = a[i - 1]
        for j in range(1, len(b) + 1):
            if ai == b[j - 1]:
                v = prev[j - 1] + 1
                cur[j] = v
                if v > best:
                    best, best_i = v, i
        prev = cur
    return best, a[best_i - best:best_i]


def maximal_matches(a: str, b: str, k: int = DEFAULT_THRESHOLD) -> list[str]:
    """找出 b 中所有长度 ≥ k 且出现在 a 中的**极大**公共子串。

    "极大"指不能再向两端延伸。比逐位置 LCS 快得多，适合长文本。
    """
    res, n, i = [], len(b), 0
    while i < n:
        if len(b) - i >= k and b[i:i + k] in a:
            j = i + k
            while j <= n and b[i:j] in a:
                j += 1
            res.append(b[i:j - 1])
            i = j - 1
        else:
            i += 1
    return res


# ──────────────────────────────────────────────────────────────────────────
# 2. 重合片段分类
# ──────────────────────────────────────────────────────────────────────────
def classify_overlap(text: str, sub: str) -> str:
    """判断一个重合片段属于哪一类。

    :returns: ``quote`` 引文内 / ``title`` 书名号专名 / ``risk`` 需继续处理
    """
    if len(sub) >= 2 and sub.startswith(('“', '‘')) and sub.endswith(('”', '’')):
        return 'quote'
    idx = text.find(sub)
    while idx != -1:
        before = text[:idx]
        if before.count('“') > before.count('”'):
            return 'quote'
        # 落在书名号内（或片段自身含书名号）的也属专名：书名、标准全称、
        # 法规名等专用名称，无法也不该替换。
        if before.rfind('《') > before.rfind('》'):
            return 'title'
        idx = text.find(sub, idx + 1)
    if '《' in sub and '》' in sub:
        return 'title'
    return 'risk'


OVERLAP_LABEL = {'quote': '[引文内-合规]', 'title': '[专名-须保留]',
                 'risk': '!! 需处理  '}


# ──────────────────────────────────────────────────────────────────────────
# 3. 段落角色识别
# ──────────────────────────────────────────────────────────────────────────
ROLE_LABEL = {
    'title': '论文标题', 'heading': '小标题', 'abstract': '摘要',
    'keywords': '关键词', 'body': '正文', 'reference': '参考文献条目',
    'ref_title': '参考文献标题', 'empty': '空段', 'table': '表格', 'other': '其他',
}

#: 默认不建议改写的角色（改了要么违规，要么无意义）
LOCKED_ROLES = ('title', 'heading', 'keywords', 'reference', 'ref_title',
                'empty', 'table')


def classify_role(p: dict, total: int) -> str:
    """根据段落内容与格式猜测其角色。``p`` 来自 ``Document.paragraphs()``。"""
    text = (p.get('text') or '')
    t = text.strip()
    if p.get('kind') == 'table':
        return 'table'
    if not t:
        return 'empty'
    if REF_TITLE.match(t):
        return 'ref_title'
    if ABSTRACT.match(t):
        return 'abstract'
    if KEYWORDS.match(t):
        return 'keywords'
    if CITATION_MARK.search(t) or (REF_ITEM.match(t) and len(t) > 60):
        return 'reference'
    style = p.get('style') or ''
    if 'Heading' in style or '标题' in style:
        return 'heading'
    if HEADING_NUM.match(t) and len(t) <= 60:
        return 'heading'
    if p['p'] == 1 and len(t) <= 80:
        return 'title'
    if len(t) <= 80 and t.endswith(('。', '？', '！')) is False and len(t) <= 30:
        return 'heading'
    return 'body'


def annotate(paragraphs: list[dict]) -> list[dict]:
    """给 ``Document.paragraphs()`` 的结果补齐 role / 长短标记。"""
    total = len(paragraphs)
    for p in paragraphs:
        p.setdefault('role', classify_role(p, total))
        p['length'] = len(p.get('text') or '')
        p['locked'] = p['role'] in LOCKED_ROLES
    return paragraphs


def needs_rewrite(p: dict) -> bool:
    """默认「应该改写」的段落：正文，且有一定长度。"""
    return p.get('role') == 'body' and p.get('length', 0) >= 30


if __name__ == '__main__':
    a = '所谓"核心指标"，既包括过程数据，也涵盖结果数据，借此可完整判断效果。'
    b = '所谓"核心指标"，包含过程数据、结果数据，借此可完整判断效果。'
    print('最长重合:', lcs_length(a, b))
    for m in maximal_matches(a, b, 8):
        print('  %-6s %s' % (classify_overlap(b, m), m))
