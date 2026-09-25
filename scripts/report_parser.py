#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""report_parser —— 解析查重 / AIGC 检测报告，输出「哪些句子被命中」。

支持两类输入：

**A. PaperPass 报告包（解压后的目录）**
    报告包结构::

        PaperPass-检测报告/
          ├─ htmls/js/detaildata.js        总相似度 + 来源拆解（期刊/网络/用户库…）
          ├─ htmls/js/texthtmldata.js      逐句相似度 detailJsonData
          ├─ htmls/js/texthtmldata_ai.js   逐句 AIGC 疑似度 aiCheckSentenceList
          ├─ 报告_Word标红版/*查重报告*.docx  逐段着色原文（最权威的定位依据）
          └─ 原文件/*.docx                 送检原稿

**B. 任意「标红版」docx（知网 / 维普 / 万方 / PaperPass 的 Word 报告）**

着色约定（两家略有不同，这里统一识别）::

    w:color   FF0000 / F12828  红 → 重复
              F39800 / FF9900  橙 → 相似
              0070C0 / 0000FF  蓝 → 引用
    w:highlight  red / yellow / green / cyan / magenta

⚠️ 一个真实的坑：PaperPass 标红版用的是 ``w:color``（字体颜色），**不是**
``w:highlight``（高亮底色）。只找 highlight 会得到"整篇全黑"的结论。
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
import sys
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import docx_io                                    # noqa: E402
from docx_io import W, para_text, tag_name        # noqa: E402

# ── 颜色 → 语义 ──────────────────────────────────────────────────────────
COLOR_MAP = {
    'F12828': ('repeat', '红·重复'), 'FF0000': ('repeat', '红·重复'),
    'C00000': ('repeat', '红·重复'), 'FF3300': ('repeat', '红·重复'),
    'F39800': ('similar', '橙·相似'), 'FF9900': ('similar', '橙·相似'),
    'FFC000': ('similar', '橙·相似'), 'E36C0A': ('similar', '橙·相似'),
    '0070C0': ('cite', '蓝·引用'), '0000FF': ('cite', '蓝·引用'),
    '00B0F0': ('cite', '蓝·引用'),
    '000000': ('normal', ''), '333333': ('normal', ''), 'auto': ('normal', ''),
}
HIGHLIGHT_MAP = {
    'red': ('repeat', '红·重复'), 'darkRed': ('repeat', '红·重复'),
    'yellow': ('similar', '黄·相似'), 'green': ('similar', '绿·相似'),
    'cyan': ('cite', '青·引用'), 'magenta': ('repeat', '品红·重复'),
    'blue': ('cite', '蓝·引用'), 'none': ('normal', ''),
}
EDIT_KINDS = ('repeat', 'similar')


# ── A. 报告包里的 JS 数据 ────────────────────────────────────────────────
def _load_js_var(path: str, varname: str):
    """从 js 文件里取 ``var NAME = <json>;`` 并解析。"""
    if not os.path.exists(path):
        return None
    s = open(path, encoding='utf-8', errors='ignore').read()
    m = re.search(r'var\s+%s\s*=\s*(\[.*?\]|\{.*?\});' % re.escape(varname), s, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except Exception:
        return None


def find_file(root: str, *names: str):
    """在目录下按文件名子串查找。``names`` 按**优先级**排列：
    先把所有文件与 names[0] 比，命中不了再退到 names[1]。

    这个"逐轮匹配"不是洁癖——报告目录里同时有
    ``*_查重报告_*.docx`` 与 ``*_AIGC检测报告_*.docx``，
    如果按文件遍历顺序匹配「标红版」，很容易抓到 AIGC 那份。
    """
    allf = []
    for dirpath, _, files in os.walk(root):
        for f in files:
            allf.append(os.path.join(dirpath, f))
    for n in names:
        low = n.lower()
        for p in allf:
            if low in os.path.basename(p).lower():
                return p
    return None


def parse_paperpass(root: str) -> dict:
    """解析 PaperPass 报告目录，返回结构化结果。"""
    out: dict = {'root': root, 'sentences': [], 'ai_sentences': [],
                 'sources': [], 'marked': [], 'summary': {}}

    js = os.path.join(root, 'htmls', 'js', 'detaildata.js')
    if os.path.exists(js):
        s = open(js, encoding='utf-8', errors='ignore').read()
        for k in ('score', 'localScore', 'pScore', 'tScore', 'cScore', 'bScore',
                  'uScore', 'netScore', 'aiScore', 'contentLength',
                  'sentenceCount', 'sectionCount', 'author', 'title'):
            m = re.search(r"var\s+%s\s*=\s*['\"]([^'\"]*)['\"]"
                          r"|var\s+%s\s*=\s*(-?[\d.]+)" % (k, k), s)
            if m:
                out['summary'][k] = (m.group(1) if m.group(1) is not None
                                     else float(m.group(2)))
        m = re.search(r'var\s+similarResources\s*=\s*(\{.*?\});\s*\n', s, re.S)
        if m:
            try:
                r = json.loads(m.group(1))
                out['sources'] = (r.get('local') or []) + (r.get('net') or [])
            except Exception:
                pass

    out['sentences'] = _load_js_var(
        os.path.join(root, 'htmls', 'js', 'texthtmldata.js'), 'detailJsonData') or []
    out['ai_sentences'] = _load_js_var(
        os.path.join(root, 'htmls', 'js', 'texthtmldata_ai.js'),
        'aiCheckSentenceList') or []

    wp = find_file(root, '标红版_查重报告', '查重报告', '标红版')
    if wp and wp.lower().endswith('.docx'):
        out['marked_file'] = wp
        out['marked'] = parse_marked_docx(wp)
    return out


# ── B. 通用标红 docx ─────────────────────────────────────────────────────
def parse_marked_docx(path: str) -> list[dict]:
    """解析标红版 docx，返回逐段着色片段。

    :returns: ``[{'text': 整段文本, 'segments': [(color, text), ...],
                 'marked': [(kind, label, text), ...]}, ...]``（跳过全空段）
    """
    root = ET.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    out = []
    for p in root.find(docx_io.q('body')):
        if p.tag != docx_io.q('p'):
            continue
        segs: list[list] = []          # [color, text]
        for r in p.iter(docx_io.q('r')):
            color = _run_color(r)
            txt = ''.join(x.text or '' for x in r.iter(docx_io.q('t')))
            if not txt:
                continue
            if segs and segs[-1][0] == color:
                segs[-1][1] += txt
            else:
                segs.append([color, txt])
        if not any(s[1].strip() for s in segs):
            continue
        marked = []
        for c, t in segs:
            kind, label = _resolve(c)
            if kind in EDIT_KINDS:
                marked.append({'kind': kind, 'label': label, 'color': c, 'text': t})
        out.append({'text': ''.join(t for _, t in segs),
                    'segments': [(c, t) for c, t in segs],
                    'marked': marked})
    return out


def _run_color(r) -> str:
    """取一个 run 的标记色：优先字体颜色，其次高亮底色，都没有则 000000。"""
    rpr = r.find(docx_io.q('rPr'))
    if rpr is None:
        return '000000'
    c = rpr.find(docx_io.q('color'))
    if c is not None:
        v = (c.get(W + 'val') or '').upper()
        if v and v not in ('AUTO', '000000'):
            return v
    h = rpr.find(docx_io.q('highlight'))
    if h is not None:
        v = h.get(W + 'val')
        if v and v != 'none':
            return 'HL:' + v
    return '000000'


def _resolve(color: str) -> tuple[str, str]:
    if color.startswith('HL:'):
        return HIGHLIGHT_MAP.get(color[3:], ('marked', color[3:]))
    return COLOR_MAP.get(color.upper(), ('marked', color))


# ── 渲染成人类可读文本 ───────────────────────────────────────────────────
def render_paperpass(info: dict, show_paras: bool = True) -> str:
    L = []
    s = info.get('summary') or {}
    L.append('=' * 68)
    L.append('PaperPass 报告解析：%s' % info.get('root'))
    L.append('=' * 68)
    if s:
        L.append('')
        L.append('## 概览')
        L.append('  作者         : %s' % s.get('author', '?'))
        L.append('  送检标题     : %s' % s.get('title', '?'))
        L.append('  正文字数     : %s' % s.get('contentLength', '?'))
        L.append('  总相似度     : %s%%' % s.get('score', '?'))
        L.append('    来源拆解   : 期刊 %s%% / 网络 %s%% / 用户库 %s%% / 学位 %s%% / 其他 %s%%'
                 % (s.get('pScore'), s.get('netScore'), s.get('uScore'),
                    s.get('bScore'), s.get('tScore')))
        L.append('  AIGC 疑似度  : %s%%' % s.get('aiScore', '?'))
        hi = [x for x in info['sentences'] if (x.get('score') or 0) > 0]
        L.append('  命中句子     : %s / %s 句' % (len(hi), s.get('sentenceCount', '?')))
    srcs = info.get('sources') or []
    if srcs:
        L.append('')
        L.append('## Top 相似来源')
        for x in sorted(srcs, key=lambda y: -float(y.get('score') or 0))[:10]:
            L.append('  %5s%%  %s%s' % (x.get('score'), x.get('source', ''),
                                        (x.get('title') or '').strip()))
    S = info.get('sentences') or []
    hi = [x for x in S if (x.get('score') or 0) > 0]
    if hi:
        tot = sum(len(x.get('content') or '') for x in S)
        hit = sum(len(x.get('content') or '') for x in hi)
        L.append('')
        L.append('## 查重逐句命中（%d 句 / %d 字，占正文 %.1f%%）'
                 % (len(hi), hit, hit * 100.0 / max(tot, 1)))
        for x in sorted(hi, key=lambda y: -(y.get('score') or 0)):
            L.append('  [%5.1f] %s' % (x['score'], (x.get('content') or '')[:78]))
    A = info.get('ai_sentences') or []
    if A:
        ah = [x for x in A if (x.get('score') or 0) > 0]
        L.append('')
        L.append('## AIGC 逐句疑似（共 %d 句，疑似 %d 句）' % (len(A), len(ah)))
        for x in ah:
            L.append('  [%5.1f] %s' % (x['score'], (x.get('content') or '')[:78]))
        if not ah:
            L.append('  全部句子评分 0 —— AIGC 侧无需处理')
    mk = info.get('marked') or []
    if mk:
        L.append('')
        L.append('## 逐段着色标记（%s）' % os.path.basename(info.get('marked_file', '')))
        pal, tot = {}, 0
        for para in mk:
            for c, t in para['segments']:
                pal[c] = pal.get(c, 0) + len(t)
                tot += len(t)
        for c, n in sorted(pal.items(), key=lambda x: -x[1]):
            if _resolve(c)[0] != 'normal':
                L.append('  %-10s %-8s %5d 字 %5.1f%%'
                         % (c, _resolve(c)[1], n, n * 100.0 / max(tot, 1)))
        L.append('  未着色 %d 字 %.1f%%'
                 % (pal.get('000000', 0), pal.get('000000', 0) * 100.0 / max(tot, 1)))
        if show_paras:
            L.append('')
            for i, para in enumerate(mk, 1):
                if not para['marked']:
                    L.append('p%03d [未标记] %s' % (i, para['text'][:60]))
                    continue
                buf = []
                for c, t in para['segments']:
                    k, lab = _resolve(c)
                    buf.append(t if k == 'normal' else '【%s|%s】' % (lab, t))
                L.append('p%03d %s' % (i, ''.join(buf)))
    L.append('')
    L.append('提示：橙、红是改写对象；标题、专名全称、引文原文属合规保留，'
             '不要为了压数字改动它们。')
    return '\n'.join(L)


def main():
    import sys
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    tgt = sys.argv[1]
    show = '--paras' in sys.argv
    if os.path.isdir(tgt):
        print(render_paperpass(parse_paperpass(tgt), show_paras=show))
    else:
        mk = parse_marked_docx(tgt)
        hit = sum(len(m['text']) for p in mk for m in p['marked'])
        tot = sum(len(p['text']) for p in mk)
        print('标红版解析：%s' % tgt)
        print('段落 %d，命中 %d 字 / %d 字 = %.1f%%'
              % (len(mk), hit, tot, hit * 100.0 / max(tot, 1)))
        for i, p in enumerate(mk, 1):
            for m in p['marked']:
                print('p%03d [%s] %s' % (i, m['label'], m['text'][:70]))


if __name__ == '__main__':
    main()
