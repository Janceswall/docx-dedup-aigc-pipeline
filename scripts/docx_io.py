#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx_io —— 零依赖的 DOCX 读取 / 段落替换 / 保真保存引擎。

设计要点
--------
1. **只改 word/document.xml**：styles / numbering / media / rels / settings /
   fontTable 等其余部件全部按原字节复制。格式零破坏不是靠"改完再修"，而是
   从机制上就不可能破坏。
2. **run 级替换**：新文本写进该段第一个含 ``w:t`` 的 run，继承它的 ``rPr``
   （字体、字号、加粗、颜色），其余 run 只清空文本载体、保留元素壳。
   段落属性 ``pPr`` 与首 run 属性因此 100% 不变。
3. **前缀守卫**：原段以 ``摘要：`` / ``关键词：`` 这类标签开头而新文本漏掉时
   自动补回。这是实战踩过的坑——整段替换极易吃掉标签。
4. **复杂内容拒改**：含图片 / 域 / 文本框 / 超链接 / 数学公式的段落默认拒绝
   替换并抛错，避免把不可重建的东西弄丢。
5. **不依赖段落坐标**：持有 Element 引用，替换顺序无关，不存在"坐标漂移"。
   只有走外部编辑器（腾讯文档本地通道）时才需要从后往前替换。

依赖：Python 3.8+ 标准库。无第三方包。
"""
from __future__ import annotations

import re
import shutil
import zipfile
from xml.etree import ElementTree as ET

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_NS = 'http://www.w3.org/XML/1998/namespace'
W = '{%s}' % W_NS

#: 段落内这些标签承载可见文本
TEXT_TAGS = (W + 't', W + 'tab', W + 'br', W + 'cr', W + 'noBreakHyphen')

#: 出现这些标签的段落一律拒绝替换
COMPLEX_TAGS = ('drawing', 'pict', 'object', 'fldSimple', 'instrText', 'fldChar',
                'txbxContent', 'hyperlink', 'oMath', 'oMathPara', 'ruby')

#: 段落前缀守卫：命中则新文本必须保留同样的前缀
PREFIX_RE = re.compile(
    r'^\s*((?:摘\s*要|关\s*键\s*词|关\s*键\s*字|内容提要|'
    r'Abstract|ABSTRACT|Key\s*words|KEY\s*WORDS)\s*[：:]\s*)')

_NS_REGISTERED = False


# ──────────────────────────────────────────────────────────────────────────
# XML 工具
# ──────────────────────────────────────────────────────────────────────────
def register_namespaces(raw: bytes) -> None:
    """从原文里提取所有 xmlns 声明并注册，避免序列化时冒出 ns0/ns1 前缀。

    ElementTree 不注册的前缀会被重写成 ``ns0:`` 这类形式。虽然 XML 仍然合法、
    Word 也读得进，但会让产出的文件与原文差异巨大、难以做字节级比对，
    也容易被某些严格解析器拒绝。
    """
    global _NS_REGISTERED
    head = raw[:8000].decode('utf-8', 'ignore')
    for prefix, uri in re.findall(r'xmlns:([A-Za-z0-9_.\-]+)\s*=\s*"([^"]+)"', head):
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            pass
    _NS_REGISTERED = True


def q(tag: str) -> str:
    """把 ``p`` 变成 ``{ns}p``。"""
    return W + tag


def para_text(p) -> str:
    """取一个 ``w:p`` 的纯文本（tab → \\t，换行符 → \\n）。"""
    out = []
    for node in p.iter():
        if node.tag == W + 't':
            out.append(node.text or '')
        elif node.tag == W + 'tab':
            out.append('\t')
        elif node.tag in (W + 'br', W + 'cr'):
            out.append('\n')
    return ''.join(out)


def tag_name(el) -> str:
    """返回不带命名空间的标签名。"""
    return el.tag.split('}')[-1] if isinstance(el.tag, str) else ''


# ──────────────────────────────────────────────────────────────────────────
# Document
# ──────────────────────────────────────────────────────────────────────────
class DocxError(Exception):
    """文档结构问题（不可安全替换）。"""


class Document:
    """一个 .docx 的可读写句柄。

    用法::

        doc = Document('论文.docx')
        for i, t in enumerate(doc.block_texts(), 1):
            print(i, t)
        doc.replace_paragraph(5, '改写后的第五段')
        doc.save('论文-改写.docx')
    """

    def __init__(self, path: str):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self._order = list(z.namelist())
            self._parts = {n: z.read(n) for n in self._order}
        if 'word/document.xml' not in self._parts:
            raise DocxError('不是有效的 .docx（缺少 word/document.xml）：%s' % path)
        self._raw = self._parts['word/document.xml']
        register_namespaces(self._raw)
        self.root = ET.fromstring(self._raw)
        self.body = self.root.find(q('body'))
        if self.body is None:
            raise DocxError('document.xml 缺少 w:body：%s' % path)
        self._dirty = False
        self.changed: list[int] = []

    # ── 读取 ─────────────────────────────────────────────────────────────
    def blocks(self):
        """返回 body 顶层块列表 ``[(标签名, Element), ...]``。

        段落号即此列表的下标 + 1。**表格也占一个号**——这与 Word 里
        "第 N 个块"的直觉一致，也与腾讯文档本地通道的段号口径一致。
        """
        return [(tag_name(c), c) for c in self.body
                if c.tag in (q('p'), q('tbl'))]

    def para_elements(self):
        return [c for c in self.body if c.tag == q('p')]

    def block_texts(self) -> list[str]:
        """按块号顺序返回文本；表格位置返回 ``<表格>``。"""
        return [para_text(c) if t == 'p' else '<表格>' for t, c in self.blocks()]

    def paragraphs(self) -> list[dict]:
        """返回结构化段落清单，供生成改写工作表使用。"""
        out = []
        styles = {}
        for i, (tag, el) in enumerate(self.blocks(), start=1):
            if tag != 'p':
                out.append({'p': i, 'kind': 'table', 'text': '<表格>',
                            'style': None, 'numbered': False})
                continue
            ppr = el.find(q('pPr'))
            style = numbered = None
            if ppr is not None:
                ps = ppr.find(q('pStyle'))
                style = ps.get(W + 'val') if ps is not None else None
                numpr = ppr.find(q('numPr'))
                numbered = numpr is not None
            text = para_text(el)
            out.append({'p': i, 'kind': 'p', 'text': text, 'style': style,
                        'numbered': bool(numbered),
                        'runs': [len(para_text(r)) for r in el.findall(q('r'))]})
        return out

    # ── 安全闸 ───────────────────────────────────────────────────────────
    def check_replaceable(self, idx: int, allow: tuple = ()) -> None:
        """确认第 idx 块是普通段落、可以整段替换；否则抛 DocxError。"""
        blocks = self.blocks()
        if not 1 <= idx <= len(blocks):
            raise DocxError('段落号 %d 越界（共 %d 块）' % (idx, len(blocks)))
        tag, el = blocks[idx - 1]
        if tag != 'p':
            raise DocxError('第 %d 块是表格，不能按段落替换' % idx)
        for bad in COMPLEX_TAGS:
            if bad in allow:
                continue
            for node in el.iter():
                if tag_name(node) == bad:
                    raise DocxError(
                        '第 %d 段含 <%s>（图片/域/超链接/公式等无法重建的内容），'
                        '拒绝替换。若确认安全，请用 allow=(%r,)' % (idx, bad, bad))
        return el

    # ── 替换 ─────────────────────────────────────────────────────────────
    def replace_paragraph(self, idx: int, text: str, guard: bool = True,
                          allow: tuple = ()) -> str:
        """把第 idx 段整段文本换成 ``text``，返回真正写入的文本。

        :param guard: 是否启用前缀守卫（保留 摘要：/关键词： 等标签）
        :param allow: 白名单，允许段落中存在哪些「复杂标签」
        """
        el = self.check_replaceable(idx, allow)
        old = para_text(el)

        if guard:
            m = PREFIX_RE.match(old)
            if m and not text.startswith(m.group(1)):
                text = m.group(1) + text.lstrip()

        runs = el.findall(q('r'))
        hosts = [r for r in runs if r.find(q('t')) is not None]
        if not hosts:
            # 段落里没有可写的 run：挑第一个 run（可能是空的）或新建一个
            if runs:
                host = runs[0]
            else:
                host = ET.SubElement(el, q('r'))
            hosts = [host]
            if host.find(q('t')) is None:
                ET.SubElement(host, q('t'))
        host = hosts[0]

        # 1) 清掉 host 里除首个 w:t 之外的所有文本载体
        first_t = host.find(q('t'))
        for node in list(host):
            if node.tag in TEXT_TAGS and node is not first_t:
                host.remove(node)

        # 2) 写文本（含换行 → w:br + w:t）
        lines = text.split('\n')
        first_t.text = lines[0] or ''
        first_t.set('{%s}space' % XML_NS, 'preserve')
        for extra in lines[1:]:
            br = ET.SubElement(host, q('br'))
            br.tail = None
            t = ET.SubElement(host, q('t'))
            t.text = extra or ''
            t.set('{%s}space' % XML_NS, 'preserve')

        # 3) 其余 run 只清空文本载体，保留元素壳（维持 run 数量与属性）
        for r in runs:
            if r is host:
                continue
            for node in list(r):
                if node.tag in TEXT_TAGS:
                    r.remove(node)

        self._dirty = True
        if idx not in self.changed:
            self.changed.append(idx)
        return text

    # ── 保存 ─────────────────────────────────────────────────────────────
    def save(self, out_path: str, keep_meta: bool = True) -> str:
        """另存为新 docx。除 document.xml 外所有部件按原字节复制。"""
        if out_path == self.path:
            raise DocxError('拒绝覆盖源文件，请指定新的输出路径')
        xml = self.document_xml()
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for name in self._order:
                z.writestr(name, xml if name == 'word/document.xml'
                           else self._parts[name])
        return out_path

    def document_xml(self) -> bytes:
        """序列化当前的 document.xml（带 XML 声明）。"""
        body = ET.tostring(self.root, encoding='utf-8')
        return (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + body)


# ──────────────────────────────────────────────────────────────────────────
# 便捷函数
# ──────────────────────────────────────────────────────────────────────────
def read_texts(path: str) -> list[str]:
    """快速取一份 docx 的段落文本列表（表格记为 ``<表格>``）。"""
    return Document(path).block_texts()


#: 开关型元素：缺 val 等价于 val="1"
TOGGLE_TAGS = {'b', 'bCs', 'i', 'iCs', 'strike', 'u', 'caps', 'smallCaps',
               'adjustRightInd', 'autoSpaceDE', 'autoSpaceDN', 'kinsoku',
               'overflowPunct', 'snapToGrid', 'wordWrap', 'widowControl',
               'topLinePunct', 'vanish', 'emboss', 'imprint', 'outline'}

_TRUE_VALS = ('true', 'on', '1')


def canon_element(el):
    """把 XML 元素规范化成可比较的嵌套元组。

    这一步不可省。编辑器 SDK 重写文档时会把 ``<w:b/>`` 序列化成
    ``<w:b w:val="1"/>``、把 ``<w:kinsoku/>`` 写成 ``<w:kinsoku w:val="1"/>``
    ——语义完全等价，但字符串比对会报「所有段落格式都变了」的假警报。
    这里把开关型元素缺 val 一律视为 ``1``，再比结构化元组。
    """
    if el is None:
        return None
    tag = tag_name(el)
    attrs = dict((k.split('}')[-1], v) for k, v in el.attrib.items())
    if tag in TOGGLE_TAGS:
        if not attrs:
            attrs = {'val': '1'}
        elif attrs.get('val', '').lower() in _TRUE_VALS:
            attrs = {'val': '1'}
    kids = tuple(sorted((canon_element(c) for c in el), key=repr))
    return (tag, tuple(sorted(attrs.items())), (el.text or '').strip() or None, kids)


def paragraph_signature(p) -> dict:
    """取一个 ``w:p`` 的「格式指纹」：pPr + 各 run 的 rPr（不含文本）。"""
    return {
        'pPr': canon_element(p.find(q('pPr'))),
        'runs': [canon_element(r.find(q('rPr'))) for r in p.findall(q('r'))],
        'text': para_text(p),
    }


def copy_docx(src: str, dst: str) -> str:
    """复制工作副本——任何改写前都先复制，绝不在原稿上动手。"""
    shutil.copyfile(src, dst)
    return dst


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    d = Document(sys.argv[1])
    for i, t in enumerate(d.block_texts(), 1):
        print('%03d | %s' % (i, t))
