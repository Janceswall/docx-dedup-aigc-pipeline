#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""抽取 docx 全文（段号 + 样式 + 文本），绕开编辑器预览的字数上限。

用法::

    python3 extract_docx_text.py <file.docx> [输出文件]

输出格式::

    001 | style=None | 文档标题
    002 | style=None | 摘要：……
    ...
    段落数(含表格): <N> 正文总字符: <M>

⚠️ 这是**兼容壳**。推荐使用功能更完整的::

    python3 dedup.py prepare <file.docx> -o work/     # 同时识别段落角色、生成改写工作表
    python3 dedup.py roles <file.docx>                # 只看段落角色
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import docx_io                                                     # noqa: E402


def main(path: str, out_path: str | None = None) -> None:
    doc = docx_io.Document(path)
    blocks = doc.blocks()
    lines, total = [], 0
    i = 0
    for tag, el in blocks:
        i += 1
        if tag != 'p':
            lines.append('%03d | <表格> | ' % i)
            continue
        text = docx_io.para_text(el)
        total += len(text)
        ppr = el.find(docx_io.q('pPr'))
        style = num = None
        if ppr is not None:
            ps = ppr.find(docx_io.q('pStyle'))
            style = ps.get(docx_io.W + 'val') if ps is not None else None
            numpr = ppr.find(docx_io.q('numPr'))
            if numpr is not None:
                nid = numpr.find(docx_io.q('numId'))
                ilvl = numpr.find(docx_io.q('ilvl'))
                num = 'numId=%s lvl=%s' % (
                    nid.get(docx_io.W + 'val') if nid is not None else '?',
                    ilvl.get(docx_io.W + 'val') if ilvl is not None else '0')
        mark = ' [自动编号 %s]' % num if num else ''
        lines.append('%03d | style=%s%s | %s' % (i, style, mark, text))
    lines.append('---')
    lines.append('段落数(含表格): %d 正文总字符: %d' % (i, total))
    out = '\n'.join(lines)
    if out_path:
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(out)
    else:
        print(out)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
