#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验「编辑器通道」返回的字符坐标与文档真实文本是否逐段对齐。

⚠️ **只有走编辑器通道（腾讯文档本地通道 / editor_sdk）时才需要这个脚本。**
本仓库默认的零依赖通道不需要坐标——它直接持有段落元素引用，
替换顺序完全无关，不存在坐标漂移问题。

用法::

    # 1) 先用编辑器 SDK 拿结构，落成 JSON
    #    <edsdk> call doc_resolve_document_structure \
    #      --json '{"file_id":"...","mode":"compact","limit":0}' > structure.json
    # 2) 校验
    python3 verify_align.py <file.docx> <structure.json> [ranges.json]

对齐关系::

    span = end_index - start_index = len(text) + 1

``end_index`` 是**闭端**，末尾含段落结束符。因此可替换范围是
``[start_index, start_index + len(text))``。

任何一段对不上就**停下来排查**（可能混入了表格 / 图片 / 域节点），
不要凭感觉往下替换。
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import docx_io                                                     # noqa: E402


def main(docx: str, struct: str, out_ranges: str | None = None) -> int:
    texts = docx_io.read_texts(docx)
    nodes = json.load(open(struct, encoding='utf-8'))['nodes']
    print('docx 段落数 = %d | 编辑器节点数 = %d' % (len(texts), len(nodes)))
    bad = 0
    for i, (t, n) in enumerate(zip(texts, nodes), start=1):
        span = n['end_index'] - n['start_index']
        if span != len(t) + 1:
            bad += 1
            print('!! p%02d span=%d textlen=%d start=%d end=%d | %s'
                  % (i, span, len(t), n['start_index'], n['end_index'], t[:30]))
    print('对齐异常段落数 =', bad)
    if bad:
        print('=> 坐标不可信，先排查（可能混入表格/图片/字段节点），不要直接替换')
        return 2
    rows = [{'p': i, 'begin': n['start_index'], 'end': n['start_index'] + len(t),
             'len': len(t)}
            for i, (t, n) in enumerate(zip(texts, nodes), start=1) if t.strip()]
    if out_ranges:
        with open(out_ranges, 'w', encoding='utf-8') as f:
            json.dump(rows, f, ensure_ascii=False, indent=1)
        print('已写出 %s，可替换段落数 = %d' % (out_ranges, len(rows)))
    return 0


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(__doc__)
        raise SystemExit(1)
    raise SystemExit(main(sys.argv[1], sys.argv[2],
                          sys.argv[3] if len(sys.argv) > 3 else None))
