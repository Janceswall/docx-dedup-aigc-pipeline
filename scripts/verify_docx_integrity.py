#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""改写前后文档的结构与格式保真校验。

用法::

    python3 verify_docx_integrity.py <原稿.docx> <改写稿.docx> [--changed 2,4,5,...]

检查项::

    1. 段落/块数量是否一致
    2. 哪些段落被改动、哪些保持原样
    3. 段落属性 pPr 是否零改动
    4. 未改写段落的 run 属性是否变化
    5. 空段数量、参考文献条目是否保留
    6. 总字数变化

⚠️ pPr / rPr 比对**必须做语义归一化**：编辑器会把 ``<w:b/>`` 写成
``<w:b w:val="1"/>``，直接字符串比对会报「全部段落格式都变了」的假警报。
本脚本通过 :func:`docx_io.canon_element` 处理了这一点。

兼容壳；推荐直接使用 ``python3 dedup.py verify ...``
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import dedup                                                       # noqa: E402


def main():
    argv = list(sys.argv[1:])
    exp = None
    if '--changed' in argv:
        i = argv.index('--changed')
        exp = [int(x) for x in argv[i + 1].split(',')]
        argv = argv[:i] + argv[i + 2:]
    if len(argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    res = dedup.do_verify(argv[0], argv[1], exp)
    dedup._print_verify(res)
    raise SystemExit(0 if res['ok'] else 2)


if __name__ == '__main__':
    main()
