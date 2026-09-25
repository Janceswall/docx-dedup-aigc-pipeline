#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""13 字重合检测：找出改写稿与原文之间所有 ≥ 阈值的连续重合片段。

知网连续 13 字相同即标红，所以这个检测器是「改到位没有」的核心判据。

用法::

    python3 check_overlap.py <原稿.docx> <改写稿.docx> [阈值，默认13] [--only 2,4,5] [--skip 1,3] [--all]

判读::

    [引文内-合规]  落在中文引号内            → 引文原文，保留
    [专名-须保留]  含书名号                   → 书名/专名全称，无法也不该替换
    !! 需处理      其他长重合                 → 必须回去做深度重构

默认**跳过两稿一致的段落**（标题、关键词、参考文献表、本轮没动的正文段）——
它们本来就与原文逐字相同，报出来只是噪音。要全量显示加 ``--all``。

⚠️ 第二轮（按检测报告精修）时，本检测只作参考——正确的验收口径是
「报告标记的句子是否被打断」，用 ``dedup.py residue``。

兼容壳；推荐直接使用 ``python3 dedup.py check ...``
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import dedup                                                       # noqa: E402


def main():
    argv = list(sys.argv[1:])
    only = skip = None
    include_unchanged = '--all' in argv
    if include_unchanged:
        argv.remove('--all')
    for flag in ('--only', '--skip'):
        if flag in argv:
            i = argv.index(flag)
            vals = [int(x) for x in argv[i + 1].split(',')]
            if flag == '--only':
                only = vals
            else:
                skip = vals
            argv = argv[:i] + argv[i + 2:]          # 摘掉选项与其取值
    if len(argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    k = int(argv[2]) if len(argv) > 2 and argv[2].isdigit() else 13
    res = dedup.do_check(argv[0], argv[1], k, only, skip, include_unchanged)
    dedup._print_check(res)
    raise SystemExit(2 if res['risky'] else 0)


if __name__ == '__main__':
    main()
