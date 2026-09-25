#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""解析 PaperPass 检测报告包（或任意标红版 docx）。

用法::

    python3 parse_paperpass_report.py "<报告文件夹>"            # 概览 + 逐段标记
    python3 parse_paperpass_report.py "<报告文件夹>" --no-paras  # 只要概览
    python3 parse_paperpass_report.py "<标红版.docx>"            # 直接解析 Word 标红版

报告包结构（PaperPass 下载解压后）::

    PaperPass-检测报告/
      ├─ htmls/js/detaildata.js         → 总相似度、来源拆解、Top 相似来源
      ├─ htmls/js/texthtmldata.js       → 逐句相似度（detailJsonData）
      ├─ htmls/js/texthtmldata_ai.js    → 逐句 AIGC 疑似度
      ├─ 报告_Word标红版/*查重报告*.docx  → 逐段着色原文（最权威的定位依据）
      └─ 原文件/*.docx                   → 送检原稿

Word 标红版的着色约定::

    w:color  F39800 橙=相似   F12828/FF0000 红=重复   0070C0 蓝=引用
    ⚠️ PaperPass 用字体颜色（w:color）着色，**不是**高亮底色（w:highlight）。
       只找 highlight 会得到「整篇全黑」的错误结论。

兼容壳；推荐直接使用 ``python3 dedup.py report <目标>``
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import dedup                                                       # noqa: E402


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    dedup.do_report(sys.argv[1], as_json=False,
                    show_paras='--no-paras' not in sys.argv)


if __name__ == '__main__':
    main()
