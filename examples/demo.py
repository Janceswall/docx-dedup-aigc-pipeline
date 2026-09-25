#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一键演示：合成一份示例文档 → 跑完整条降重链路 → 产出交付物。

用法::

    python3 examples/demo.py

产出（都在 examples/out/ 下）::

    示例文档-原稿.docx         合成的输入文档
    改写稿-示例文档-原稿.docx   改写结果
    改写对照表-*.docx          逐段对照表（横向）
    自检报告-*.md              结论与验收数据

这份演示**不需要任何真实文档或检测报告**，clone 下来就能跑。
示例内容为通用主题的合成文本，与任何真实材料无关。
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPTS = os.path.join(ROOT, 'scripts')
sys.path.insert(0, SCRIPTS)

import docx_writer                                                  # noqa: E402

DOC = '示例文档-原稿.docx'
WORK = os.path.join(HERE, '_demo_work')
OUT = os.path.join(HERE, 'out')

# ── 示例原文（刻意写入若干"模板化表述"，用来演示重合检测与降重） ──────────
SOURCE = [
    ('h1', '线上协作工具在小团队项目推进中的应用观察'),
    ('p', '摘要：线上协作工具把分散的成员、任务与文档集中到同一处，使小团队在缺少专职管理岗位'
          '的情况下也能维持项目节奏。本文结合三个小团队的落地过程，梳理工具选型与推广中的常见'
          '误区，并提出可操作的改进做法。'),
    ('p', '关键词：协作工具；项目推进；小团队'),
    ('p', ''),
    ('h2', '一、协作工具为什么能提高效率'),
    ('p', '从协作本身的特点来看，信息的流转效率决定了项目推进的速度。讨论、任务与文件分散在'
          '多个渠道时，成员需要反复确认上下文，沟通成本远高于处理成本。线上协作工具的作用正是'
          '把这几条线并到一处，让每个人在同一个界面上看到同一份进度。'),
    ('p', '此外，线上协作工具还有助于减少无效沟通。当任务状态与责任人写在明处时，追问与重复'
          '汇报自然减少，团队可以把精力放在推进本身。因此，在选型时应当优先考虑团队已有习惯，'
          '而不是功能清单的长度。'),
    ('h2', '二、三条可操作的落地做法'),
    ('p', '第一条做法是先把任务拆到可交付的粒度。把"完成第一版方案"拆成交提纲、补数据、出配图'
          '这样的具体动作，责任人与时间点才有地方挂。这种方式把抽象的目标变成了可以逐条勾选的'
          '清单。'),
    ('p', '第二条做法是固定节奏的同步。每周用同一时间、同一份模板过一遍进度，让偏差在早期就被'
          '看见，而不是等到交付前才暴露。实践证明，这种方式对项目按期完成的作用明显。'),
    ('p', ''),
    ('h2', '参考文献'),
    ('p', '[1] 张某. 团队协作与信息流转效率研究[M]. 北京: 某出版社, 2018.'),
    ('p', '[2] 李某. 小团队项目管理实践手册[M]. 上海: 某出版社, 2020.'),
    ('p', '[3] 王某. 协作工具在项目推进中的应用观察[J]. 某期刊, 2019(12): 34-36.'),
    ('p', ''),
]

# ── 对应的改写方案（段号 → 改写文本） ────────────────────────────────────
REWRITES = {
    6: '项目能不能往前推，很大程度上取决于信息流转得顺不顺。讨论在群里、任务在表格里、'
       '文件在网盘里，成员每动一步都得先问清楚上下文从哪来——拖慢进度的其实是这些来回确认，'
       '不是干活本身。协作工具做的事很简单：把这几条线收进同一个界面，让所有人看到的是'
       '同一份进度。',
    7: '无效沟通的减少是另一层收获。状态和责任人一旦摆在明面上，追问和重复汇报就少了，'
       '省下的时间才可能真用在推进上。选工具时值得掂量的其实是团队已经形成的习惯，'
       '而不是功能清单有多长。',
    9: '先说拆解。目标写得再漂亮，落不到具体动作上就没法分配。"第一版方案"是个结果，'
       '不是动作；得把它拆成一条条能勾掉的小事——先列提纲，再找数据，最后配图，'
       '责任人和时间点才有地方挂。抽象的目标变成可勾选的清单，推进这件事才好谈。',
    10: '再说节奏。每周固定一个时间、用同一份模板过一遍进度，听着刻板，但偏差只有在还来得及'
        '调整的时候被看见才有意义。等到交付前才发现问题，能做的就只剩解释了。',
}

TECHNIQUES = {
    6: '句式重构：长句拆短 + 换喻体 + 口语化节奏',
    7: '视角转换：第三人称客观陈述 → 实践者口吻；删除"此外/因此"等连接词',
    9: '结构重排：抽象概括 → 具体场景；加入短句打破等长节奏',
    10: '结构重排 + 口语化；拆掉"实践证明…作用明显"这类模板收尾',
}

CONFIRM = [
    '本文档为合成示例，仅用于演示流程，不含任何真实材料的内容。',
    '实际使用时，请在运行前向用户确认：检测系统、红线数字、是否有标红报告。',
]


def build_source(path: str) -> str:
    b = docx_writer.DocxBuilder(orientation='portrait')
    for kind, text in SOURCE:
        if kind == 'h1':
            b.heading(text, 1)
        elif kind == 'h2':
            b.heading(text, 2)
        else:
            b.paragraph(text, indent_first=2 if text and not text.startswith('[') else 0)
    b.save(path, title='示例文档', author='demo')
    return path


def run(*args) -> int:
    cmd = [sys.executable, os.path.join(SCRIPTS, 'dedup.py')] + list(args)
    line = ' '.join(os.path.basename(a) if a.startswith(SCRIPTS) else a
                    for a in cmd[1:])
    print('$ ' + line, flush=True)
    sys.stdout.flush()
    return subprocess.call(cmd)


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    doc = build_source(os.path.join(HERE, DOC))
    print('① 已生成示例文档：%s' % doc)

    if run('prepare', doc, '-o', WORK) != 0:
        return 1

    sheet_path = os.path.join(WORK, 'sheet.json')
    sheet = json.load(open(sheet_path, encoding='utf-8'))
    print('\n② 示例：自动识别出的段落角色')
    for it in sheet['paragraphs']:
        if it['p'] <= 12:
            print('   p%02d %-10s %s %4d字 | %s'
                  % (it['p'], it['role'], '锁定' if it['locked'] else '可改',
                     it['length'], (it['text'] or '(空)')[:34]))

    n = 0
    for it in sheet['paragraphs']:
        if it['p'] in REWRITES:
            it['rewrite'] = REWRITES[it['p']]
            it['techniques'] = TECHNIQUES.get(it['p'], '')
            n += 1
    sheet['confirm'] = CONFIRM
    json.dump(sheet, open(sheet_path, 'w', encoding='utf-8'),
              ensure_ascii=False, indent=2)
    print('\n③ 已向工作表的 %d 段填入改写文本（其余段一律不动）' % n)

    print('\n④ 跑完整条链路')
    code = run('run', sheet_path, '-o', OUT, '--prefix', '改写稿-' +
               os.path.splitext(DOC)[0])

    print('\n完成，产出目录：%s' % OUT)
    print('（结论为「通过」表示：13 字重合已清零、格式零破坏。'
          '退出码 %d）' % code)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
