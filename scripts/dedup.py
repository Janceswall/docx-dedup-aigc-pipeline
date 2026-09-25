#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dedup —— 端到端降重 + 降 AI 命令行入口。

    # ① 准备：抽取全文 + 生成改写工作表（零依赖通道，不需要任何编辑器）
    python3 dedup.py prepare 论文.docx -o work/

    # ② 在 work/sheet.json 里给要改的段落填 rewrite 字段（人或 AI 来填）

    # ③ 一条命令跑完全部：应用改写 → 13 字检测 → 格式保真校验 → 生成对照表
    python3 dedup.py run work/sheet.json -o out/

    # 拿到检测报告后的第二轮（按标红位置精修）
    python3 dedup.py report "PaperPass-检测报告/" --paras
    python3 dedup.py run work/sheet_r2.json -o out2/ --report "PaperPass-检测报告/"

子命令一览
----------
prepare   抽取文本、识别段落角色、生成 sheet.json 工作表
apply     把 sheet.json 里的 rewrite 落回 docx
check     13 字连续重合检测（改写稿 vs 原稿）
residue   报告驱动验收：检测报告标记的句子是否还残留在新稿里
verify    结构与格式保真校验（段落数 / pPr / rPr / 参考文献 / 空段）
compare   生成横向逐段对照表 docx
report    解析 PaperPass 报告包或任意标红版 docx
run       端到端：apply → check → verify → compare（可带 residue）
roles     列出段落角色，辅助决定改哪些段

退出码：0 通过；1 参数/文件错误；2 检查未通过（仍有需处理项）。
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import docx_io                                                     # noqa: E402
import docx_writer                                                 # noqa: E402
import report_parser                                               # noqa: E402
import textutil                                                    # noqa: E402

try:                                                               # Windows 控制台
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

VERSION = '1.3.0'

#: 残留项的处置状态（用于自检报告）
STATUS_CN = {'risk': '需继续处理', 'retained': '合规保留', 'broken': '已打断'}


# ──────────────────────────────────────────────────────────────────────────
# 工具
# ──────────────────────────────────────────────────────────────────────────
def _now() -> str:
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _read_json(path: str) -> dict:
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def _write_json(path: str, obj) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(path)) or '.', exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def _stem(path: str) -> str:
    return os.path.splitext(os.path.basename(path))[0]


# ──────────────────────────────────────────────────────────────────────────
# 核心步骤
# ──────────────────────────────────────────────────────────────────────────
def do_prepare(src: str, workdir: str, threshold: int = 13) -> dict:
    """抽取全文、识别角色、生成改写工作表。"""
    doc = docx_io.Document(src)
    paras = textutil.annotate(doc.paragraphs())

    # 顺手做一次「可替换性体检」——有段落含图片/域就先告诉用户
    blocked = []
    for p in paras:
        if p['kind'] != 'p':
            continue
        try:
            doc.check_replaceable(p['p'])
        except docx_io.DocxError as e:
            blocked.append({'p': p['p'], 'reason': str(e)})

    sheet = {
        'schema': 'docx-dedup-aigc-pipeline/sheet@1',
        'generated_at': _now(),
        'source': os.path.abspath(src),
        'threshold': threshold,
        'blocked': blocked,
        'paragraphs': [
            {'p': p['p'], 'role': p['role'], 'locked': p['locked'],
             'length': p['length'], 'text': p['text'],
             'rewrite': '', 'techniques': '', 'note': ''}
            for p in paras
        ],
        'confirm': [],          # 需要用户拍板的事项，交付时原样列出
    }
    os.makedirs(workdir, exist_ok=True)
    _write_json(os.path.join(workdir, 'sheet.json'), sheet)

    lines = ['%03d | %-10s %s | %s' % (p['p'], textutil.ROLE_LABEL.get(p['role'], p['role']),
                                       '【锁定】' if p['locked'] else '        ', p['text'])
             for p in paras]
    open(os.path.join(workdir, 'full_text.txt'), 'w', encoding='utf-8').write(
        '\n'.join(lines) + '\n')
    return sheet


def do_apply(sheet: dict, out_path: str, base: str | None = None) -> dict:
    """把 sheet 里的 rewrite 落回文档。未填 rewrite 的段落一律不动。"""
    src = base or sheet.get('source')
    if not src or not os.path.exists(src):
        raise SystemExit('找不到源文档：%s' % src)
    doc = docx_io.Document(src)
    applied, skipped = [], []
    for item in sheet.get('paragraphs', []):
        new_text = (item.get('rewrite') or '').strip()
        if not new_text:
            continue
        p = item['p']
        final = doc.replace_paragraph(p, item['rewrite'])
        applied.append({'p': p, 'role': item.get('role'),
                        'before': item.get('text', ''),
                        'after': final,
                        'techniques': item.get('techniques', '')})
        skipped.append(p)
    doc.save(out_path)
    return {'output': out_path, 'applied': applied,
            'applied_count': len(applied), 'unchanged_count':
                len(doc.block_texts()) - len(applied)}


def do_check(orig: str, new: str, k: int = 13, only=None, skip=None,
             include_unchanged: bool = False) -> dict:
    """13 字连续重合检测。

    默认**跳过两稿间完全一致的段落**——它们本来就与原文逐字相同（标题、关键词、
    参考文献表、以及本轮没动的正文段），报出来只是噪音。要看全部请传
    ``include_unchanged=True``。
    """
    A = docx_io.read_texts(orig)
    B = docx_io.read_texts(new)
    only, skip = set(only or []), set(skip or [])
    findings, unchanged = [], 0
    for i, (a, b) in enumerate(zip(A, B), start=1):
        if (only and i not in only) or i in skip:
            continue
        if not include_unchanged and a == b:
            unchanged += 1
            continue
        for m in textutil.maximal_matches(a, b, k):
            if len(m) < k:
                continue
            kind = textutil.classify_overlap(b, m)
            findings.append({'p': i, 'kind': kind, 'len': len(m), 'text': m})
    risky = [f for f in findings if f['kind'] == 'risk']
    return {'orig': orig, 'new': new, 'k': k,
            'paragraph_count': [len(A), len(B)],
            'findings': findings, 'total': len(findings), 'risky': len(risky),
            'skipped_unchanged': unchanged}


def do_residue(report_target: str, new: str, k: int = 13) -> dict:
    """报告驱动验收：检测报告标记过的句子，是否还残留在新稿里。

    这是第二轮的**正确验收口径**。不要用「新稿和上一版有多少重合」来判断——
    上一版刻意保留的未标记文字本来就该重合。
    """
    if os.path.isdir(report_target):
        info = report_parser.parse_paperpass(report_target)
        segs = []
        for para in info.get('marked') or []:
            for m in para['marked']:
                if len(m['text']) >= k:
                    segs.append(m)
    else:
        segs = []
        for para in report_parser.parse_marked_docx(report_target):
            for m in para['marked']:
                if len(m['text']) >= k:
                    segs.append(m)
    doc = docx_io.Document(new)
    texts = doc.block_texts()
    roles = [p['role'] for p in textutil.annotate(doc.paragraphs())]
    parts, spans, pos = [], [], 0
    for t, role in zip(texts, roles):
        spans.append((pos, pos + len(t), role))
        parts.append(t)
        pos += len(t)
    full = ''.join(parts)

    def _compliant(x: str) -> bool:
        """残留片段是否属「合规保留」。

        两条判据：① 片段本身是引文/书名（引号或书名号包裹）；
        ② 片段落在标题、关键词、参考文献这类**本来就不该改**的段落里。
        """
        if textutil.classify_overlap(full, x) != 'risk':
            return True
        at = full.find(x)
        if at >= 0:
            for a, b, role in spans:
                if a <= at < b:
                    return role in textutil.LOCKED_ROLES
        return False

    items = []
    for m in segs:
        res = sorted((x for x in textutil.maximal_matches(full, m['text'], k)
                      if len(x) >= k), key=len, reverse=True)
        # 残留片段本身也要分类：属合规保留的不计入「需继续处理」——
        # 否则交付结论永远是"仍有待处理项"。
        risky = [x for x in res if not _compliant(x)]
        items.append({'label': m['label'], 'source_len': len(m['text']),
                      'max_residue': max((len(x) for x in res), default=0),
                      'residue': res[:3], 'risky': risky,
                      'status': ('risk' if risky
                                 else ('retained' if res else 'broken')),
                      'text': m['text']})
    bad = [x for x in items if x['status'] == 'risk']
    return {'target': report_target, 'new': new, 'k': k,
            'marked_count': len(items),
            'residue_count': sum(1 for x in items if x['residue']),
            'risky_count': len(bad), 'items': items}


def do_verify(orig: str, new: str, expected_changed=None) -> dict:
    """结构与格式保真校验。"""
    A, B = docx_io.Document(orig), docx_io.Document(new)
    sa = [docx_io.paragraph_signature(c) for c in A.para_elements()]
    sb = [docx_io.paragraph_signature(c) for c in B.para_elements()]
    ta, tb = A.block_texts(), B.block_texts()

    changed = [i for i, (a, b) in enumerate(zip(ta, tb), 1) if a != b]
    same = [i for i, (a, b) in enumerate(zip(ta, tb), 1) if a == b]
    exp_bad = {}
    if expected_changed:
        exp, got = set(expected_changed), set(changed)
        exp_bad = {'not_changed': sorted(exp - got), 'over_changed': sorted(got - exp)}

    ppr_bad = [i for i, (a, b) in enumerate(zip(sa, sb), 1) if a['pPr'] != b['pPr']]
    same_run_bad = [i for i in same
                    if i <= len(sa) and sa[i - 1]['runs'] != sb[i - 1]['runs']]

    def refs(ts):
        return [t for t in ts if textutil.CITATION_MARK.search(t)]

    ra, rb = refs(ta), refs(tb)
    return {
        'paragraph_count': [len(ta), len(tb)],
        'changed': changed, 'unchanged': same,
        'expected_mismatch': exp_bad,
        'ppr_changed': ppr_bad,
        'run_props_changed_in_unchanged': same_run_bad,
        'empty_paragraphs': [sum(1 for t in ta if not t.strip()),
                             sum(1 for t in tb if not t.strip())],
        'references': [len(ra), len(rb)], 'references_identical': ra == rb,
        'chars': [sum(map(len, ta)), sum(map(len, tb))],
        'ok': (len(ta) == len(tb) and not ppr_bad and not same_run_bad
               and ra == rb),
    }


def do_compare(orig: str, new: str, out_path: str, sheet: dict | None = None,
               title: str = '降重降 AI 改写对照表') -> str:
    """生成横向逐段对照表 docx。"""
    A, B = docx_io.Document(orig), docx_io.Document(new)
    ta, tb = A.block_texts(), B.block_texts()
    role_of, note_of = {}, {}
    if sheet:
        for it in sheet.get('paragraphs', []):
            role_of[it['p']] = it.get('role')
            note_of[it['p']] = (it.get('techniques') or it.get('note') or '')
    else:
        for it in textutil.annotate(A.paragraphs()):
            role_of[it['p']] = it['role']

    changed = [i for i, (a, b) in enumerate(zip(ta, tb), 1) if a != b]
    doc = docx_writer.DocxBuilder(orientation='landscape')

    doc.heading(title, 1)
    doc.paragraph('原稿：%s' % os.path.basename(orig), size=9, color='595959')
    doc.paragraph('改写稿：%s' % os.path.basename(new), size=9, color='595959')
    doc.paragraph('生成时间：%s　　工具：docx-dedup-aigc-pipeline v%s'
                  % (_now(), VERSION), size=9, color='595959')
    doc.spacer()

    doc.heading('一、总体情况', 2)
    doc.table(['项目', '数值'], [
        ['段落/块总数', '%d（原稿） / %d（改写稿）' % (len(ta), len(tb))],
        ['实际改写段落', '%d 段' % len(changed)],
        ['保持原样段落', '%d 段' % (len(ta) - len(changed))],
        ['总字数', '%d → %d' % (sum(map(len, ta)), sum(map(len, tb)))],
        ['逐段对照覆盖', '下文表格仅列实际改动的段落'],
    ], widths=[0.22, 0.78], size=9)

    doc.heading('二、逐段对照', 2)
    if changed:
        rows = []
        for i in changed:
            rows.append(['p%02d' % i,
                         textutil.ROLE_LABEL.get(role_of.get(i, ''), role_of.get(i, '')),
                         ta[i - 1], tb[i - 1], note_of.get(i, '')])
        doc.table(['段号', '角色', '原文', '改写后', '手法/说明'], rows,
                  widths=[0.045, 0.075, 0.385, 0.385, 0.11], size=8.5, zebra='F2F6FA')
    else:
        doc.paragraph('（无改动段落）', size=9)

    doc.heading('三、未改动部分清单', 2)
    untouched = [i for i in range(1, len(ta) + 1) if i not in changed]
    doc.paragraph('以下 %d 段逐字未动：' % len(untouched), size=9)
    for i in untouched:
        doc.paragraph('p%02d ［%s］%s' % (i, textutil.ROLE_LABEL.get(
            role_of.get(i, ''), role_of.get(i, '')), (ta[i - 1] or '(空段)')[:110]),
            size=8.5, color='595959')

    if sheet and sheet.get('confirm'):
        doc.heading('四、需要确认的事项', 2)
        for j, c in enumerate(sheet['confirm'], 1):
            doc.paragraph('%d. %s' % (j, c), size=9)
    doc.save(out_path, title=title)
    return out_path


def do_report(target: str, as_json: bool = False, show_paras: bool = True):
    if os.path.isdir(target):
        info = report_parser.parse_paperpass(target)
        if as_json:
            data = dict(info)
            data['marked'] = [{**p, 'segments': [list(s) for s in p['segments']]}
                              for p in (info.get('marked') or [])]
            print(json.dumps(data, ensure_ascii=False, indent=2))
        else:
            print(report_parser.render_paperpass(info, show_paras=show_paras))
        return info
    mk = report_parser.parse_marked_docx(target)
    if as_json:
        print(json.dumps(mk, ensure_ascii=False, indent=2))
        return mk
    hit = sum(len(m['text']) for p in mk for m in p['marked'])
    tot = sum(len(p['text']) for p in mk)
    print('标红版解析：%s' % target)
    print('段落 %d，命中 %d / %d 字 = %.1f%%'
          % (len(mk), hit, tot, hit * 100.0 / max(tot, 1)))
    for i, p in enumerate(mk, 1):
        for m in p['marked']:
            print('p%03d [%s] %s' % (i, m['label'], m['text'][:70]))
    return mk


# ──────────────────────────────────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────────────────────────────────
def _print_check(res: dict) -> None:
    na, nb = res['paragraph_count']
    if na != nb:
        print('!! 段落数不一致（原稿 %d / 改写稿 %d）——先查结构' % (na, nb))
    for f in sorted(res['findings'], key=lambda x: (x['p'], -x['len'])):
        print('p%02d %s 长度%3d: %s' % (f['p'], textutil.OVERLAP_LABEL[f['kind']],
                                        f['len'], f['text']))
    print('---')
    print('>=%d 字重合片段 %d 处，其中需继续处理 %d 处'
          % (res['k'], res['total'], res['risky']))
    if res.get('skipped_unchanged'):
        print('（已跳过 %d 段两稿一致的段落；它们是刻意未改的，加 --all 可一并显示）'
              % res['skipped_unchanged'])
    if res['risky']:
        print('=> 对这些片段做深度重构（拆句重组，不是换同义词）')
    else:
        print('=> 通过：剩余重合均为引文/专名等合规保留项')


def _print_verify(res: dict) -> None:
    print('段落数: 原稿 %d / 改写稿 %d -> %s'
          % (*res['paragraph_count'], '一致' if res['paragraph_count'][0]
             == res['paragraph_count'][1] else '!! 不一致'))
    print('发生改动 %d 段: %s' % (len(res['changed']), res['changed']))
    print('保持原样 %d 段: %s' % (len(res['unchanged']), res['unchanged']))
    if res['expected_mismatch']:
        em = res['expected_mismatch']
        if em.get('not_changed'):
            print('!! 预期要改但没改:', em['not_changed'])
        if em.get('over_changed'):
            print('!! 不该改却改了:', em['over_changed'])
    print('段落属性 pPr 被改动:', res['ppr_changed'] or '无')
    print('未改写段落中 run 属性变化:', res['run_props_changed_in_unchanged'] or '无')
    print('空段: %d / %d   参考文献: %d / %d 条 -> %s'
          % (*res['empty_paragraphs'], *res['references'],
             '完全一致' if res['references_identical'] else '!! 不一致'))
    print('字数: %d -> %d' % tuple(res['chars']))
    print('=> %s' % ('通过：结构与格式零破坏' if res['ok'] else '!! 存在问题，见上'))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog='dedup', description='端到端 DOCX 降重 + 降 AIGC 流水线',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__.split('子命令一览')[1] if '子命令一览' in __doc__ else '')
    ap.add_argument('--version', action='version', version=VERSION)
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('prepare', help='抽取全文并生成改写工作表')
    p.add_argument('src'); p.add_argument('-o', '--out', default='dedup_work')
    p.add_argument('-k', '--threshold', type=int, default=13)

    p = sub.add_parser('apply', help='把工作表里的改写落回 docx')
    p.add_argument('sheet'); p.add_argument('-o', '--out', required=True)
    p.add_argument('--base', help='覆盖工作表里记录的源文档')

    p = sub.add_parser('check', help='13 字连续重合检测')
    p.add_argument('orig'); p.add_argument('new')
    p.add_argument('-k', type=int, default=13)
    p.add_argument('--only', help='只检查这些段，如 2,4,5')
    p.add_argument('--skip', help='跳过这些段，如 1,3,6')
    p.add_argument('--all', action='store_true',
                   help='连两稿一致的段落也一并显示（默认跳过）')
    p.add_argument('--json', action='store_true')

    p = sub.add_parser('residue', help='报告驱动验收：标记句是否还残留')
    p.add_argument('report'); p.add_argument('new')
    p.add_argument('-k', type=int, default=13)
    p.add_argument('--json', action='store_true')

    p = sub.add_parser('verify', help='结构与格式保真校验')
    p.add_argument('orig'); p.add_argument('new')
    p.add_argument('--changed', help='预期改动的段号，如 2,4,5')
    p.add_argument('--json', action='store_true')

    p = sub.add_parser('compare', help='生成横向逐段对照表 docx')
    p.add_argument('orig'); p.add_argument('new')
    p.add_argument('-o', '--out', required=True)
    p.add_argument('--sheet'); p.add_argument('--title', default='降重降 AI 改写对照表')

    p = sub.add_parser('report', help='解析检测报告')
    p.add_argument('target'); p.add_argument('--json', action='store_true')
    p.add_argument('--no-paras', action='store_true')

    p = sub.add_parser('roles', help='列出段落角色')
    p.add_argument('src')

    p = sub.add_parser('run', help='端到端：apply + check + verify + compare')
    p.add_argument('sheet'); p.add_argument('-o', '--out', default='dedup_out')
    p.add_argument('-k', type=int, default=13)
    p.add_argument('--report', help='检测报告目录或标红版 docx，附带做残留验收')
    p.add_argument('--prefix', default='')

    a = ap.parse_args(argv)

    if a.cmd == 'prepare':
        sheet = do_prepare(a.src, a.out, a.threshold)
        n = len(sheet['paragraphs'])
        todo = [x for x in sheet['paragraphs'] if not x['locked'] and x['length'] >= 30]
        print('已生成 %s/sheet.json（%d 块，建议改写 %d 段）' % (a.out, n, len(todo)))
        print('已生成 %s/full_text.txt' % a.out)
        if sheet['blocked']:
            print('!! 有 %d 段不可安全替换（含图片/域/超链接）：'
                  % len(sheet['blocked']))
            for b in sheet['blocked']:
                print('   p%02d %s' % (b['p'], b['reason'][:70]))
        print('下一步：编辑 sheet.json 的 rewrite 字段，然后运行 '
              'dedup.py run %s/sheet.json -o out/' % a.out)
        return 0

    if a.cmd == 'roles':
        for it in textutil.annotate(docx_io.Document(a.src).paragraphs()):
            print('p%02d %-10s %s %5d字 | %s'
                  % (it['p'], textutil.ROLE_LABEL.get(it['role'], it['role']),
                     '锁定' if it['locked'] else '可改', it['length'],
                     (it['text'] or '(空)')[:56]))
        return 0

    if a.cmd == 'apply':
        sheet = _read_json(a.sheet)
        res = do_apply(sheet, a.out, a.base)
        print('已写出 %s（改写 %d 段，未动 %d 段）'
              % (res['output'], res['applied_count'], res['unchanged_count']))
        _write_json(os.path.splitext(a.out)[0] + '.applied.json', res)
        return 0

    if a.cmd == 'check':
        only = [int(x) for x in a.only.split(',')] if a.only else None
        skip = [int(x) for x in a.skip.split(',')] if a.skip else None
        res = do_check(a.orig, a.new, a.k, only, skip,
                       include_unchanged=a.all)
        if a.json:
            print(json.dumps(res, ensure_ascii=False, indent=2))
        else:
            _print_check(res)
        return 2 if res['risky'] else 0

    if a.cmd == 'residue':
        res = do_residue(a.report, a.new, a.k)
        if a.json:
            print(json.dumps(res, ensure_ascii=False, indent=2))
        else:
            for it in res['items']:
                flag = {'risk': '!! 残留待处理', 'retained': ' = 残留(合规保留)',
                        'broken': 'OK 已打断'}[it['status']]
                print('%s 标记%3d字 最长残留%3d  [%s] %s'
                      % (flag, it['source_len'], it['max_residue'], it['label'],
                         it['text'][:56]))
            print('---')
            print('标记片段 %d 处：残留 %d 处，其中需继续处理 %d 处'
                  % (res['marked_count'], res['residue_count'],
                     res['risky_count']))
        return 2 if res['risky_count'] else 0

    if a.cmd == 'verify':
        ch = [int(x) for x in a.changed.split(',')] if a.changed else None
        res = do_verify(a.orig, a.new, ch)
        if a.json:
            print(json.dumps(res, ensure_ascii=False, indent=2))
        else:
            _print_verify(res)
        return 0 if res['ok'] else 2

    if a.cmd == 'compare':
        sheet = _read_json(a.sheet) if a.sheet else None
        out = do_compare(a.orig, a.new, a.out, sheet, a.title)
        print('已生成对照表：%s' % out)
        return 0

    if a.cmd == 'report':
        do_report(a.target, a.json, not a.no_paras)
        return 0

    if a.cmd == 'run':
        sheet = _read_json(a.sheet)
        os.makedirs(a.out, exist_ok=True)
        src = sheet['source']
        stem = a.prefix or ('改写稿-' + _stem(src))
        new_docx = os.path.join(a.out, stem + '.docx')
        cmp_docx = os.path.join(a.out, '改写对照表-' + stem + '.docx')
        log_md = os.path.join(a.out, '自检报告-' + stem + '.md')

        print('① 应用改写 …')
        ap_res = do_apply(sheet, new_docx)
        print('   改写 %d 段，未动 %d 段' % (ap_res['applied_count'],
                                            ap_res['unchanged_count']))

        role_of = {it['p']: it.get('role') for it in sheet.get('paragraphs', [])}
        changed = [x['p'] for x in ap_res['applied']]
        # 只把「正文段」纳入重合检测：标题、关键词这类即便改过也必然与原文
        # 大面积重合（题目本来就那几个字），纳入只会产生假警报。
        body_changed = [p for p in changed if role_of.get(p) == 'body'] or changed
        print('② 13 字重合检测 …')
        ck = do_check(src, new_docx, a.k, only=body_changed)
        print('   重合片段 %d 处，需继续处理 %d 处（口径：与原稿的重合%s）'
              % (ck['total'], ck['risky'],
                 '；本轮请以 ④ 的残留验收为准' if a.report else ''))

        print('③ 格式保真校验 …')
        vf = do_verify(src, new_docx, changed)
        print('   %s' % ('结构与格式零破坏' if vf['ok'] else '!! 存在问题'))

        rs = None
        if a.report:
            print('④ 报告驱动残留验收 …')
            rs = do_residue(a.report, new_docx, a.k)
            print('   标记 %d 处：残留 %d 处，其中需继续处理 %d 处'
                  % (rs['marked_count'], rs['residue_count'],
                     rs['risky_count']))

        print('⑤ 生成对照表 …')
        do_compare(src, new_docx, cmp_docx, sheet)

        # 通过与否的口径按轮次区分：有报告时以「标记句是否被打断」为准，
        # 无报告时以「与原稿是否还有长重合」为准。
        ok = bool(vf['ok']
                  and (not rs['risky_count'] if rs else not ck['risky']))
        print('⑥ 结论：%s' % ('通过' if ok else '仍有待处理项'))

        md = ['# 自检报告', '',
              '## 结论：**%s**' % ('通过' if ok else '仍有待处理项'), '',
              '- 主判据：%s' % ('报告标记句是否被打断（第二节仅作参考）' if rs
                               else '与原稿的 13 字连续重合'),
              '- 源文档：`%s`' % os.path.basename(src),
              '- 改写稿：`%s`' % os.path.basename(new_docx),
              '- 生成时间：%s' % _now(), '',
              '## 一、改写概览', '',
              '| 项目 | 数值 |', '|---|---|',
              '| 改写段落 | %d 段 |' % ap_res['applied_count'],
              '| 未改动段落 | %d 段 |' % ap_res['unchanged_count'],
              '| 字数变化 | %d → %d |' % tuple(vf['chars']),
              '| 段落数 | %d → %d |' % tuple(vf['paragraph_count']), '']
        md += ['## 二、重合检测（阈值 %d 字）' % a.k, '']
        if rs:
            md += ['> ⚠️ 本轮为**报告驱动精修**。此处的「与原稿重合」只作参考——'
                   '上一轮刻意保留的未标红文字本来就该重合。\n'
                   '> **主判据见第四节「报告驱动残留验收」**。', '']
        else:
            md += ['> 本轮为**预降重**（尚无检测报告）。通过与否以此节为准。', '']
        md += ['| 分类 | 数量 |', '|---|---|']
        kinds = {}
        for f in ck['findings']:
            kinds[f['kind']] = kinds.get(f['kind'], 0) + 1
        md.append('| 引文内-合规 | %d |' % kinds.get('quote', 0))
        md.append('| 专名-须保留 | %d |' % kinds.get('title', 0))
        md.append('| **需继续处理** | **%d** |' % kinds.get('risk', 0))
        md.append('')
        if kinds.get('risk'):
            md.append('需继续处理的片段：')
            for f in ck['findings']:
                if f['kind'] == 'risk':
                    md.append('- p%02d（%d 字）%s' % (f['p'], f['len'], f['text']))
            md.append('')
        md += ['## 三、格式保真', '',
               '| 检查项 | 结果 |', '|---|---|',
               '| 段落数 | %s |' % ('一致' if vf['paragraph_count'][0]
                                    == vf['paragraph_count'][1] else '不一致'),
               '| pPr 改动 | %s |' % (vf['ppr_changed'] or '无'),
               '| 未改段 run 属性 | %s |' % (vf['run_props_changed_in_unchanged'] or '无'),
               '| 参考文献 | %s |' % ('完全一致' if vf['references_identical'] else '不一致'),
               '| 空段 | %d → %d |' % tuple(vf['empty_paragraphs']), '']
        if rs:
            md += ['## 四、报告驱动残留验收', '',
                   '标记片段 %d 处；残留 %d 处，其中**需继续处理 %d 处**，'
                   '其余为引文/专名等合规保留。'
                   % (rs['marked_count'], rs['residue_count'],
                      rs['risky_count']), '']
            md.append('| 标记 | 原长 | 最长残留 | 处置 | 片段 |')
            md.append('|---|---|---|---|---|')
            for it in rs['items']:
                md.append('| %s | %d | %d | %s | %s |'
                          % (it['label'], it['source_len'], it['max_residue'],
                             STATUS_CN[it['status']],
                             it['text'][:40].replace('|', '/')))
            md.append('')
        if sheet.get('confirm'):
            md += ['## 五、需要确认的事项', '']
            md += ['%d. %s' % (i, c) for i, c in enumerate(sheet['confirm'], 1)]
            md.append('')
        open(log_md, 'w', encoding='utf-8').write('\n'.join(md))

        print('')
        print('产出：')
        for f in (new_docx, cmp_docx, log_md):
            print('  %s' % f)
        return 0 if ok else 2

    return 1


if __name__ == '__main__':
    raise SystemExit(main())
