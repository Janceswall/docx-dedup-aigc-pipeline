---
name: docx-dedup-aigc-pipeline
version: 1.3.0
agent_created: true
display_name: DOCX 降重降 AIGC 流水线
display_name_en: DOCX Dedup & De-AIGC Pipeline
license: MIT
description_zh: "对本地 .docx 论文/报告做「降查重率 + 降 AIGC 检测率」的端到端流水线：分段改写 → 13 字重合检测迭代 → 报告驱动精修 → 格式保真校验 → 输出改写稿与逐段对照表。零第三方依赖，只改 document.xml，其余部件按原字节复制。"
description: "当用户要求对本地 Word 文档（.docx/.doc）降查重率、降 AIGC 检测率、降 AI 味，尤其是强调「保持原有结构不变」「输出对照表」「根据查重报告再改一次」时使用。触发词：降重、降AIGC、降AI率、去AI味、查重率太高、论文改写、标红段落处理、查重报告、PaperPass、知网查重、保持模板结构不变。不适用：纯文本层改写建议（用子技能 text-paraphrase）、只查重复率不做改写、代写整篇论文、篡改数据或结论。"
---

# DOCX 降重 + 降 AIGC 流水线

把一份本地 Word 文档改写成「查重率更低、AI 痕迹更弱」的版本，**结构零增删、格式零破坏**，
并交付一份可逐段核对的对照表。

## 这个技能解决什么问题

网上讲"怎么降重"的资料很多，讲"改完怎么回到 .docx 里而不把排版弄坏"的几乎没有。
本技能的差别就在于**落地**：段落怎么定位、格式怎么保、怎么验证真的改到位了、
怎么产出一份用户敢逐字核对的交付物。

| | 子技能 `text-paraphrase` | 子技能 `de-ai-flavor` | 本技能 |
|---|---|---|---|
| 管什么 | 换一种说法而不改意思 | 改得像人话而不是 AI | 定位 / 回写 / 验证 / 交付 |

三者是**协作**关系：本技能在 Step 3 按需加载两个子技能产出改写文本，
再自己负责把文本安全地落回文档并验证。

## 前置

- Python 3.8+（**仅标准库，无需 pip install 任何包**）
- 一份 `.docx` 文档
- 可选：查重 / AIGC 检测报告（PaperPass 报告包或任意"标红版" docx）

## 开工前必问三项

缺一项就不该动手：

1. **检测系统是哪一家**（知网 / 维普 / 万方 / PaperPass）——以学校或期刊指定者为准
2. **红线数字是多少**（如查重 ≤15%、AIGC ≤20%）——目标不明就无法判断改到什么程度算够
3. **有没有标红检测报告**
   - 有 → 走**第二轮**（报告驱动精修），只动被标记的句子
   - 没有 → 只能做**第一轮预降重**（按角色改写高风险段）

⚠️ **必须向用户说明这是两轮工作**，不要承诺一轮到位。
⚠️ 提醒用户：不要用淘宝/第三方"知网查重"，有论文被收入比对库的风险，会和自己撞车。

## 快速开始

```bash
# ① 抽取全文 + 识别段落角色 + 生成改写工作表
python3 scripts/dedup.py prepare 论文.docx -o work/

# ② 在 work/sheet.json 里给要改的段落填 rewrite 字段（人或 AI 来填）
#    只填要改的；没填的一律不动。可用子技能取改写文本。

# ③ 一条命令跑完：应用改写 → 重合检测 → 格式校验 → 生成对照表
python3 scripts/dedup.py run work/sheet.json -o out/

# ④ 拿到检测报告后的第二轮
python3 scripts/dedup.py report "PaperPass-检测报告/" --paras
python3 scripts/dedup.py run work/sheet_r2.json -o out2/ --report "PaperPass-检测报告/"
```

产出三个文件：`改写稿-*.docx`（可直接提交复检）、`改写对照表-*.docx`（横向逐段对照）、
`自检报告-*.md`（含结论与验收数据）。

## 铁律

1. **段落数、段落顺序、标题、关键词、参考文献表一律不动。**
   零依赖通道从机制上保证了这一点——它只改段落内的 run 文本，
   连 styles / numbering 等部件都是按原字节复制过去的。
2. **绝不覆盖原稿。** 先复制工作副本，输出一律写到新文件。
   （`docx_io.save()` 会拒绝写入源文件路径。）
3. **不改数据、不改结论、不改人名年份、不篡改引文内容。**
   压缩引文必须用省略号标明删节。
4. **引文原文、书名、专名全称（标准/法规/机构名）属合规保留对象**，不算"没改干净"，
   交付时单独列出说明。
5. **剩余重合要如实报告来源**，让用户知道数字降不下去的部分是为什么。

## 执行流程

详细做法、原理与坑位见 **`references/workflow.md`**。要点：

| 步骤 | 命令 | 判据 |
|---|---|---|
| ① prepare | `dedup.py prepare` | 段落角色自动识别，锁定段默认不改 |
| ② apply | `dedup.py apply` | 只动填了 `rewrite` 的段；自动保留 `摘要：` 等前缀 |
| ③ check | `dedup.py check` | ≥13 字连续重合必须清零（引文/专名除外） |
| ④ residue | `dedup.py residue` | 标记句的最长残留 <13 字；落在标题/引文/专名内的残留属合规保留，不计入待处理（第二轮主判据） |
| ⑤ verify | `dedup.py verify` | 段落数一致、`pPr` 零改动、参考文献逐字节相同 |
| ⑥ compare | `dedup.py compare` | 程序化抽取，绝不人工转录 |

**核心质量门是 ③/④ 的迭代**：第一轮"同义替换式"改写通常仍留下大量长重合
（经验值：常为数十处，最长一处可达上百字；数量随文档与改写质量而变），
必须改 2~3 轮才能压干净。
**不要跳过检测器凭感觉判断"改够了"。**

### 验收口径按轮次区分

- **第一轮（无报告）**：看 `check`——改写稿与原稿是否还有 ≥13 字的非引用重合
- **第二轮（有报告）**：看 `residue`——**不是**看"和上一版有多少重合"。
  上一版刻意保留的未标红文字本来就该重合，用错口径会得到满屏假警报。

## 子技能

本技能自带两个子技能，Step 3 改写时按需加载：

| 子技能 | 路径 | 作用 |
|---|---|---|
| 保意改写（压查重） | `skills/text-paraphrase/` | 同义替换、句式转换、段落重构、引述转述 |
| 去 AI 腔（压 AIGC） | `skills/de-ai-flavor/` | 消除连接词堆叠、句式等长、模板收尾等八类 AI 指纹 |

**两件事必须同一轮做完**：只换词能降查重但对 AIGC 率几乎无效；
只打散句式则重合片段照样标红。合流手法见
`references/aigc-reduction-techniques.md`。

## 两条执行通道

默认走**零依赖通道**（纯 Python 直接改写 `word/document.xml`），
环境无要求、替换顺序无关。若在带编辑器 SDK 的环境里、需要用户实时看到改动，
可以改走**编辑器通道**——那时必须按段落号**降序**替换，并先用
`scripts/verify_align.py` 校验字符坐标对齐。两条通道产出的文档内容实测完全等价。

## 文件一览

```
SKILL.md                       本文件
README.md                      开源仓库说明
scripts/
  dedup.py                     ★ 统一 CLI（prepare/apply/check/residue/verify/compare/report/run/roles）
  docx_io.py                   ★ 零依赖 docx 读写引擎（文本抽取 / run 级替换 / 保真保存）
  docx_writer.py               从零构造 docx（生成横向对照表）
  textutil.py                  重合检测算法 / 引文分类 / 段落角色识别
  report_parser.py             PaperPass 报告包与通用标红 docx 解析
  extract_docx_text.py         兼容壳：抽全文
  check_overlap.py             兼容壳：13 字重合检测
  verify_docx_integrity.py     兼容壳：格式保真校验
  parse_paperpass_report.py    兼容壳：报告解析
  verify_align.py              编辑器通道专用：校验字符坐标对齐
skills/
  text-paraphrase/             子技能：保意改写手法
  de-ai-flavor/                子技能：去 AI 腔
references/
  workflow.md                  ★ 端到端流程、原理、全部坑位
  report-driven-round.md       ★ 第二轮精修详解
  aigc-reduction-techniques.md 两套手法的合流说明
examples/                      合成示例与自检脚本
```

## 边界

**不适用**：纯文本层改写建议（用子技能 `text-paraphrase`）、只测重复率不做改写、
代写整篇论文、篡改数据或结论。

**效果说明**：改写对 AIGC 检测只有间接作用（打散句式、降低连接词密度、增加具体细节），
最终数值**以学校或期刊认可的系统复检为准**，不要向用户承诺具体分数。
