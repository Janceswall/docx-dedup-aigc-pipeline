# docx-dedup-aigc-pipeline

**本地 .docx 论文的「降查重率 + 降 AIGC 检测率」端到端流水线。**

改完之后格式不会坏——这不是"改完再修"，而是从机制上就不可能坏：
只改 `word/document.xml` 里的段落文本，styles / numbering / media / rels
等其余部件全部按**原字节**复制。

**零第三方依赖**，只用 Python 标准库。不需要 `pip install` 任何东西。

```bash
python3 scripts/dedup.py prepare 论文.docx -o work/     # 抽取 + 识别段落角色 + 生成工作表
#   ← 在 work/sheet.json 里填改写文本（人或 AI 来填）
python3 scripts/dedup.py run work/sheet.json -o out/    # 一条命令产出全部交付物
```

产出三份：

| 文件 | 用途 |
|---|---|
| `改写稿-*.docx` | 可直接提交复检 |
| `改写对照表-*.docx` | 横向 Word，逐段「原文 → 改写后 + 手法」，供逐字核对 |
| `自检报告-*.md` | 结论 + 重合检测 + 格式保真 + 残留验收数据 |

---

## 为什么用它

「降重」本身不难，难的是**改完怎么回到 .docx 里而不把排版弄坏**，以及
**怎么证明真的改到位了**。这个项目把这两件事做成了可验证的工程流程：

- **格式零破坏，靠机制而非靠小心**<br>
  只重写 `word/document.xml`，其余 ZIP 部件逐字节原样搬运。
  段落替换是 run 级的——新文本写进该段第一个含 `w:t` 的 run，
  继承它的字体字号，段落属性 `pPr` 因此 100% 不变。

- **有量化质量门，不靠手感**<br>
  知网连续 13 字相同即标红。`check` 命令把改写稿与原稿的所有 ≥13 字连续
  重合片段全部列出，并自动区分「引文/专名（合规保留）」与「需继续处理」。
  经验规律：单纯换词的"同义替换式"改写，第一轮往往仍留下**数十处**长重合，
  最长一处可达**上百字**——不靠检测器迭代 2~3 轮是压不干净的。
  （具体数量取决于文档与改写质量，故不给固定值。）

- **验收口径按轮次区分**<br>
  拿到检测报告后的第二轮，正确口径是「**被标记的句子是否被打断**」，
  而不是「新稿和上一版有多少重合」——后者会给出满屏假警报，
  因为上一版刻意保留的未标红文字本来就该重合。残留还会被自动分成
  「已打断 / 合规保留 / 需继续处理」三类，只有第三类为 0 才算过关。

- **对照表程序化生成**<br>
  逐段对照内容由程序从两份 .docx 直接抽取，没有人工转录，
  用户可以逐字核。

---

## 安装

无需安装。克隆下来即可用：

```bash
git clone https://github.com/Janceswall/docx-dedup-aigc-pipeline.git
cd docx-dedup-aigc-pipeline
python3 scripts/dedup.py --help
```

要求 Python **3.8+**，仅标准库。

---

## 用法

### 第一轮：预降重（还没有检测报告）

```bash
python3 scripts/dedup.py prepare 论文.docx -o work/
```

`work/sheet.json` 会自动给每个段落打上**角色**，并给出默认锁定建议：

| 角色 | 默认 | 说明 |
|---|---|---|
| `title` / `heading` / `keywords` | 🔒 锁定 | 题目与小标题不参与降重 |
| `reference` | 🔒 锁定 | 参考文献条目逐字保留 |
| `empty` | 🔒 锁定 | 空段要保留，段落数结构不能变 |
| `body` | ✅ 可改 | 正文 |

给要改的段落填 `rewrite`（整段新文本），没填的一律不动：

```json
{ "p": 5, "role": "body", "length": 147,
  "text": "原文……",
  "rewrite": "改写后的整段文本……",
  "techniques": "句式重构：长句拆短 + 加入实践者视角" }
```

然后：

```bash
python3 scripts/dedup.py run work/sheet.json -o out/
```

`run` 会依次执行：应用改写 → 13 字重合检测 → 格式保真校验 →（可选）残留验收
→ 生成对照表与自检报告，并给出**通过 / 仍有待处理项**的结论。

### 第二轮：按检测报告精修

```bash
python3 scripts/dedup.py report "PaperPass-检测报告/" --paras
python3 scripts/dedup.py run work/sheet_r2.json -o out2/ --report "PaperPass-检测报告/"
```

`report` 支持 PaperPass 报告包（解压后的目录）以及任意"标红版" `.docx`：

```
## 概览
  总相似度     : 12.0%
    来源拆解   : 期刊 9.0% / 网络 1.0% / 用户库 1.0% / 学位 1.0% / 其他 0.0%
  AIGC 疑似度  : 4.0%
  命中句子     : 18 / 120 句

## Top 相似来源
    3.2%  期刊 《某领域的通行研究综述》
    2.1%  网络 某主题的实践总结
```

> 上面的数字与来源标题只是**格式示例**，不是任何真实报告的内容。

第二轮的验收：

```bash
python3 scripts/dedup.py residue "PaperPass-检测报告/" 改写稿v2.docx
```

```
 = 残留(合规保留) 标记 18字 最长残留 18  [橙·相似] 某主题的实践观察与思考
OK 已打断       标记 42字 最长残留  0  [红·重复] 举例而言，被标记的整句在新稿中已被拆开，……
---
标记片段 18 处：残留 2 处，其中需继续处理 0 处
```

残留分三类：**已打断**（残留 0）、**合规保留**（残留在标题/引文/专名里，不计入待处理）、
**需继续处理**（其余长残留，必须回炉）。只有第三类为 0，这一轮才算过。

### 独立命令

```bash
python3 scripts/dedup.py roles    论文.docx                    # 只看段落角色
python3 scripts/dedup.py check    原稿.docx 改写稿.docx --only 2,4,5
python3 scripts/dedup.py verify   原稿.docx 改写稿.docx --changed 2,4,5
python3 scripts/dedup.py compare  原稿.docx 改写稿.docx -o 对照表.docx
```

`check` 默认跳过两稿一致的段落（标题、关键词、参考文献表、本轮没动的正文段），
只报真正改过的地方；要看全量加 `--all`。

所有子命令都支持 `--json`，便于接进其他流程。

### 演示

```bash
python3 examples/demo.py
```

会用一份合成的示例论文跑完整条链路，产出放在 `examples/out/`。

---

## 工作原理

```
论文.docx
   │  prepare ── docx_io 解析 word/document.xml，逐块抽取文本
   ▼             textutil 识别角色（标题/摘要/正文/参考文献/空段）
sheet.json ←─ 人或 AI 填 rewrite
   │  apply ─── 只改填了 rewrite 的段；run 级替换，继承 rPr
   ▼
改写稿.docx
   │  check ─── 与原文的 ≥13 字连续重合（引文/专名自动分类）
   │  residue ─ 报告标记句的最长残留（第二轮主判据）
   │  verify ── 段落数 / pPr / 未改段 run 属性 / 参考文献 / 空段
   │  compare ─ 程序化抽取 → 横向对照表 docx
   ▼
改写稿.docx + 对照表.docx + 自检报告.md
```

### 两条执行通道

| | 零依赖通道（默认） | 编辑器通道（可选） |
|---|---|---|
| 依赖 | 仅标准库 | 编辑器 SDK（如腾讯文档本地通道） |
| 替换方式 | 直接改写 `word/document.xml` | 按字符坐标调 API 替换 |
| 顺序 | 无序（持有元素引用） | 必须**从后往前**（坐标会漂移） |
| 优势 | 环境无要求、无坐标漂移风险 | 用户在编辑器里实时可见 |

两条通道产出的文档内容实测**完全等价**（逐段文本一致）。

### 内置子技能

Step 3 改写时按需加载：

- **`skills/text-paraphrase/`** —— 保意改写手法（压查重）：同义替换、句式转换、
  段落重构、引述转述的三层手法库，附前后对照示例。
- **`skills/de-ai-flavor/`** —— 去 AI 腔（压 AIGC）：识别并消除连接词堆叠、
  句式等长、完美并列、三段式排比、模板收尾等八类 AI 写作指纹。

两件事**同一轮做完**：只换词能降查重但对 AIGC 率几乎无效；只打散句式，
重合片段照样标红。

---

## 能力与限制

**能做**：本地 `.docx` / `.doc` 长文（论文、研究报告、各类正式文档）的保结构改写与验证。

**不能做 / 不做**：

- 不代写整篇论文，不篡改数据与结论，不洗稿
- 不保证具体分数——查重结果取决于检测系统算法，**以学校或期刊认可的系统复检为准**
- 含图片、域、文本框、公式、超链接的段落会被安全闸拒绝替换
  （这些内容无法从纯文本重建），需要人工处理

> ⚠️ 关于检测渠道：**不要**使用淘宝或第三方"知网查重"服务，
> 存在论文被收入比对库的风险，真到学校检测时会和自己撞车。
> 请走学校统一检测或知网官方个人入口。

---

## 项目结构

```
SKILL.md                        技能定义（Agent 读的入口）
scripts/
  dedup.py                      统一 CLI
  docx_io.py                    零依赖 docx 读写引擎
  docx_writer.py                从零构造 docx（对照表）
  textutil.py                   重合检测 / 引文分类 / 角色识别
  report_parser.py              检测报告解析
  extract_docx_text.py  ┐
  check_overlap.py      │
  verify_docx_integrity.py   ├─ 兼容壳，保留原命令行接口
  parse_paperpass_report.py  │
  verify_align.py       ┘       编辑器通道：坐标对齐校验
skills/
  text-paraphrase/              子技能：保意改写
  de-ai-flavor/                 子技能：去 AI 腔
references/
  workflow.md                   端到端流程 + 原理 + 全部坑位
  report-driven-round.md        第二轮精修详解
  aigc-reduction-techniques.md  两套手法合流说明
examples/
  demo.py                       一键跑通演示
```

---

## 致谢

`skills/de-ai-flavor` 内联自社区技能「去AI味助手」（原作者署名见该文件 frontmatter）。
`skills/text-paraphrase` 的手法库整理自论文降重改写的通用实践。

## 许可

MIT，见 [LICENSE](LICENSE)。
