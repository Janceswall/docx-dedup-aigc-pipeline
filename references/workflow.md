# 端到端工作流

本文是主技能的执行手册。SKILL.md 给的是"做什么"，这里给的是"怎么做、为什么这么做、
哪里会踩坑"。

---

## 0. 开工前的三个必问项

缺任何一项都不该动手：

1. **检测系统是哪一家**（知网 / 维普 / 万方 / PaperPass）
   各系统阈值与算法不同，以学校或期刊**指定**的那家为准。对着别家降是白费功夫。
2. **红线数字是多少**（如查重 ≤15%、AIGC ≤20%）
   目标不明确就无法判断"改到什么程度算够"。注意：**不必追求 0%**，
   降到红线以下留 5~8 个点的缓冲即可，改过头会伤学术表达的准确性。
3. **有没有标红检测报告**
   - **有报告** → 按标红位置精准精修（第二轮）
   - **没有报告** → 只能做**预降重**（按段落角色全文高风险段改写）

> ⚠️ **必须向用户说清楚这是两轮工作**，不要承诺一轮到位。
> 也别让用户用淘宝/第三方"知网查重"——有论文被收入比对库的风险，
> 真到学校检测时会和自己撞车。走学校统一检测或知网官方个人入口。

---

## 1. 两条执行通道

本工具链支持两条通道，产出的 .docx 内容完全等价（实测逐段文本一致）。

| | 零依赖通道（默认） | 编辑器通道（可选） |
|---|---|---|
| 依赖 | 仅 Python 标准库 | 腾讯文档本地通道 / editor_sdk |
| 替换方式 | 直接改写 `word/document.xml` | 调用编辑器 API 按字符坐标替换 |
| 顺序要求 | 无序（持有元素引用） | **必须从后往前**（坐标会漂移） |
| 用户可见性 | 改完打开文件才看到 | 在编辑器里**实时可见**，所见即所得 |
| 适用 | 批处理、CI、任何环境 | 需要在编辑器里实时校对 |

**默认用零依赖通道**——它对环境没有要求，而且从机制上排除了坐标漂移这一类错误。

### 编辑器通道要点（如果你要用）

先用 `doc_resolve_document_structure` 拿坐标，再用
`scripts/verify_align.py` 校验对齐：

```bash
# 1) 拿结构（file_id 用编辑器返回的）
<edsdk> call doc_resolve_document_structure \
  --json '{"file_id":"<file_id>","mode":"compact","limit":0}' > structure.json
# 2) 校验 span == len(text) + 1
python3 scripts/verify_align.py workspace_copy.docx structure.json
```

替换时必须**按段落号降序**处理：

```python
for p in sorted(REWRITES, reverse=True):        # 从后往前
    n = nodes[p - 1]
    call('doc_replace_text', {
        "file_id": FID,
        "ranges": [{"begin": n['start_index'], "end_index": n['end_index']}],
        "text": REWRITES[p]})
```

对齐关系是 `span = end_index - start_index = len(text) + 1`——
`end_index` 是**闭端**，末尾含段落结束符。所以可替换范围是
`[start_index, start_index + len(text))`。

**任何一段对不上就停下来排查**（可能混入表格 / 图片 / 域节点），
不要凭感觉往下替换。

---

## 2. 流程

```
  原稿.docx
     │
     │ ① prepare —— 抽取全文、识别段落角色、生成改写工作表
     ▼
  work/sheet.json  ←──── 人或 AI 在此填 rewrite 字段
     │
     │ ② apply —— 把改写落回 docx（只动填了 rewrite 的段）
     ▼
  改写稿.docx
     │
     ├─ ③ check —— 13 字连续重合检测（改写稿 vs 原稿）
     ├─ ④ residue —— 报告驱动验收（标记句是否被打断）※第二轮
     ├─ ⑤ verify —— 结构与格式保真校验
     └─ ⑥ compare —— 生成横向逐段对照表 docx
     ▼
  改写稿.docx + 改写对照表.docx + 自检报告.md
```

一条命令跑完 ②~⑥：

```bash
python3 scripts/dedup.py run work/sheet.json -o out/ \
    --report "PaperPass-检测报告/"        # 第二轮才需要 --report
```

### ① prepare

```bash
python3 scripts/dedup.py prepare 论文.docx -o work/
```

产出 `work/sheet.json`（改写工作表）与 `work/full_text.txt`（带段号全文）。

工作表会自动给每个段落打上**角色**：

| 角色 | 默认处理 |
|---|---|
| `title` 论文标题 | 🔒 锁定——题目不能改 |
| `heading` 小标题 | 🔒 锁定（报告标红时可单独考虑） |
| `keywords` 关键词 | 🔒 锁定 |
| `reference` 参考文献条目 | 🔒 锁定，逐字保留 |
| `empty` 空段 | 🔒 保留（段落数结构不能变） |
| `body` 正文 | ✅ 改写对象 |

锁定是**默认建议**，不是硬约束。要改标题就把 `locked` 改掉。

⚠️ `prepare` 会顺手做一次**可替换性体检**：含图片、域、文本框、超链接、
数学公式的段落会列入 `blocked` 并拒绝替换——这些内容无法从纯文本重建，
强行替换会把它们弄丢。遇到这种段落需要人工处理。

### ② apply

在 `sheet.json` 里给要改的段落填 `rewrite`（整段新文本），未填的一律不动：

```json
{ "p": 5, "role": "body", "length": 147,
  "text": "原文……", "rewrite": "改写后的整段文本……",
  "techniques": "句式重构：长句拆短 + 加入实践者视角" }
```

```bash
python3 scripts/dedup.py apply work/sheet.json -o 改写稿.docx
```

**前缀守卫**：原段以 `摘要：` / `关键词：` 开头而新文本漏掉时，会自动补回。
这是实战踩过的坑——整段替换极易吃掉标签。

### ③ check —— 13 字重合检测

知网连续 **13 字**相同即标红，所以这是"改到位没有"的量化判据：

```bash
python3 scripts/dedup.py check 原稿.docx 改写稿.docx -k 13 --only 2,4,5,7
```

输出把每个重合片段分成三类：

| 标记 | 含义 | 处理 |
|---|---|---|
| `[引文内-合规]` | 落在中文引号内 | 引文原文，**保留** |
| `[专名-须保留]` | 含书名号 | 书名/专名全称，**无法也不该替换** |
| `!! 需处理` | 其他长重合 | **必须回去做深度重构** |

不做这个分类会把三处无法规避的引用误报成问题，导致无意义的反复改写。

默认会**跳过两稿一致的段落**——标题、关键词、参考文献表、以及本轮没动的正文段，
它们本来就与原文逐字相同，报出来只是噪音。想看全量加 `--all`，
也可以反过来用 `--only 2,4,5` 只看本次改动的段。

**实践规律**：第一轮"同义替换式"改写通常仍留下大量长重合，
常为**数十处**、最长一处可达**上百字**（具体随文档与改写质量而变）。
**必须靠这个检测器迭代 2~3 轮**才能压干净。

### ④ residue —— 报告驱动验收（第二轮）

拿到检测报告后，**正确的验收口径是"被标记的句子是否被打断"**，
而不是"新稿和上一版有多少重合"——上一版刻意保留的未标红文字本来就该重合。

```bash
python3 scripts/dedup.py residue "PaperPass-检测报告/" 改写稿v2.docx
```

它对报告里每一处红色（重复）/橙色（相似）标记，检查在新稿全文中的
**最长残留连续长度**，要求 < 13 字。

残留要分两类看：落在**标题、关键词、参考文献**这类"本来就不该改"的段落里，
或本身是**引文原文、书名/专名全称**的，判为「合规保留」，不计入待处理项；
只有其余的长残留才需要回炉。不做这一步，交付结论会永远是"仍有待处理项"。

详见 `report-driven-round.md`。

### ⑤ verify —— 格式保真校验

```bash
python3 scripts/dedup.py verify 原稿.docx 改写稿.docx --changed 2,4,5,7
```

逐项检查：段落/块数量、改动清单、`pPr` 是否零改动、未改写段落的 run 属性
是否变化、参考文献条目、空段数量、字数变化。

⚠️ `pPr` / `rPr` 比对**必须做语义归一化**。编辑器会把 `<w:b/>` 序列化成
`<w:b w:val="1"/>`、把 `<w:kinsoku/>` 写成 `<w:kinsoku w:val="1"/>`——
语义完全等价，但字符串比对会报"全部段落格式都变了"的假警报。
`scripts/docx_io.py` 的 `canon_element()` 处理了这一点。

### ⑥ compare —— 生成对照表

```bash
python3 scripts/dedup.py compare 原稿.docx 改写稿.docx -o 对照表.docx --sheet work/sheet.json
```

产出**横向页面**的 Word：总体情况 + 逐段对照（段号/角色/原文/改写后/手法）+
未改动部分清单 + 需要确认的事项。表格内容是程序从两份 docx **直接抽取**的，
没有人工转录，可以逐字核。

---

## 3. 段落号口径

**段落号 = body 顶层块的序号，表格也占一个号。**

这与 Word 里"第 N 个块"的直觉一致。注意检测报告里的段号**不等于**文档段号——
报告会把空段去掉重新编号，两者通常差若干个空段。**先建映射再动手**。

---

## 4. 已知的坑（都踩过）

| 现象 | 原因 | 对策 |
|---|---|---|
| 抽不出全文，每段只有 66 字 | 编辑器结构预览的 `text_preview` 按字节折算，上限约 66 汉字 | 用 `scripts/extract_docx_text.py` 直接解析 `document.xml` |
| 编辑器替换后坐标错位 | `doc_replace_text` 会让后续坐标漂移 | 按段落号**降序**替换 |
| `摘要：` 前缀消失 | 整段替换吃掉前缀 | 前缀守卫自动补回；替换后重新抽取核对 |
| 格式校验报"全部段落都变了" | `<w:b/>` 与 `<w:b w:val="1"/>` 的序列化差异 | 语义归一化后再比 |
| 渲染成图报 file not found | 渲染服务不支持中文路径 | 复制成 ASCII 文件名再渲染 |
| 渲染服务报"繁忙/转换失败" | 云端渲染不稳定 | 改用 XML 层校验，别死等渲染 |
| 对照表内容叠了两遍 | 复用同一文档实例反复插入 | 每次重建文档都重新创建实例 |
| 报告标红版解析出"整篇全黑" | PaperPass 用 `w:color` 着色，**不是** `w:highlight` | 读 `w:color@val`：`F39800` 橙、`F12828` 红 |
| 解析到了 AIGC 标红版而非查重版 | 报告目录里两份文件名都含"标红版" | 文件名匹配按**优先级逐轮**匹配，不要按目录遍历顺序 |
| `create_doc` 后调 doc 工具报"document is not open" | 返回的 `file_id` 只是标识，实例还没读进来 | 先 `open_file(wait=True)` |
| 正则从 `create_doc` 返回值抠 `file_id` 抠错 | `(\S+)` 把尾随逗号也抠进去了 | 用 `file_id=([A-Za-z0-9_\-]+)` |

---

## 5. 交付话术要点

- 明确本轮是**预降重**还是**按标红报告精修**，别让用户误以为一轮就能达标
- 主动列出**需要用户拍板**的事项：被压缩的引文、疑似错误的官方名称、顺带修的笔误
- 说明改写对 **AIGC 检测**只有间接作用（打散句式、降低连接词密度、增加具体细节），
  最终数值以学校认可的系统复检为准
- 如实报告**剩余重合的来源**（标题、专名全称、引文原文），
  让用户知道数字降不下去的部分是为什么
- 提醒：锁定未改的标题/关键词若报告显示被标红，可再单独处理

---

## 6. 文件索引

| 文件 | 作用 |
|---|---|
| `scripts/dedup.py` | 统一 CLI：prepare / apply / check / residue / verify / compare / report / run / roles |
| `scripts/docx_io.py` | 零依赖 docx 读写引擎（文本抽取、run 级替换、保真保存、XML 规范化） |
| `scripts/docx_writer.py` | 从零构造 docx（生成对照表用，支持横向页面与表格） |
| `scripts/textutil.py` | 重合检测算法、引文/专名分类、段落角色识别 |
| `scripts/report_parser.py` | PaperPass 报告包与通用标红 docx 解析 |
| `scripts/extract_docx_text.py` | 兼容壳：抽取全文 |
| `scripts/check_overlap.py` | 兼容壳：13 字重合检测 |
| `scripts/verify_docx_integrity.py` | 兼容壳：格式保真校验 |
| `scripts/parse_paperpass_report.py` | 兼容壳：报告解析 |
| `scripts/verify_align.py` | 编辑器通道专用：校验字符坐标对齐 |
| `skills/text-paraphrase/` | 子技能：保意改写手法（压查重） |
| `skills/de-ai-flavor/` | 子技能：去 AI 腔（压 AIGC） |
| `references/aigc-reduction-techniques.md` | 两套手法的合流说明 |
| `references/report-driven-round.md` | 第二轮精修详解 |
