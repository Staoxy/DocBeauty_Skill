# DocBeauty Skill V2 — AI 可读技术设计方案

- **版本**：V2.0（基于 V1 审查结论 + 开源竞品调研全面修订）
- **项目类型**：Agent Tool / Skill
- **核心技术**：Python ≥3.9 + python-docx ≥1.1.0 + lxml + jsonschema
- **主要输入**：DOCX
- **主要输出**：DOCX + JSON 处理报告（含内容一致性证明）
- **阅读指引**：Agent 只需读 §0、§3、§19、§21；实现者通读全文，重点 §5–§18、§22–§24。

---

## 0. 快速参考（必读）

**一句话定位**：DocBeauty 是一个可以被 AI Agent 调用的**确定性 Word 排版引擎**——AI 负责理解用户，DocBeauty 负责在不改变内容的前提下执行排版，并用 JSON 报告自证"内容没有被改动"。

**CLI**：

```text
python main.py <input.docx> [-c config.json] [-o OUTDIR] [--report PATH] [--dry-run] [--batch DIR]
```

**退出码**：`0` 成功 · `2` 部分成功 · `3` 内容护栏失败（**不产出文件**）· `1` 错误。

**stdout** 只输出 report.json 内容（UTF-8）；人类可读日志走 stderr。

**五条默认行为**（Agent 必须知道）：

1. `content_protection` 默认 `true`：除白名单区域（目录、页码域）外，**任何正文文字变化都会导致拒绝产出文件**。
2. `document_type`/`style` 默认 `auto`：Skill 自动检测；不确定时回退 `generic` + `formal`，绝不强行套模板。
3. 已有页码、已有目录域、页眉页脚：**默认不覆盖、不删除**，只补缺失项。
4. 标点/符号规范化**默认关闭**（它会改文字），必须显式开启且关闭内容保护才生效。
5. `status != "success"` 时，Agent 必须如实转述 report.json 中的 warnings/errors，不得宣称成功。

---

## 1. 项目定位

DocBeauty 不是"一个会写 Word 的 AI"，而是：

> **一个可以被 AI Agent 调用的确定性 Word 文档处理引擎。**

- 核心目标：在尽可能不改变用户原始内容的前提下，自动分析 DOCX 结构，规范化、美化和排版。
- V1 的三条定位保留不变：独立可运行、可被 Agent 结构化调用、输出机器可读报告。
- V2 新增强调：**报告必须能自证内容一致性**（指纹 + 逐项 diff），这是与全部已知竞品的本质差异（竞品调研见 §28 依据列）。

## 2. 核心原则

### 2.1 内容保护原则

默认**只修改格式，不修改内容**。禁止默认执行：改写句子、增删正文、修改事实/数字/引用/标题文字。

**允许的操作白名单**（都会在报告 `content_guard.excluded_generated` 中逐项登记）：

- 插入目录（"目录"标题 + TOC 域 + 可选提示行 + 分页）
- 页脚/页眉中插入页码域
- 删除满足 §12 条件的空白段落
- 页面设置、字体、段落、样式、表格样式、图片尺寸

**唯一的例外操作**：`punctuation_normalize`（§10.5），默认关闭，开启时必须同时 `content_protection=false`，且每个改动逐条进入报告。

### 2.2 AI 负责理解，代码负责执行

- LLM/Agent：理解需求、读用户的格式要求文档、判断文档类型、生成结构化参数（含 `overrides`）、向用户转述报告。
- DocBeauty：读取/分析/修改/生成 DOCX、自检、出报告。
- **禁止**：LLM 直接生成或修改 Word XML；Skill 改写正文文字；Skill 自行解析格式要求文档（那是 Agent 的活，Skill 只接收合并后的参数）。
- 借鉴 Office-Word-MCP-Server 的教训（45+ 细粒度工具导致 Agent 几十次调用难以收敛）：DocBeauty 提供**少量高层原子操作**，一次调用完成一类任务。

### 2.3 确定性与幂等

- 相同输入 + 相同配置 → 相同输出（报告中不含随机时序影响的字段，时间戳单独存放）。
- **幂等**：对 `processed.docx` 再跑一次，报告应显示各操作 `changed ≈ 0`（§24 有专门测试）。

### 2.4 失败不落盘

- 全程在临时目录的**文件副本**上操作；只有质量检查 + 内容护栏全部通过后才原子写出（写 `.tmp` 后 `os.replace`）。
- 护栏失败或致命错误 → **删除临时输出，不产出 DOCX**，报告如实说明。
- 单个非致命操作失败 → 继续，标记 `partial_success`，逐项进 `errors[]`。

### 2.5 边界诚实

Skill 处理不了的東西（浮动图片、文本框、公式、脚注内容、宏）必须在报告中**明确列出"未处理"**，不得假装已处理、也不得静默破坏。

### 2.6 不懂的不动

对 XML 中不理解的部分（未知部件、自定义 XML、宏、内容控件 w:sdt）保持原样。只通过 python-docx/lxml DOM 修改，禁止对包内 XML 字节做正则替换（§22）。

## 3. 输入

### 3.1 DOCX 文件（必需）

- 仅支持 `.docx`。`.doc`/`.wps`/加密文档 → 报错（错误码见 §20；`.doc/.wps` 的 COM 转换是 Phase 7 可选项，V2 不承诺）。
- 输入前校验：存在性、大小（默认 >100MB 报 `FILE_TOO_LARGE`，可配）、能否作为 ZIP 打开、能否被 python-docx 加载、是否加密（`FILE_NOT_FOUND` / `UNSUPPORTED_FILE` / `FILE_TOO_LARGE` / `INVALID_DOCX` / `PASSWORD_PROTECTED`）。

### 3.2 用户自然语言要求（可选，Agent 的输入而非 Skill 的输入）

例如"按课程论文格式排版，不要改正文"。Agent 将其翻译为 §3.4 的结构化配置。

### 3.3 格式要求文件（可选，`format_file`）

用户常会给出学校格式要求（docx/pdf/网页）。**Skill 不读它**。Agent 负责阅读并把其中的格式规定（字体、字号、行距、页边距、页码等）翻译为 `overrides` JSON 传入（借鉴 paper_format_agent 的 `--format-file` 思路，但把"读文档"这一步交还给 LLM——这正是 §2.2 的分工）。`format_file` 字段仅用于报告回显追溯。

### 3.4 结构化配置（完整 Schema）

```json
{
  "task": "beautify",
  "input_file": "/path/to/input.docx",
  "output_dir": "./output",
  "document_type": "auto | academic | experiment | report | meeting | generic",
  "style": "auto | academic | formal | clean",
  "content_protection": true,
  "operations": "auto" | ["normalize_font", "format_headings", "..."],
  "format_file": null,
  "overrides": { "font/paragraph/page/headings/captions/tables/page_number/toc 的任意子集，见 §9.7" },
  "options": {
    "delete_blank_paragraphs": true,
    "heading_promote": "off | high_confidence | all",
    "strip_style_numbering": true,
    "keep_inline_emphasis": true,
    "page_number_skip_existing": true,
    "table_repeat_header": true,
    "table_borders": "auto | grid | threeline | none",
    "image_max_width_percent": 100,
    "toc_replace_manual": false,
    "toc_refresh_hint": true,
    "toc_page_break": true,
    "punctuation_normalize": false
  }
}
```

**字段职责界定**（消除 V1 的歧义）：

- `style` 是**排版参数的唯一来源**（模板选择）。
- `document_type` 只影响**检测与默认模板映射**（§9.5），本身不携带排版参数。
- `operations` 允许 `"auto"`（字符串）或数组（schema 用 `oneOf`）。
- 配置经 jsonschema 校验，非法即 `CONFIG_INVALID`，不产出任何文件。

**操作目录**（输入动词 ↔ 报告完成时态固定映射）：

| 输入 operations 值 | 报告 operations 值 | 说明 |
|---|---|---|
| `page_setup` | `page_setup_done` | 页面尺寸/页边距 |
| `detect_headings` | `headings_detected` | 结构识别（含提升） |
| `format_headings` | `headings_formatted` | 应用 Heading 样式 |
| `normalize_font` | `font_normalized` | 字体统一 |
| `normalize_paragraph` | `paragraph_normalized` | 段落统一 |
| `normalize_lists` | `lists_normalized` | 列表策略 |
| `format_captions` | `captions_formatted` | 图注/表注 |
| `format_tables` | `tables_formatted` | 表格美化 |
| `format_images` | `images_formatted` | 图片规范化 |
| `page_number` | `page_number_added` | 页码 |
| `header_footer` | `header_footer_done` | 页眉（可选文字） |
| `toc` | `toc_added` | 原生目录 |
| `cleanup_blank_paragraphs` | `blank_paragraphs_cleaned` | 空段清理 |
| `punctuation_normalize` | `punctuation_normalized` | 仅显式开启时可用 |
| `diagnose` | `diagnosis_done` | 只诊断不修改（配 dry-run） |

`operations: "auto"` 时按 §5.2 全序执行，但逐个操作带**前置条件**（如 `toc` 需要 §16 条件满足），条件不满足则记 `skipped` + 原因，不算失败。

## 4. 输出

```text
<output_dir>/
├── <输入文件名>_beautified.docx    # 保留原文件名便于用户对认
└── <输入文件名>_report.json        # 与 docx 同名，避免批量时混淆
```

- 原子写入：先写 `<name>.tmp`，全部检查通过后 `os.replace`。
- `--dry-run` 时不产出 docx，只产出 `status="dry_run"` 的报告。

## 5. 架构与执行流水线

### 5.1 模块图

```text
Agent / CLI
    │ config.json
    ▼
main.py ──► config.py（schema 校验 + overrides 合并）
    ▼
analyzer.py ──► structure.py（区域 + 标题 + 图注 + 列表 → ExclusionZones）
    ▼
guard.py（基线快照，任何修改之前）
    ▼
┌────────────── 排版操作层（固定顺序，见 5.2） ──────────────┐
│ layout.py → headings.py → formatter.py → lists.py        │
│ → captions(结构内) → tables.py → images.py → toc.py      │
│ → page_number(layout.py) → cleanup.py                    │
└──────────────────────────────────────────────────────────┘
    ▼
checker.py（Q01–Q16） ──► guard.py（护栏校验） ──► 原子写盘
    ▼
report.py ──► report.json
```

### 5.2 固定执行顺序（实现者必须按此排序）

| # | 步骤 | 模块 | 是否修改文档 | 排序原因 |
|---|---|---|---|---|
| 1 | 加载 + 护栏基线快照 | guard | 否 | 快照必须在任何修改前 |
| 2 | 文档分析 | analyzer | 否 | 提供后续全部依据 |
| 3 | 结构检测（区域/标题/图注/列表/已有目录/已有页码） | structure | 否 | 产出 ExclusionZones |
| 4 | 配置解析与合并（模板 + overrides + options） | config | 否 | — |
| 5 | 页面设置 | layout | 是 | 页边距定了才有"可用宽度"，表格/图片依赖它 |
| 6 | 标题样式应用 | headings | 是 | 先定标题身份，段落归一才能正确跳过标题 |
| 7 | 字体统一 | formatter | 是 | — |
| 8 | 段落统一 | formatter | 是 | 依赖 #6 的标题集合与 ExclusionZones |
| 9 | 列表策略 | lists | 是 | — |
| 10 | 图注/表注格式 | structure+formatter | 是 | — |
| 11 | 表格美化 | tables | 是 | 依赖 #5 的可用宽度 |
| 12 | 图片规范化 | images | 是 | 依赖 #5 |
| 13 | 目录 | toc | 是 | 依赖 #6 的最终标题集合 |
| 14 | 页码/页眉 | layout | 是 | 在 TOC 之后，方便三段式分区 |
| 15 | 空白段落清理 | cleanup | 是 | 最后清理，避免影响前面各步的段落索引 |
| 16 | 质量检查 → 内容护栏 → 原子写盘 → 报告 | checker/guard/report | 写盘 | — |

任何一步抛出未捕获异常：按 §20 分类——单操作失败降级 `partial_success`；写盘阶段失败 → 不产出文件。

### 5.3 关键架构决策（实现前必须确认）

| 编号 | 决策 | 内容 |
|---|---|---|
| D1 | **样式双轨制** | 同时做两件事：①修改 styles.xml 中样式定义（Normal、Heading 1–3、Caption、Table）与 docDefaults；②清除正文 run/段落上的**残留直接格式**，让样式真正生效。只做①会被直接格式压制（"设了样式没变化"），只做②会让用户后续编辑困难。 |
| D2 | **ExclusionZones 一级公民** | 结构检测产出统一的"排除区注册表"（§7.2），所有格式化操作必须先查它。 |
| D3 | **护栏三段式** | 修改前快照（步骤1）→ 修改后校验（步骤16）→ 失败不写盘。 |
| D4 | **原子输出** | 临时目录副本操作；`os.replace` 落盘。 |
| D5 | **高层原子操作** | 面向 Agent 的操作粒度 = "一类任务一次调用"，不做细粒度 set-font-of-char 工具。 |
| D6 | **OOXML 红线** | 只经 python-docx/lxml DOM 修改；禁止对包内 XML 字节做正则/字符串替换（§22）。 |
| D7 | **幂等目标** | 重复运行不产生视觉变化，报告能识别"已符合"状态。 |

## 6. Document Analyzer

`analyzer.py` 输出结构化 `analysis` 对象（进入报告）。V1 指标全部保留，V2 扩充为：

```json
{
  "paragraph_count": 126,
  "table_count": 4,
  "inline_image_count": 7,
  "floating_image_count": 1,
  "textbox_count": 0,
  "omml_equation_count": 3,
  "footnote_count": 12,
  "hyperlink_count": 12,
  "heading_counts": {"h1": 3, "h2": 8, "h3": 6, "outline_lvl_other": 1},
  "heading_counts_by_text_pattern": {"cn_gov_l1": 3, "cn_gov_l2": 8, "l3": 6},
  "style_usage_histogram": {"Normal": 90, "Heading 1": 3, "自定义样式A": 12},
  "numbering": {"styles_with_numPr": ["Heading 1", "Heading 2"], "num_ids_in_use": [1, 4]},
  "existing_toc_field": false,
  "existing_manual_toc": false,
  "existing_page_number_fields": ["footer2.xml"],
  "section_count": 2,
  "sections_pgnumtype": [null, {"fmt": "decimal", "start": 1}],
  "tracked_changes": {"ins": 5, "del": 2},
  "comment_count": 4,
  "cjk_char_ratio": 0.93,
  "font_size_histogram_pt": {"12": 90, "14": 8, "16": 3},
  "cover_like_first_section": true,
  "has_encryption": false
}
```

新增指标的用途：

- `styles_with_numPr`：§10.3 防双重编号的输入。
- `existing_toc_field` / `existing_page_number_fields`：§16/§11.2 跳过依据。
- `tracked_changes` / `comment_count` > 0 → 报告 `warnings`："文档包含未接受的修订/批注，排版结果以 Word 打开接受修订后可能变化，建议先接受修订再排版"（含修订的文档不拒绝处理，但必须警告——竞品普遍忽略这一点）。
- `floating_image_count` / `textbox_count` / `omml_equation_count` / `footnote_count`：§2.5 边界声明的依据。
- `cjk_char_ratio` < 0.3 时，正文西文字体权重提高（模板仍生效，但警告用户当前模板以中文排版为假设）。
- `cover_like_first_section`（首节段落数 ≤ 10 且存在 ≥22pt 居中段落）：封面识别信号之一。

## 7. Structure Detection 与 ExclusionZones

### 7.1 区域识别

| 区域 | 识别规则 | 后续策略 |
|---|---|---|
| 封面 cover | 首节 + `cover_like_first_section` + 大字号居中段落群；或到第一个 Heading/分节符为止 | 整区进排除区：不做缩进/对齐/字体归一（封面格式被正文归一化毁掉是竞品通病） |
| 前置 front_matter | 摘要/Abstract/关键词/Keywords 标题至正文第一个 Heading 前 | 摘要标题套 Heading 样式但**不编号、进目录可选**；正文段仍参与归一 |
| 已有目录区 toc_existing | 存在 TOC 域 → 域所在段落群；或"目录"标题 + 后续点线/页码模式段落群（手工目录） | TOC 域：只补 `updateFields`；手工目录：默认仅警告（`toc_replace_manual=false`），不删除 |
| 参考文献 references | 标题匹配 `^(参考文献|References|Bibliography)` 至下一 Heading | 标题套 Heading；条目段落**跳过首行缩进**（悬挂缩进不动），仍做字体统一 |
| 附录 appendix | `^(附录|Appendix)` 至文档尾 | 同正文 |

区域识别**只做归类，不删除、不改文字**。低置信度区域标注 `confidence` 并放入报告。

### 7.2 ExclusionZones 注册表（D2）

结构检测的最终产物，所有格式化操作必须查询：

```json
{
  "paragraphs_skip_indent":    [],   // 列表项、图注、参考文献条目、封面段、目录段
  "paragraphs_skip_all":       [],   // 封面段、目录域段（除字体外的所有归一都跳过）
  "paragraphs_are_captions":   [],
  "paragraphs_are_headings":   [],   // 含"拟提升"候选
  "table_cell_paragraphs":     true, // 表格内段落一律跳过正文缩进/对齐规则
  "runs_keep_emphasis":        true  // 全局：粗/斜/下划线/颜色保留
}
```

## 8. Heading Detection

### 8.1 检测优先级（从高到低）

1. **原生 Heading 样式**（pStyle = Heading1–3）→ 直接采信。
2. **大纲级别** `w:outlineLvl` 0–2 → 视为 H1–H3。
3. **挂了多级列表编号的段落**（numPr + 编号模板匹配）→ 按编号层级映射。
4. **文本模式启发式**（§8.3）→ 打分制，见 §8.5。

### 8.2 编号模式模板（可配置，兼容 Word-Formatter-Pro 的四级体系）

```json
{
  "cn_zhang":  {"l1": "^第[一二三四五六七八九十百\\d]+章", "l2": "^第[一二三四五六七八九十百\\d]+节", "l3": "^[一二三四五六七八九十]+、", "l4": "^（[一二三四五六七八九十]+）"},
  "cn_gov":    {"l1": "^[一二三四五六七八九十]+、", "l2": "^（[一二三四五六七八九十]+）", "l3": "^(\\d+)[\\.、]\\s", "l4": "^（\\d+）"},
  "numeric":   {"l1": "^(\\d+)[\\.、]\\s", "l2": "^\\d+\\.\\d+", "l3": "^\\d+\\.\\d+\\.\\d+", "l4": "^\\d+\\.\\d+\\.\\d+\\.\\d+"},
  "auto_detect": true
}
```

`auto_detect`：对候选段落试配每个模板，选"层级序列最连续（相邻级差 ≤1）且命中数最多"的模板；都不理想 → 只采信原生样式，启发式结果降级为候选。

### 8.3 启发式打分（正文段落 → 标题候选）

| 信号 | 分值 |
|---|---|
| 命中编号模板（且层级序列连续） | +3 |
| 段落整体加粗 | +2 |
| 字号 ≥ 正文众数 + 1pt | +2 |
| 长度 ≤ 40 字符 | +1 |
| 不以句末标点（。！？…）结尾 | +1 |
| 以句末标点结尾 | **-4（一票否决）** |
| 命中图注/表注模式（§15） | **一票否决**（图注检测先于标题启发式执行） |
| 以冒号结尾且长度 ≤ 12（如"注意："） | **一票否决**（强调行，不是标题） |
| 段落含图片/域/书签 | 不参与 |

### 8.4 低置信度处理（把 V1"不确定时不得破坏"落成机制）

- 得分 ≥5：高置信度，按 `options.heading_promote`（默认 `high_confidence`）自动套 Heading 样式。
- 得分 3–4：**只进报告** `candidates` 列表（段落序号 + 文本 + 得分 + 建议级别），不修改。
- `heading_promote: "off"`：全部只报告不修改；`"all"`：候选也自动套（仅当用户明确要求）。
- 层级跳跃检测：最终标题序列出现 H1→H3 等跳级 → 质量检查 Q08 记 warning（借鉴 document-format-skills 的层级跳跃诊断）。

## 9. 排版模板

### 9.1 参数表约定

每个模板是一个**扁平参数字典**（`templates/academic.py` 等），overrides 按 §9.7 深合并。字号同时给 pt 与中文名；缩进用"字符"单位（实现见 §10.2）。

### 9.2 academic（课程论文 / 学术作业 / 调查报告 / 实验报告）

| 参数 | 值 |
|---|---|
| 页面 | A4；边距 上下 2.5cm、左 3.0cm、右 2.5cm |
| 正文 | eastAsia 宋体 / ascii+hAnsi Times New Roman / 小四 12pt / 1.5 倍行距 / 首行缩进 2 字符 / 两端对齐 |
| H1 | eastAsia 黑体 / 三号 16pt / 居中 / 段前 24pt 段后 18pt / keepNext |
| H2 | 黑体 / 四号 14pt / 左对齐 / 段前 12pt 段后 6pt / keepNext |
| H3 | 黑体 / 小四 12pt / 左对齐 / 段前 6pt 段后 6pt / keepNext |
| 图注表注 | 宋体 / 五号 10.5pt / 居中 |
| 表格 | 三线表（顶/底 1.5pt、栏目线 0.75pt、无竖线）；表头行加粗居中 + `tblHeader` 重复 |
| 页码 | 页脚居中 / 宋体小五 9pt |
| 目录 | TOC 域 `\o "1-3" \h \z \u`；学术文档检测到封面时启用三段式页码（§11.2） |

### 9.3 formal（正式报告 / 会议纪要 / 工作材料）

A4；边距 2.54cm 四边；正文宋体小四 12pt / 1.5 倍 / 首行缩进 2 字符 / 两端对齐；H1 黑体三号左对齐段前 18 段后 12；H2 黑体四号段前 12 段后 6；H3 黑体小四段前 6 段后 6；表格全边框 grid；页脚居中小五；其余同 academic。

### 9.4 clean（资料 / 笔记 / 长文档）

A4；边距 2.54cm；正文等线 11pt / 1.3 倍 / **无首行缩进** / 段后 6pt / 两端对齐；H1 微软雅黑 14pt 加粗段前 12 段后 6；H2 微软雅黑 12pt 加粗；H3 等线 11pt 加粗；表格 grid；页脚居中。

### 9.5 document_type ↔ style 映射（修复 V1 中 meeting 无模板的不一致）

| document_type | 默认 style | 典型证据 |
|---|---|---|
| academic / experiment | academic | 摘要+关键词+参考文献 ≥2 项；实验目的/步骤/结果 ≥2 项 |
| report / meeting | formal | 会议时间+参会人员+议题 ≥2 项；报告/总结/分析类词汇 |
| generic | formal | 证据不足时的**安全回退**（V1 规则保留） |

自动检测只在 `document_type=auto` 时执行；证据与置信度写入报告 `config.document_type`。

### 9.7 覆盖机制（补上 V1 缺失的参数通道）

```json
"overrides": {
  "font":    {"chinese": "仿宋", "western": "Times New Roman", "size_pt": 12, "size_cn": "小四"},
  "paragraph": {"line_spacing": 1.5, "first_line_chars": 2, "align": "justify", "before_pt": 0, "after_pt": 0},
  "page":    {"margins_cm": {"top": 2.5, "bottom": 2.5, "left": 3.0, "right": 2.5}},
  "headings": {"h1": {"font": "黑体", "size_pt": 16, "align": "center"}, "h2": {}, "h3": {}},
  "captions": {"font": "宋体", "size_pt": 10.5},
  "tables":  {"borders": "threeline", "repeat_header": true},
  "page_number": {"position": "footer", "align": "center", "font": "宋体", "size_pt": 9},
  "toc":     {"levels": "1-3"}
}
```

合并规则：深合并（dict 递归、标量覆盖），`overrides` > 模板默认；`size_cn` 与 `size_pt` 同时给出时以 `size_pt` 为准并校验一致性（映射见附录 B）。合并结果整体回显进报告 `config.overrides_applied`。

## 10. Formatting Engine（附 OOXML 实现细则）

### 10.1 字体统一

**核心规则**（python-docx 不会替你做的事，必须经 lxml 完成，参见附录 C-1）：

- 同时设置 `w:rFonts` 的 `w:ascii` + `w:hAnsi`（西文）与 `w:eastAsia`（中文）；**必须删除** `w:asciiTheme` / `w:hAnsiTheme` / `w:eastAsiaTheme` / `w:cstheme` 主题字体引用，否则部分文档设置不生效。
- 双轨（D1）：①修改 Normal 样式与 docDefaults 的 rFonts（让后续新输入继承）；②逐 run 清理残留直接字体（含 rStyle 带入的字体覆盖）。
- **强调保留清单**（`keep_inline_emphasis=true` 默认）：加粗 `w:b`、斜体 `w:i`、下划线 `w:u`、删除线、颜色 `w:color`、高亮 `w:highlight`、上下标 `w:vertAlign` 全部保留——字体归一只动 rFonts 与 sz/szCs。此清单来自 Word-Formatter-Pro 的实证需求。
- 超链接 run：正常设置字体，但**不得动其 Hyperlink 字符样式**（颜色/下划线由样式控制）。
- 公式（m:oMath 内部 run）、文本框（w:txbxContent）、脚注内容：**不处理、不破坏**（§2.5）。
- 中文字号映射表见附录 B；`size` 一律 pt，报告与 schema 不接受含糊单位。

### 10.2 段落统一

- 行距：`w:spacing w:line="360" w:lineRule="auto"` = 1.5 倍（python-docx `line_spacing=1.5` 可用）。
- 段前/段后：pt（`w:before`/`w:after`）；**不用** `beforeLines/afterLines`（渲染器支持差）。
- 首行缩进 2 字符：**必须写 `w:ind w:firstLineChars="200"`**，并附带 `w:firstLine="480"`（2 × 12pt 字号换算 twips）作为忽略 Chars 属性的老渲染器回退（附录 C-2）。禁止只写 twips——字号不同的文档会缩进错位。
- 对齐：正文 `w:jc="both"`；仅应用于**非排除区**段落（§7.2）。
- 表格单元格内段落：跳过缩进/对齐规则（只做字体统一与垂直居中，见 §13）。

### 10.3 标题样式应用算法（防双重编号是硬性要求）

对每个被采信的标题段落，按序执行：

1. **查目标样式 numPr**（analyzer 的 `styles_with_numPr`）：若 Heading 样式绑定了编号定义，在该样式 pPr 写入 `<w:numPr><w:numId w:val="0"/></w:numPr>` 禁用继承编号（numId=0 = 无编号，比删除样式属性安全），记入报告 `numbering_disabled_on_styles`。
2. 段落自身若有 `w:numPr`（自动编号）：移除段落级 numPr（这是格式不是文字），记入报告。**手工编号文字属于内容，绝不删除**——与样式编号叠加会造成"一、一、研究背景"，本步骤是唯一防线。
3. 清除段落直接格式（缩进、对齐、spacing、边框）与 run 直接格式（rFonts/sz/颜色/主题字体），使样式完全接管。
4. `pStyle = Heading N`；确认样式携带 `w:outlineLvl`。
5. 段落 pPr 加 `w:keepNext` + `w:keepLines`（§11.4 标题孤行）。
6. 若 run 仍有 rStyle 字符样式干扰字体：按模板补写 run 级 rFonts/eastAsia 兜底（可关）。
7. 登记进 ExclusionZones 的 `paragraphs_are_headings`。

### 10.4 列表策略

- 识别：段落带 `w:numPr`，或以项目符号字符（•·▪◦-–）开头。
- 策略：跳过首行缩进与两端对齐；字体统一正常应用；**不重排编号、不改编号文字、不动 numId 定义**（编号重排 = 改内容，禁止）。
- 编号与正文同段混排的段落（"1. xxx 内容…"）不视为列表，参与正文归一。

### 10.5 标点/符号规范化（唯一的改内容操作，默认关闭）

- 开启条件：`options.punctuation_normalize=true` **且** `content_protection=false`，否则 `CONFIG_INVALID`。
- 范围：中英文括号/引号/冒号等按上下文切换（借鉴 document-format-skills）。
- **保护例外**（绝不修改）：数字与单位间（100%）、英文缩写（e.g./i.e.）、URL、邮箱、代码段、公式。
- 每处改动：段落号 + 原文 + 新文，逐条进 `report.content_guard.punctuation_changes`，并从指纹 diff 的失败判定中排除（属于预期变更）。

## 11. Page Layout Engine

### 11.1 页面设置

- A4 + 模板页边距，逐节应用；不删除/合并任何 `w:sectPr`。
- 多节文档：每节独立设置边距；节的数量与分节符位置在操作前后必须一致（Q16）。

### 11.2 页码

- **检测优先**（`page_number_skip_existing=true` 默认）：扫描全部 footer/header 部件中的 PAGE / NUMPAGES / fldSimple 域；已存在且用户未明确要求改样式 → 跳过并记录。
- 插入：目标节 footer 段落写入 PAGE 域（附录 C-5，`PAGE \* MERGEFORMAT`；**禁止** `\* decimal` ——它不是合法域开关，会导致 WPS/Word 显示"1decimal"），字体按模板（宋体小五 9pt），对齐默认居中。
- **三段式页码**（academic 模板 + 检测到封面时启用）：封面节无页码 → 前置节（目录等）罗马数字 `pgNumType fmt="upperRoman" start="1"` + 后处理 footer instrText 为 `PAGE \* ROMAN \* MERGEFORMAT` → 正文节 `pgNumType fmt="decimal" start="1"` + instrText `PAGE \* arabic \* MERGEFORMAT`。依据：WPS 忽略 `pgNumType` 的 fmt，必须双写；见附录 A。
- 多节普通文档：只给**缺页码**的节补页码，不改任何节的起始编号。

### 11.3 页眉

- V2 保留既有页眉不动；`header_footer` 操作 + `overrides.header_text` 时，给无页眉的节添加简单文字页眉（默认无下划线装饰，底部边框可选）。装饰性页眉设计不在 V2 范围。

### 11.4 孤行控制

- 所有 Heading 1–3：`keepNext` + `keepLines`（§10.3 步骤 5）。
- 图片段落：若下一段是图注 → 图片段落 `keepNext`（图注跟随图片）；表注在表上方 → 表注段落 `keepNext`。

## 12. 空白段落清理（硬排除清单）

"空白"定义：无文字、无域、无图片/图形（w:drawing）、无分页符、无书签的段落。

**硬排除（绝不允许删除）**：

1. 段落 pPr 携带 `w:sectPr`（删了会合并节、毁页眉页码设置）；
2. 含 `w:br w:type="page"` 或 `w:pageBreakBefore`（排版意义的分页）;
3. **位于两个表格之间**（删除后 Word 打开会把两表合并成一个表——竞品均未防此坑）；
4. 含 `w:fldChar`/`w:fldSimple`/`w:instrText` 的段落；
5. ExclusionZones `paragraphs_skip_all` 内（封面、目录区）；
6. 文档首段、任意单元格首段。

**允许动作**：正文中连续 ≥2 个空白段 → 收敛为 1 个（保守策略；不做全删）。每次删除登记段落序号，数量进入 Q02 对账。

## 13. Table Processor

原则：**不修改单元格中的文字内容**（字体字号除外）。

- 宽度适配：可用宽度 = 所在节页宽 − 左右边距。写入三件套：`tblW`（dxa）、`tblGrid` 列定义、每格 `tcW`（dxa）；`tblLayout fixed`。**合并单元格表**（gridSpan/vMerge）：只设置 tblW + tblGrid + tblLayout，跳过逐格 tcW 并记 warning（逐格重算易错）。
- 表头：首行加粗、居中、浅色底纹（模板色）；`w:tblHeader` 跨页重复（`table_repeat_header=true` 默认）。
- 行：`w:cantSplit` 防行内截断；行高超过一页的行跳过并 warning。
- 单元格：垂直居中；水平对齐**保持原样**（数据语义优先）；字体字号统一照 §10.1，跳过缩进。
- 边框：`table_borders` — academic 默认 `threeline`（三线表：顶/底 1.5pt、栏目线 0.75pt、无竖线），formal/clean 默认 `grid`；`auto` = 模板默认。改边框不改内容。
- 宽度溢出检查留给 Q09。

## 14. Image Processor（inline only，边界声明）

- 处理范围：**仅 inline 图片**（`document.inline_shapes`）。浮动图片（wp:anchor）与文本框内图片：只计数、只报告"未处理"（§2.5），绝不移动/删除。
- 超宽收缩：inline 图宽度 > 可用宽度 × `image_max_width_percent/100` 时，按 EMU 等比缩小（宽高同乘系数，禁止只改一边导致变形）；**只缩小不放大**。
- 居中：图片所在段落水平居中（该段跳过首行缩进）。
- 图注跟随：§11.4。
- EMF/WMF 矢量图：允许，只改 extent 尺寸。
- 每次缩放登记：图序号、旧尺寸 EMU → 新尺寸 EMU。

## 15. Caption Detection

- 模式（段落起始）：`^图\s*[\d\-\.．]+`、`^表\s*[\d\-\.．]+`、`^Figure\s*\d+`、`^Table\s*\d+`、`^Fig\.\s*\d+`。
- 归类后：套图注样式（模板 captions 参数，宋体五号居中）；进 ExclusionZones（跳过缩进、跳过标题启发式）。
- 位置语义：表注在表上方 → `keepNext`；图注在图下方 → 由图片段 `keepNext` 保证跟随（§11.4）。
- 图注/表注的编号文字属于内容，**不重排、不更新**。

## 16. TOC

**自动执行条件**（量化，消除 V1 的"结构明确"模糊性）：`operations` 含 `toc` **且**（`heading_count ≥ 3` 且 `h1_count ≥ 1`）。用户显式指定时跳过条件检查但仍需 ≥1 个 Heading。

**插入流程**：

1. 已有 TOC 域 → 只确保 `settings.xml` 含 `<w:updateFields w:val="true"/>`，不插入第二个域。
2. 已有手工文字目录（点线+页码模式段落群）→ 默认**不删除**（内容保护），报告 warning 建议用户手动删除后重跑；`options.toc_replace_manual=true` 时才删除该区域并登记。
3. 插入位置：有前置区（封面/摘要后）→ 前置区末尾；否则文档开头第一个正文 Heading 之前。
4. 构成：①"目录"标题段（直接格式：黑体三号居中，**不用 Heading 1**，避免目录收录自身）→ ②TOC 域 `TOC \o "1-3" \h \z \u`（fldChar begin/separate/end，separate 与 end 之间放占位文本"（打开文档后按 F9 或右键'更新域'生成目录）"）→ ③可选灰色斜体提示行（`toc_refresh_hint`，默认 true）→ ④分页符（`toc_page_break`，默认 true；官方 docx skill 实证：缺分页是目录排版第一失败原因）。
5. `settings.xml` 写入 `updateFields=true`——否则目录打开是空的。
6. 目录区段落全部登记进 `excluded_generated`（护栏 diff 排除区）。

## 17. Content Guard

**这是本 Skill 的核心差异化模块，必须最先实现快照部分（Phase 1）。**

### 17.1 快照（修改前，流水线步骤 1）

- 提取范围：正文段落文本（含超链接内文本，python-docx ≥1.1 的 `paragraph.text` 已覆盖，须自测）+ 全部表格单元格文本（按表序、行序、列序）。
- **明确不进快照**（排版本身会合法修改）：页眉页脚、脚注尾注、文本框、批注。
- 归一化（消指纹误报，借鉴 paper_format_agent）：NFC → 连续空白（含全角空格）折叠为单空格 → strip。

### 17.2 指纹与逐项 diff

- 整体指纹：`sha256("\n".join(归一化段落文本))`；同时保存逐段哈希序列，用于定位 diff。
- 表格指纹：每表 `sha256` 拼接单元格归一化文本。

### 17.3 diff 的预期变更排除

以下差异**不算护栏失败**，但必须逐项登记：

- `excluded_generated`：目录标题/域占位/提示行/分页段落（toc 操作登记）；
- 空白段落删除（Q02 用"删除数对账"校验，而非指纹）；
- `punctuation_changes`（仅当该操作开启）；
- 页码域、页眉文字（不在快照范围）。

### 17.4 完整性计数（不可减少原则）

处理前后对比以下计数，**只允许增加或不变**：`w:hyperlink`、`w:bookmarkStart`、`fldChar/fldSimple`、表格数、inline 图片数、浮动图片数、节数、批注数。任何减少 → 护栏失败（借鉴 Word-Formatter-Pro 的域/书签保护，扩展为通用计数）。

### 17.5 失败策略

- 判定失败 → 删除临时 docx，**不产出输出文件**，报告 `status="guard_failed"`、`exit_code=3`，附 diff 明细（段落号 + before/after 摘要，最多 20 条）。
- `content_changed` 字段如实反映，Skill 不得假装没有变化（V1 原则保留，V2 补上"拒绝落盘"的硬执行）。

## 18. Quality Checker（Q01–Q16）

| ID | 检查 | 通过标准 | 级别 |
|---|---|---|---|
| Q01 | 输出可打开 | python-docx 重新加载成功 | error |
| Q02 | 段落对账 | `后段落数 = 前段落数 − 删除空段数 + 目录新增段数`（±0） | error |
| Q03 | 表格数不变 | 前后相等 | error |
| Q04 | inline 图片数不变 | 前后相等 | error |
| Q05 | 浮动图片/文本框数不变 | 前后相等 | error |
| Q06 | 超链接不减 | after ≥ before | error |
| Q07 | 域/书签不减 | after ≥ before | error |
| Q08 | 标题层级 | headings>0 时 h1≥1；无跳级（级差>1） | warning |
| Q09 | 表格宽度 | tblW ≤ 所在节可用宽度（逐节） | warning |
| Q10 | 残留空段簇 | 不存在连续 >2 空白段 | warning |
| Q11 | 页码 | 请求页码时目标节 footer 存在 PAGE 域 | warning |
| Q12 | 目录 | 请求 TOC 时域存在且 updateFields=true | warning |
| Q13 | 字体覆盖 | 抽样正文 run eastAsia 设置率 ≥95% | warning |
| Q14 | 缩进一致 | 非排除区正文段 firstLineChars=200 占比 ≥90%（clean 模板跳过） | info |
| Q15 | 行距一致 | 主流行距占比 ≥90% | info |
| Q16 | 兼容红线 | instrText 无 `\* decimal`；无空 `<w:pgNumType/>`；节数不变 | warning |

任一 error → 该项进入 `errors[]`，整体状态降级；Q01/Q02/Q03 失败时若护栏亦失败 → 不落盘（§2.4）。

## 19. Report Schema（完整）

```json
{
  "status": "success | partial_success | guard_failed | dry_run | error",
  "exit_code": 0,
  "version": {"skill": "2.0.0", "python": "3.11.6", "python_docx": "1.1.2"},
  "input":  {"file": "input.docx", "size_bytes": 48213, "sha256": "ab12..."},
  "output": {"file": "input_beautified.docx", "size_bytes": 51002, "written": true},
  "config": {
    "style_resolved": "academic",
    "document_type": {"value": "academic", "confidence": 0.8,
                      "evidence": ["关键词'摘要'", "标题'参考文献'", "3 级标题结构"]},
    "operations_requested": "auto",
    "operations_executed": ["page_setup_done", "headings_formatted", "..."],
    "operations_skipped": [{"op": "toc", "reason": "heading_count < 3"}],
    "operations_failed": [],
    "overrides_applied": {"font": {"chinese": "宋体"}}
  },
  "analysis": { "...§6 全部指标..." },
  "zones": {"cover": [0, 5], "front_matter": [6, 12], "toc_existing": null,
            "references": [110, 126], "confidence_notes": []},
  "statistics": {
    "before": {"paragraphs": 126, "tables": 4, "inline_images": 7, "headings": 0, "hyperlinks": 12},
    "after":  {"paragraphs": 122, "tables": 4, "inline_images": 7, "headings": 12, "hyperlinks": 12},
    "reconciliation": {"blank_deleted": 7, "toc_added": 3, "delta_explained": true}
  },
  "operations_detail": [
    {"op": "headings_formatted", "status": "done", "changed": {"promoted": 12, "style_numbering_disabled": ["Heading 1"]}, "warnings": []},
    {"op": "tables_formatted", "status": "done", "changed": {"tables": 4}, "warnings": ["Table 3 含合并单元格，仅做整表宽度适配"]}
  ],
  "content_guard": {
    "enabled": true,
    "fingerprint_before": "e3b0...",
    "fingerprint_after": "e3b0...",
    "content_changed": false,
    "excluded_generated": {"paragraphs": 3, "preview": ["目录", "（打开文档后按 F9...）"]},
    "punctuation_changes": [],
    "integrity": {"hyperlinks": {"before": 12, "after": 12},
                  "bookmarks": {"before": 2, "after": 2},
                  "fields": {"before": 0, "after": 2},
                  "tables": {"before": 4, "after": 4},
                  "inline_images": {"before": 7, "after": 7}},
    "diffs": []
  },
  "quality_checks": [{"id": "Q01", "status": "pass", "detail": "..."}],
  "untouched_boundary": {"floating_images": 1, "textboxes": 0, "equations": 3, "footnotes": 12,
                         "note": "以上对象未处理亦未破坏"},
  "warnings": ["文档包含 5 处未接受修订，建议先在 Word 中接受修订再排版"],
  "errors": [],
  "timing": {"total_ms": 4120, "per_op_ms": {"normalize_font": 380}}
}
```

规则：`status=partial_success` 时 `errors[]` 必须非空且逐项可读；`guard_failed` 时 `output.written=false` 且 diff 非空；报告以 UTF-8 写出（`ensure_ascii=False`）。

## 20. 错误处理与错误码

| 错误码 | 场景 | exit_code | 是否产出文件 |
|---|---|---|---|
| `FILE_NOT_FOUND` | 输入不存在 | 1 | 否 |
| `UNSUPPORTED_FILE` | 非 .docx（.doc/.wps/.pdf/...） | 1 | 否 |
| `FILE_TOO_LARGE` | > 上限（默认 100MB） | 1 | 否 |
| `PASSWORD_PROTECTED` | 加密文档 | 1 | 否 |
| `INVALID_DOCX` | ZIP 损坏 / python-docx 无法加载 | 1 | 否 |
| `EMPTY_DOCUMENT` | 无正文段落 | 1 | 否 |
| `CONFIG_INVALID` | schema 校验失败 / 标点操作未满足门控 | 1 | 否 |
| `OUTPUT_WRITE_FAILED` | 落盘失败（权限/磁盘） | 1 | 否 |
| `GUARD_FAILED` | 内容护栏失败 | 3 | **否（已删除）** |
| `OP_FAILED:*` | 单操作失败（partial_success 前缀错误码） | 2 | 是 |

单个非致命操作失败（如某表合并单元格超复杂）→ 跳过该子任务 + `errors[]` 登记 + 继续，**不得因单表失败丢弃整个输出**（V1 原则保留，落地为 §2.4 机制）。

## 21. CLI 与 Agent 集成

### 21.1 CLI

```text
python main.py <input.docx> [-c config.json] [-o OUTDIR] [--report PATH]
              [--dry-run] [--batch DIR|GLOB] [--print-schema]
```

- stdout：仅 report.json（Agent 解析入口）；stderr：人类日志。
- 退出码：0/2/3/1（§0）。`--print-schema` 输出配置 JSON Schema。
- Windows 注意：所有文件 IO 显式 `encoding="utf-8"`；报告 `ensure_ascii=False`。

### 21.2 --dry-run（Agent 协作的关键）

执行流水线步骤 1–4（加载/分析/结构检测/配置合并），**不修改任何东西**，输出 `status="dry_run"` 报告，额外包含：

- 将执行的模板与合并后参数；
- 标题提升计划（将套样式的高置信段落清单 + 候选清单）；
- 将插入的目录/页码/将执行的每个操作及跳过原因。

Agent 推荐流程：`--dry-run` → 向用户展示计划 → 确认后正式运行 → 转述报告。

### 21.3 SKILL.md 规范

- frontmatter：`name: docbeauty`；`description` 中英双语、含触发场景（论文排版/Word 美化/格式统一/毕业论文格式/beautify or format a Word document without changing its content）与"不适用"声明（用户要改内容时不要调用本 Skill）。撰写时使用 skill-creator 调优触发。
- 正文必须包含：§0 五条默认行为、dry-run 工作流、报告转述要求（`status != success` 必须如实转述 warnings/errors，guard_failed 必须告知用户"未产出文件"）、format_file 分工（Agent 读格式要求文档 → 生成 overrides）。

### 21.4 批量模式

`--batch DIR`：对目录内全部 .docx 独立执行，各自输出 `*_beautified.docx` + `*_report.json` + 汇总 `batch_summary.json`（每文件 status/exit_code）。单文件失败不影响其余。

## 22. 实现安全守则（OOXML 红线）

1. 只通过 python-docx / lxml DOM 修改；**禁止**解包 zip 后对 document.xml 等做正则/字符串替换。
2. 全程操作临时副本；`os.replace` 原子落盘。
3. 不理解的部分不动：未知包部件、自定义 XML、宏（vbaProject）、内容控件（w:sdt）原样保留。
4. 绝不删除：`w:sectPr`、书签、域、批注部件、脚注部件、超链接关系。
5. 依赖锁定：`python-docx>=1.1.0,<2.0`、`lxml>=4.9`、`jsonschema>=4`。Python ≥3.9。
6. 所有 XML 命名空间经 `docx.oxml.ns.qn()` 取得，禁止手写 URI。

## 23. 项目结构

```text
docbeauty/
├── SKILL.md                  # 按 §21.3 撰写
├── README.md
├── requirements.txt
├── main.py                   # CLI 入口（§21.1）
├── config.py                 # schema 校验 + overrides 深合并 + options 门控
├── constants.py              # 中文字号映射、错误码、操作目录、默认 options
├── analyzer.py
├── structure.py              # 区域识别 + 标题/图注/列表检测 + ExclusionZones
├── guard.py                  # 快照 / 指纹 / diff / 完整性计数
├── formatter.py              # 字体 + 段落（§10.1–10.2）
├── headings.py               # 检测打分 + 样式应用算法（§8/§10.3）
├── lists.py
├── tables.py
├── images.py
├── layout.py                 # 页面/页码/页眉/keepNext（§11）
├── toc.py
├── cleanup.py                # 空白段落清理（§12）
├── checker.py                # Q01–Q16
├── report.py
├── utils/ooxml.py            # qn 封装、rFonts/ind/域 构造器（附录 C 的实现）
├── templates/
│   ├── academic.py
│   ├── formal.py
│   └── clean.py
├── schemas/
│   ├── input_schema.json
│   └── output_schema.json
├── examples/
│   ├── configs/              # 典型配置 3–5 个（含 format_file → overrides 示例）
│   └── input/  output/
└── tests/                    # fixture 全部程序化生成（§24）
```

## 24. 测试计划

**Fixture 策略**：全部用 python-docx 程序化构造，不提交二进制样例。必备 fixture：

手工编号标题文档 · 双表夹空段文档 · 含 sectPr 段落文档 · 含超链接文档 · 含浮动图/文本框/OMML 公式文档 · 含修订与批注文档 · 已有 TOC 域文档 · 已有手工目录文档 · 含 Heading 样式绑定多级列表的文档 · 加密文档（跳过用）· 空文档。

**必测矩阵**：

| 测试 | 断言 |
|---|---|
| T-guard-no-false-positive | 插入目录/页码后护栏通过，excluded_generated 登记完整 |
| T-guard-catch | 人工篡改一段文字 → guard_failed 且无输出文件 |
| T-double-numbering | Heading 样式挂 numPr + 手工编号 → 不产生双重编号 |
| T-table-merge | 两表之间空段在清理后仍存在 |
| T-sectpr | 含 sectPr 的空段不被删除 |
| T-hyperlink | 超链接文本被格式化且计数不减 |
| T-emphasis | 粗/斜/下划线/颜色在字体统一后保留 |
| T-floating | 浮动图计数不变、位置不动 |
| T-equation | OMML 公式渲染不受影响（XML 级比对） |
| T-idempotent | 对输出再跑一次，各操作 changed=0 |
| T-manual-toc | 手工目录被识别并警告，未被删除 |
| T-tracking | 修订/批注触发 warning |
| T-reconcile | Q02 对账公式精确成立 |
| T-wps-proxy | 输出能被 LibreOffice 打开并转 PDF（可选 CI） |

## 25. 开发阶段

| Phase | 内容 | 交付判据 |
|---|---|---|
| 0 | 骨架 + config schema + CLI + constants + report 骨架 | --print-schema 可用 |
| 1 | 原子读写 + analyzer + **guard 快照** + dry-run | T-reconcile 前置就绪；dry-run 报告正确 |
| 2 | 双轨样式 + 字体 + 段落 + 标题 + 列表 | T-double-numbering / T-emphasis 通过 |
| 3 | 页面 + 页码（多节/三段式）+ 页眉 | Q11/Q16 通过 |
| 4 | 表格 + 图片 + 图注 | Q09 通过；T-table-merge/T-floating 通过 |
| 5 | TOC + cleanup + checker + **guard 强制执行** | T-guard-* 全通过 |
| 6 | SKILL.md + Agent 协议 + 批量 + 幂等测试 + 文档 | T-idempotent 通过；Agent 端到端演练 |
| 7（可选） | format_file 辅助示例、.doc/.wps COM 转换、punctuation_normalize、Word COM 渲染级视觉验收、样式复刻（参考文档模式） | 每项独立验收 |

## 26. V2 明确不做的事

V1 的清单全部保留（不改写正文/AI 润色/自动生成内容/PDF/PPT/Excel/OCR/云端/账号等），新增：

- 不做装饰性封面设计与封面配方库（那是生成式 Skill 的领域，见官方 docx skill）；
- 不做样式复刻（参考文档 → 新文档的排版克隆，Phase 7 之后再议）;
- 不做参考文献格式化/GB/T 7714 校验（内容处理，独立工具域）;
- 不承诺跨平台字体渲染一致（字体名写入文档，渲染取决于打开方安装的字体，报告不为此负责）。

## 27. 成功标准（可测量）

1. **基础**：样例文档集 100% 正确读入/写出；输出可被 Word/WPS/LibreOffice 打开。
2. **自动化**：§8–§16 每个操作在对应 fixture 上有通过测试；`document_type` 自动检测在 3 类样例上给出正确值与证据。
3. **安全**：T-guard-catch 证明"篡改必拦、拦截必无文件"；T-guard-no-false-positive 证明目录/页码不误报；Q01–Q07 全部 error 级检查有测试覆盖。
4. **Agent**：`--dry-run` → 运行 → 解析 report.json 的端到端演练可复现；Agent 无需阅读源码即可完成调用。
5. **工程**：T-idempotent 通过；批量模式单文件失败不传染。

## 28. V1 → V2 变更追溯表

| # | 主题 | V1 | V2 | 依据 |
|---|---|---|---|---|
| 1 | 字体实现 | "统一中英数字字体"（无细节） | eastAsia/ascii/hAnsi/主题引用四件套 + 强调保留清单 | 审查 T1；竞品 Word-Formatter-Pro/docx-format-skill |
| 2 | 缩进实现 | "首行缩进 2 字符" | `firstLineChars=200` + twips 回退 | 审查 T1 |
| 3 | 样式策略 | 未提及 | 双轨制 D1 | 审查遗漏 1 |
| 4 | 双重编号 | 未提及 | §10.3 三步防线 | 审查 T3 |
| 5 | 空段清理 | "可删除明显无意义空段" | 硬排除清单（sectPr/分页/两表之间/域/封面） | 审查 T4 |
| 6 | 目录 | "原生 TOC Field" | + updateFields、占位、提示行、分页、已存域/手工目录处理 | 审查 T5；官方 docx skill |
| 7 | 内容保护 | "检测到变化报 warning" | 指纹 + 排除区 + 完整性计数 + **失败不落盘** | 审查 T6；竞品 paper_format_agent |
| 8 | 图片 | "图片处理" | inline 限定 + 浮动/文本框/公式边界声明 | 审查 T7；竞品全员空白 |
| 9 | 表格 | "宽度适配" | tblW+tblGrid+tcW 三件套、合并单元格降级策略、三线表选项 | 审查 T8；docx-skill-4-cn-paper |
| 10 | 参数覆盖 | §10 说"允许覆盖"但协议无字段 | `overrides` 正式入 schema + 深合并规则 | 审查遗漏 3 |
| 11 | 调用接口 | 未定义 | CLI/退出码/stdout 契约 + dry-run + 批量 | 审查遗漏 4 |
| 12 | 标题检测 | 三级示例 | 四级编号模板库 + 打分制 + 候选不落盘机制 | 竞品 Word-Formatter-Pro/WordFormatter |
| 13 | 标点规范化 | 无（原则禁止） | 显式 opt-in 操作 + 门控 + 逐条登记 | 竞品 document-format-skills |
| 14 | 页码 | "页脚居中" | 已有检测跳过 + 多节策略 + 三段式 + WPS 红线 | 审查遗漏 6；官方 docx skill WPS 笔记 |
| 15 | 列表 | 识别但无操作 | 列表策略（跳缩进、不重排编号） | 审查遗漏 2 |
| 16 | meeting 类型 | 检测到但无模板 | 映射 formal，检测证据表量化 | 审查不一致 2 |
| 17 | 质检 | 10 条模糊检查 | Q01–Q16 带判据与级别 | 审查不一致 1 |
| 18 | 报告 | 基础字段 | + errors[]、前后统计对账、护栏明细、边界声明、配置回显 | 审查不一致 3 |
| 19 | 修订/批注 | 未提及 | 检测 + 强制警告 | 审查遗漏 6；docx-skill-4-cn-paper |
| 20 | 排除区 | 散落各节 | ExclusionZones 一级公民 D2 | 审查遗漏 5 |
| 21 | 幂等/原子写 | 未提及 | D3/D4/D7 + T-idempotent | 审查小项 |
| 22 | 封面 | 未提及 | 识别 + 整区排除（不做封面设计） | 审查遗漏 5；竞品通病 |

---

## 附录 A：WPS/Word 兼容性笔记（实现时必须遵守）

1. 页码域格式开关用 `\* arabic` / `\* ROMAN`，**禁止 `\* decimal`**（非法开关，显示为"1decimal"）。
2. WPS 忽略 `pgNumType` 的 fmt → 三段式页码必须同时后处理 footer instrText。
3. 清除封面节可能残留的空 `<w:pgNumType/>`（部分生成器会写出，干扰 WPS）。
4. `w:firstLineChars` 优先于 `w:firstLine`（Word/WPS 均支持），保留 twips 值作为其他渲染器回退。
5. 禁用样式编号用 `numId=0`，不要删除共享样式上的 numPr 属性（可能被其他样式引用）。

## 附录 B：中文字号映射表

| 字号 | pt | 字号 | pt |
|---|---|---|---|
| 初号 | 42 | 四号 | 14 |
| 小初 | 36 | 小四 | 12 |
| 一号 | 26 | 五号 | 10.5 |
| 小一 | 24 | 小五 | 9 |
| 二号 | 22 | 六号 | 7.5 |
| 小二 | 18 | 小六 | 6.5 |
| 三号 | 16 | 七号 | 5.5 |
| 小三 | 15 | 八号 | 5 |

## 附录 C：关键 OOXML 片段速查

```xml
<!-- C-1 字体（四件套 + 清主题引用） -->
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>

<!-- C-2 首行缩进 2 字符（Chars 优先，twips 回退=2×字号pt×20） -->
<w:ind w:firstLineChars="200" w:firstLine="480"/>

<!-- C-3 1.5 倍行距 -->
<w:spacing w:line="360" w:lineRule="auto"/>

<!-- C-4 标题孤行控制 -->
<w:pPr><w:keepNext/><w:keepLines/></w:pPr>

<!-- C-5 页码域 -->
<w:fldSimple w:instr=" PAGE \* MERGEFORMAT "><w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>

<!-- C-6 目录域（三段 fldChar，separate 与 end 之间放占位文本） -->
<!-- instrText: TOC \o "1-3" \h \z \u -->

<!-- C-7 settings.xml 打开时更新域（目录/页码必需） -->
<w:updateFields w:val="true"/>

<!-- C-8 表头跨页重复 + 行防拆 -->
<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>

<!-- C-9 禁用样式继承编号（防双重编号） -->
<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>

<!-- C-10 正文节页码从 1 开始 -->
<w:pgNumType w:fmt="decimal" w:start="1"/>
```

---

**设计哲学（与 V1 一致，V2 补足下半句）**：

> AI 的理解能力 + 程序的确定性 = 可靠的文档自动化；
> 而"可靠"必须被证明——所以 DocBeauty 的每一次输出，都附带一份内容未被改变的自证报告。
