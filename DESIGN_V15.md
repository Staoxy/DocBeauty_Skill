# DocBeauty V1.5 设计方案（合并修订版）

- **版本**：V1.5（基于 V1.5 原始方案 + 现有代码库对照审查合并而成）
- **定位**：Agent 可调用的智能 Word 文档排版与模板引擎
- **基线**：`v1.5-dev` 分支（V1 最终态 `00de506` + 数学式防护 `a2ea3e4`），30 项测试全绿
- **阅读指引**：先读 §0 对照表（明确"已有/新增"边界），再读 §1–§3（定位与模式），实现者精读 §4–§16，§18 是执行计划，§21 是完成定义。

---

## 0. 现有能力对照表（本方案的第一原则：不重复造轮子）

V1.5 原方案的部分"新增能力"在当前代码库中已经实现。本方案的一切规划以此为前提：

| 原方案条目 | 现有实现 | V1.5 处置 |
|---|---|---|
| Content Guard（快照/diff/失败停写） | `guard.py`（指纹、排除区、完整性计数、拒落盘） | 保留，扩展模板模式的排除区登记 |
| Dry Run | `--dry-run`（含标题提升计划） | 升级为 `plan` 子命令，输出 formatting_spec.json |
| 视觉验证（Word→PDF/域更新/目录/页码） | `render_check.py`（真 Word COM） | 补 PDF→PNG 栅格化 + 客观版面指标 |
| PASS/WARNING/ERROR 质检 | `checker.py` Q01–Q16 | 保留，增加模板符合性检查 |
| 批量处理 | `--batch`（单文件失败不传染） | 保留，支持 `--template` 全局模板 |
| 结构识别（中文编号模板/混合编号） | `structure.py`（含数学式防护） | 保留，复用于**模板侧**结构识别 |
| Style vs 直接格式（目标文档侧） | 双轨制 D1（`formatter.py`/`headings.py`） | 保留；**模板侧**是新增工作 |
| 诊断 | `diagnose.py`（分级 issues） | 升级为 `analyze` 子命令；`audit` 是它的"目标 vs 模板"扩展 |
| CLI / JSON API / SKILL.md | `main.py` 参数式 + SKILL.md | 重构为子命令式（保持兼容） |
| 文档类型自动检测 | `analyzer.detect_document_type` | 保留 |
| 旧格式转换 | `legacy_convert.py`（Word COM） | 保留 |
| **Style Analyzer / Template Analyzer / Role Mapper / Rule Resolver / Formatting Spec / Audit / 视觉栅格化** | **无** | ⭐ **V1.5 的真实增量** |

---

## 1. 产品定位

DocBeauty 不再只是"自动统一字体、行距和标题的工具"，而是：

> **一个可以被 AI Agent 调用、能够理解用户排版要求、理解 Word 模板与参考风格，并自动完成专业文档排版与自证验证的确定性 Word 文档引擎。**

核心竞争力（原方案 §26，保留）：

```text
理解用户要求 + 理解目标文档 + 理解模板 + 理解参考风格
        ↓ 生成结构化排版规则
        ↓ 确定性执行
        ↓ 内容保护
        ↓ 多层验证
```

---

## 2. 核心理念（三条，第 2.2/2.3 条为本次修订新增）

### 2.1 AI 负责理解，代码负责执行（保留）

LLM/Agent：理解自然语言 → 生成结构化 Intent → 生成/确认 Formatting Spec。
DocBeauty：校验 Intent → 分析模板/参考 → 解析规则 → 确定性执行 → 护栏 → 验证。
**禁止** LLM 直接操作 Word XML；**禁止** Skill 内部调用 LLM。

### 2.2 Skill 测量，Agent 诠释（新增）

一切主观/语义判断（"学术感"、"正式"、"主色调"、"风格抽象"）由 Agent 完成；
Skill 只输出**客观测量值**（实测字体、字号、间距、颜色 RGB、结构信号）。
代码里不出现"审美映射表"。这保证同输入同输出，也保证 Reference 模式不引入主观性。

### 2.3 安全不变量高于一切规则优先级（新增）

Rule Resolver 的优先级只对**排版参数**生效。内容保护类硬约束（不改正文、
不重排编号、节数不减、超链接/表格/图片只增不减）**凌驾于所有优先级之上**，
任何规则来源都不能覆盖（与 V1 的 guard 失败拒落盘一致）。

---

## 3. 四种输入模式

统一入口：`mode: "auto" | "prompt" | "template" | "reference"`。

### MODE 1 — AUTO（保留现有默认行为）

无附加输入，走现有流水线：分析 → 类型检测 → 内置模板（academic/formal/clean）→ 执行。

### MODE 2 — PROMPT（自然语言要求）

**分工（修订）**：Agent 把用户的话翻译成受控的 Intent JSON（§4），Skill 的
`intent.py` 只做**校验与归一化**（不是 NL 解析模块）。示例：

用户："把这份实验报告排得正式一点，标题层级清晰，表格不要太拥挤，不要修改任何内容。"

Agent 产出：

```json
{
  "mode": "prompt",
  "intent": {
    "style": "formal",
    "document_type": "experiment",
    "content_protection": true,
    "preferences": {
      "heading_emphasis": "strong",
      "table_density": "comfortable"
    }
  },
  "prompt_raw": "把这份实验报告排得正式一点……"
}
```

- `preferences` 是受控词表（§4 定义），Agent 只能填枚举值；
- `prompt_raw` 仅用于报告回显，Engine 不读它；
- `preferences` 无法直接驱动 Engine，由 Rule Resolver 映射为具体参数
  （如 `table_density: comfortable` → 单元格边距加大、行距 1.3），映射表是
  确定性代码，进报告回显。

### MODE 3 — TEMPLATE（主打能力）

输入：`target.docx` + `template.docx`。

```text
Template Analyzer（§6，双路径）
 ↓ Template Specification（§7）
Target Analyzer（现有 analyzer + structure 复用）
 ↓ Role Mapping（§9）
 ↓ Rule Resolver（§10）
 ↓ Formatting Spec（§11）
 ↓ Engine 确定性执行
 ↓ Guard / Checker / Render
```

**双路径（修订，关键补充）**：真实的"学校模板"有两种形态，都必须支持：

- **规范型**：styles.xml 中有真正的样式定义（"正文"、"标题 1"…）→
  Style Analyzer 直接提取；
- **样例型**：模板本身就是一篇排好版的示范论文，格式全在直接格式化里 →
  **复用现有 `structure.detect()` 对模板做结构识别**，从实际段落提取有效
  格式（居中大字号 → Title、编号模式 → 层级、图注模式 → Caption…），
  等效于"对模板跑一遍 analyzer 再提取各角色的实际格式"。

路径判定：模板中 Heading 样式覆盖率 + 样式使用直方图；低置信时两路都跑并
合并（样式定义优先，直接格式补缺），结果标注 `template_kind` 与置信度。

### MODE 4 — REFERENCE（风格参考）

输入：`target.docx` + `reference.docx`。**Reference ≠ Template**：

- Template：尽可能严格复制规则（含页面、页眉页脚、字号精确值）；
- Reference：学习视觉风格与结构思想，允许 Agent 在客观测量之上做抽象。

按 §2.2 分工：**Skill 的 style_analyzer 只输出客观测量**（各角色实测字体/
字号/行距/缩进/颜色 RGB/表格样式信号），构成 Style Profile（§8）；
"这是学术感、主色是深蓝、强调色是灰"这类语义抽象由 Agent 在生成最终
Formatting Spec 时完成——Agent 读 Style Profile（JSON，token 友好），
改写其中的值，Skill 校验后执行。

---

## 4. Intent 规范（受控词表）

`intent` 对象经 jsonschema 校验，未知键 → warning（不失败）；枚举外值 → warning + 回退：

```json
{
  "style": "academic | formal | clean",
  "document_type": "auto | academic | experiment | report | meeting | generic",
  "content_protection": true,
  "preferences": {
    "heading_emphasis": "default | strong | subtle",
    "table_density": "compact | default | comfortable",
    "body_consistency": true,
    "caption_align": "center | left",
    "line_density": "compact | default | loose"
  },
  "notes": "Agent 的自由文本备注（仅回显）"
}
```

`preferences → 参数` 的映射表是确定性代码（`intent.py` 内），每次映射进报告
`config.intent_resolved`。Rule Resolver 中 Intent 的优先级见 §10。

---

## 5. Style Analyzer（新增 `style_analyzer.py`）

输入任意 docx，输出其 Style System 的客观画像：

```json
{
  "doc_defaults": {"eastAsia": "宋体", "ascii": "Times New Roman", "size_pt": 12},
  "styles": {
    "Normal":   {"eastAsia": "宋体", "ascii": "Times New Roman", "size_pt": 12,
                 "line": 1.5, "first_line_chars": 2, "align": "both",
                 "before_pt": 0, "after_pt": 0},
    "Heading 1": {"eastAsia": "黑体", "size_pt": 16, "align": "center",
                  "before_pt": 24, "after_pt": 18, "outline_lvl": 0,
                  "numbering": null},
    "Caption":  {"eastAsia": "宋体", "size_pt": 10.5, "align": "center"}
  },
  "style_usage": {"Normal": 90, "Heading 1": 3},
  "styles_with_numPr": [],
  "theme_fonts": {"majorEastAsia": "等线 Light", "minorEastAsia": "等线"}
}
```

规则：
- 中英文字体分离读取（`w:eastAsia` / `w:ascii` / `w:cs`），含主题字体解析；
- **声明与实际的偏差**单独输出：遍历段落统计"样式声称 X、直接格式覆盖为 Y"
  的比例（V1 已有字号直方图基础，扩展到字体/缩进/行距三个维度）；
- 编号定义（numPr → numId → abstractNum）只读快照，供策略决策（§12.1）。

---

## 6. Template Analyzer（新增 `template_analyzer.py`）

输出 Template Specification（§7）。提取维度（原方案 §4 保留）：

- 页面：size / orientation / margins / 节数与各节差异；
- 页眉页脚：有无、文字内容、页码域位置与格式（含三段式检测）；
- 字体：三路字体 + 字号（各角色）；
- 样式：Style Analyzer 全量输出；
- 编号：标题编号模式（复用 `structure.HEADING_PATTERNS` 检测）与 numPr 定义快照；
- 目录：TOC 域参数（`\\o "1-3"` 等）；
- 表格：边框模式（三线/网格）、表头样式信号；
- 图注：格式与位置（上/下）。

**只提取，不执行**；Engine 不支持的维度（如分栏、奇偶页不同页眉）照常提取，
但标进 `unsupported[]`，执行阶段转为 warning（沿用"边界诚实"原则）。

---

## 7. Template Specification（中间层，schema：`schemas/template_spec.json`）

```json
{
  "template_kind": "regular | sample",
  "confidence": 0.85,
  "page": {"size": "A4", "margins_cm": {"top": 2.5, "bottom": 2.5, "left": 3.0, "right": 2.5}},
  "sections": [{"index": 0, "page_number": null}, {"index": 1, "page_number": {"format": "decimal", "start": 1}}],
  "header_footer": {"header_text": "XX 大学课程论文", "footer_page_number": {"position": "footer", "align": "center", "size_pt": 9}},
  "roles": {
    "body":    {"eastAsia": "宋体", "ascii": "Times New Roman", "size_pt": 12, "line": 1.5, "first_line_chars": 2, "align": "both"},
    "h1":      {"eastAsia": "黑体", "size_pt": 16, "align": "center"},
    "h2":      {}, "h3": {},
    "caption": {"eastAsia": "宋体", "size_pt": 10.5, "align": "center"},
    "toc_title": {"eastAsia": "黑体", "size_pt": 16}
  },
  "tables": {"borders": "threeline", "header_bold": true},
  "numbering": {"detected_pattern": "cn_gov", "num_definitions": []},
  "toc": {"levels": "1-3"},
  "unsupported": ["odd_even_header"]
}
```

`roles` 是 Role Mapping 的目标词表（body/h1/h2/h3/caption/title/toc_title/
quote/list），**与 Engine 的 effective 参数键一一对应**——这是"Formatting
Spec 是 overrides 超集"的基础（§11）。

---

## 8. Reference Style Profile（`style_analyzer.py` 的输出复用）

对 reference.docx 跑 Style Analyzer + 结构检测，输出客观 Style Profile：

```json
{
  "measured": {
    "body": {"eastAsia": "宋体", "size_pt": 12, "line": 1.5, "first_line_chars": 2,
             "align": "both", "color_rgb": "000000"},
    "h1":   {"eastAsia": "黑体", "size_pt": 15, "align": "left", "color_rgb": "1F4E79"},
    "table_signal": {"border": "grid", "header_shade": "DEEBF7"}
  },
  "structure_signals": {"heading_depth": 3, "caption_position": "below", "numbering": "numeric"}
}
```

不含 tone/density 等语义词。Agent 读取后改写为自己的 Formatting Spec 提案
（允许把 1F4E79 抽象为"主色"再应用到不同值），Skill 校验数值合法性后执行。

---

## 9. Role Mapping（新增 `role_mapper.py`）

模板角色 → 目标角色，三级匹配策略（确定性、可回显）：

1. **精确名**：样式名完全一致（"正文" ↔ "正文"，"Heading 1" ↔ "Heading 1"）；
2. **内置同义词表**："一级标题/章标题/chapter"→ h1；"图题/图注/figure caption"
   → caption；"题目/标题/title" → title（词表进代码常量，可增量维护）；
3. **启发式**：outlineLvl → 层级；字号 + 加粗 + 居中信号 → title/h1（复用
   V1 的标题打分逻辑作用于模板）。

输出 `role_map: {"target_role": "template_role"}`，整体进报告。
匹配失败的角色：回退目标文档现状 + warning（不猜）。

---

## 10. Rule Resolver（新增 `rule_resolver.py`）

解析顺序（修订版——覆盖式深合并，高优先级覆盖低优先级）：

```text
不变量层（不可协商，见 §2.3）
P1 用户显式：CLI overrides（最精确）+ Intent preferences（经映射表）
P2 模板规则（Template Spec.roles 等）
P3 参考文档（Style Profile，经 Agent 提案）
P4 目标文档现状（已有的正确格式尽量保留：已有页码/目录/节结构）
P5 DocBeauty 内置默认（academic/formal/clean 模板）
```

- 每个最终参数记录来源（`spec.sources: {"font.size_pt": "P2:template"}`），
  报告可回溯"这个值是谁定的"；
- 同为 P1 时：overrides（结构化、精确）> intent preferences（语义、模糊）；
- `P4 目标现状`只对"模板中不存在的维度"生效（模板没规定页眉 → 目标已有的
  页眉保留），模板规定了的维度按 P2 覆盖；
- 冲突示例（原方案 §9 保留）：模板正文五号 + 用户要求小四 → 最终小四（P1 > P2）。

---

## 11. Formatting Spec（schema：`schemas/formatting_spec.json`）

**定义为现有 `overrides` 的超集**，Engine 新增"spec 输入"入口，执行路径不变：

```json
{
  "mode": "template",
  "params": {
    "page": {}, "font": {}, "paragraph": {}, "headings": {"h1": {}, "h2": {}, "h3": {}},
    "captions": {}, "tables": {}, "page_number": {}, "toc": {}, "header_footer": {}
  },
  "role_map": {"body": "正文", "h1": "标题 1"},
  "sources": {"headings.h1.size_pt": "P2:template"},
  "policy": {
    "numbering": "preserve_target",
    "sections": "preserve_target",
    "existing_page_number": "skip",
    "existing_toc": "update_only"
  },
  "unsupported": [],
  "warnings": []
}
```

`params` 的键与现有 effective config 完全同名——spec_builder 的实现就是
"Rule Resolver 输出 → 现有 build_effective 的深合并输入"。Engine 改动量最小化。

---

## 12. 关键策略决策（原方案缺失，本方案补充）

### 12.1 编号策略：`preserve_target`（保守）

模板的编号定义（numPr/abstractNum）**不迁移**到目标文档。模板模式下只套
编号的"视觉外观"（字体字号段距），编号文字/自动编号保持目标文档现状。
原因：内容保护原则禁止重排编号；编号定义迁移极易产生双重编号（V1 已有三重
防线，不应再引入这条高风险路径）。报告注明"编号体系未迁移"。

### 12.2 节结构不兼容：降级 + 警告

模板为三节（封面/目录/正文）而目标为单节时：三段式页码、分节页眉**不套用**，
标记 `sections: incompatible` + warning，页面参数（边距/纸张）仍可应用。
不自动为目标"造节"——那会移动内容，违反不变量层。

### 12.3 已有页码/目录：沿用 V1 策略

模板模式下同样 `page_number_skip_existing`、已有 TOC 域只补 updateFields、
手工文字目录默认保留。模板的页眉**文字**可以套用（P2），但如果目标已有页眉
文字则默认保留 + warning（`policy.header_footer: "preserve_if_present"`）。

### 12.4 版本命名统一

- 代码版本：`constants.SKILL_VERSION` 从 `"2.0.0"` 改为 `"1.5.0"`，此后按
  1.x 演进；
- 文档：DESIGN.md = V1 原始方案；DESIGN_V2.md = V1 的实施设计（历史文档，
  保留不改名）；本文件 DESIGN_V15.md = V1.5 设计。对外口径统一为 V1.5。

---

## 13. Audit 模式（新增 `audit.py`）

`audit` = 现有 `diagnose` 的"目标 vs 模板"扩展，**只读不改**：

```text
docbeauty audit target.docx --template template.docx
```

```json
{
  "result": "ATTENTION",
  "checks": {
    "page_size":       {"target": "Letter", "template": "A4", "match": false},
    "margins":         {"match": false, "detail": "left 3.17cm vs 3.0cm"},
    "body_font":       {"target": "宋体", "template": "仿宋", "match": false},
    "heading_h1":      {"match": true},
    "table_borders":   {"target": "grid", "template": "threeline", "match": false},
    "page_number":     {"match": true},
    "header_footer":   {"match": true}
  },
  "fixable_by_apply": true
}
```

这是对 Agent 最友好的模式：先 audit（✓/✗ 差异表）→ Agent 向用户展示 →
`apply` 修复 → `verify` 复核。完整工作流：Analyze → Audit → Plan → Apply → Verify。

---

## 14. 视觉验证补全（`render_check.py` 扩展）

现有：真 Word COM 更新域/验目录/数页数/导出 PDF。V1.5 补最后一段：

```text
PDF →（PyMuPDF 栅格化，纯 pip 依赖）→ PNG（每页）
```

客观指标（代码计算，不做审美判断——与"PASS/WARNING/ERROR 不做 Beauty Score"一致）：

| 指标 | 方法 | 判定 |
|---|---|---|
| 空白页 | 页面墨水覆盖率 < 0.5% | ERROR |
| 大面积空白 | 正文页内容 bbox 高度占页面 < 30% | WARNING |
| 内容出界 | 内容 bbox 超出页边距框 | ERROR |
| 标题孤行 | 标题文本位于页面底部 10% 且下一页起始为正文（文本层判断） | WARNING |
| 页数突变 | 字段更新前后页数差异 > 20% | WARNING |

PNG 文件路径写入报告 `render_check.pages[]`，**由 Agent 目检**（Agent 有视觉
能力，Skill 不做图像理解）。这个"代码算指标 + Agent 看图"的分工与 §2.2 一致。

---

## 15. CLI 子命令化（保持向后兼容）

```text
docbeauty analyze  <target>                        # 诊断（= 现 diagnose task）
docbeauty audit    <target> --template T.docx      # 目标 vs 模板差异（只读）
docbeauty plan     <target> [--template|--reference|--intent] # 生成 formatting_spec.json，不执行
docbeauty apply    <target> [...]                  # 执行（= 现在的默认行为）
docbeauty verify   <output.docx>                   # 独立质检（checker + guard 复核）
docbeauty render   <docx> [--pdf]                  # 渲染验证（= 现 --render-check）
```

- **无子命令 = apply**：现有调用 `main.py input.docx -o out` 完全兼容；
- `plan` 输出 spec + 现有 dry-run 计划的合并视图；
- JSON API（`-c config.json`）同步扩展：

```json
{
  "mode": "template",
  "target": "report.docx",
  "template": "template.docx",
  "reference": null,
  "intent": {"style": "formal", "preferences": {}},
  "content_protection": true,
  "dry_run": false,
  "overrides": {}
}
```

（`prompt_raw` 由 Agent 侧保存，Engine 配置里只收结构化 `intent`。）

输出 report.json 增加：`template_spec`、`role_map`、`formatting_spec`（含
sources）、`audit`（audit 模式时）、`render_check.pages`。

---

## 16. 项目结构（不做包重组）

**决定：V1.5 放弃 core/intelligence/safety/verification/cli 分包重构。**
理由：当前 20 余个平铺模块 + 30 项测试运转良好；重构 churn 会污染 V1.5 的
功能 diff、增加回归面，且无功能收益。V1.5 仅新增文件：

```text
DocBeauty_Skill/
├── SKILL.md  README.md  main.py（子命令化）  config.py  constants.py
├── analyzer.py  structure.py  formatter.py  headings.py  lists.py
├── tables.py  images.py  layout.py  toc.py  cleanup.py  punctuation.py
├── guard.py  checker.py  report.py  diagnose.py
├── legacy_convert.py  render_check.py
├── intent.py             # 新增：Intent 校验/归一化 + preferences 映射表
├── style_analyzer.py     # 新增：Style System 客观画像（§5）
├── template_analyzer.py  # 新增：双路径模板分析（§6）
├── role_mapper.py        # 新增：三级角色匹配（§9）
├── rule_resolver.py      # 新增：优先级解析（§10）
├── spec_builder.py       # 新增：Resolver 输出 → Formatting Spec（§11）
├── audit.py              # 新增：目标 vs 模板差异（§13）
├── visual_check.py       # 新增：PDF→PNG 栅格化 + 客观指标（§14）
├── schemas/              # + template_spec / formatting_spec / intent JSON Schema
├── templates/            # 保留三套内置参数
└── tests/                # 保留现有 30 项 + 新增 V1.5 测试
```

依赖新增：`PyMuPDF`（可选，视觉栅格化；无此依赖时 render_check 降级为现有
PDF 级检查）。`pywin32` 维持可选（Windows + Word）。

---

## 17. 开发阶段（按现状重排）

| Phase | 内容 | 交付判据 |
|---|---|---|
| 0 | CLI 子命令化（无子命令=apply 兼容）+ SKILL_VERSION→1.5.0 + JSON API 扩展 mode/template/reference/intent 字段 | 现有 30 项测试不改一行全绿 |
| 1 | style_analyzer + template_analyzer（规范型/样例型双路径）+ template_spec schema | 对规范型与样例型模板各产出正确 spec；单测覆盖 |
| 2 | role_mapper + rule_resolver + spec_builder + Engine `--template` 模式 | §21 验收测试第 ③–⑦ 步可跑通 |
| 3 | intent.py（校验 + preferences 映射）+ prompt 模式接线 | MODE 2 示例端到端通过 |
| 4 | audit.py + plan 输出 formatting_spec.json | audit 差异表与 spec 回显正确 |
| 5 | visual_check（PyMuPDF 栅格化 + 5 项指标）| 空白页/出界样例被正确标记 |
| 6 | SKILL.md 四模式触发更新 + §21 十三步验收固化为集成测试 | 验收测试全绿 |

每 Phase 结束跑全量回归；`reference` 模式的"Agent 提案"链路在 Phase 2 的
resolver 之上自然成立（Agent 读 Style Profile → 产出 spec → `apply --spec`），
不单列 Phase。

---

## 18. V1.5 不做的事（原方案 §23 保留 + 补充）

原清单全部保留（内容生成/改写/润色/OCR/PDF→Word/Excel/PPT 转换/云端/在线编辑等）。新增：

- × 包目录重构（core/intelligence/...，推迟到 V2）
- × docxtpl 式占位符模板填充（那是内容生成，属另一产品形态）
- × 代码内语义/审美映射（tone、配色抽象——归 Agent）
- × 编号定义迁移（§12.1）
- × 自动为目标文档造节（§12.2）
- × Beauty Score（§17 精神保留）

V2 展望（原方案 §24 保留）：模板市场、Style 市场、APA/MLA/IEEE 等规范包、跨文档 Style Transfer。

---

## 19. 完成定义（原方案 §30 十三步验收，原样保留）

输入：`我的论文.docx` + `学校论文模板.docx`，要求"按照学校模板排版，不要修改正文内容"。

```text
① 读取论文            → 已有（V1）
② 识别论文结构        → 已有（structure.py）
③ 分析学校模板        → Phase 1
④ 建立 Role Map       → Phase 2
⑤ 生成 Formatting Spec → Phase 2
⑥ 生成修改计划        → Phase 4（plan）
⑦ 执行排版            → 已有（Engine）+ Phase 2 spec 入口
⑧ 检查内容是否变化    → 已有（guard）
⑨ 检查 DOCX 是否损坏  → 已有（checker Q01）
⑩ 渲染 PDF           → 已有（render_check）
⑪ 检查视觉布局        → Phase 5
⑫ 输出最终 Word       → 已有
⑬ 输出 report.json    → 已有 + 扩展字段
```

最终产物：`formatted.docx` + `formatting_spec.json` + `report.json`，
报告告知 Agent：Status: PASS / Content Changed: NO / Formatting / Structure /
Integrity / Visual 全 PASS。此验收固化为 `tests/test_acceptance_v15.py`。

---

## 20. 原方案采纳/修订追溯表

| # | 原方案条目 | 处置 | 理由 |
|---|---|---|---|
| 1 | 四种输入模式 | ✅ 采纳，MODE 2/4 修订分工 | NL 解析归 Agent（§2.1 自洽）；测量/诠释分离（§2.2） |
| 2 | Template Analyzer | ✅ 采纳 + 双路径补充 | 真实学校模板多为样例型（§6） |
| 3 | Style Analyzer | ✅ 采纳 | 含声明/实际偏差统计 |
| 4 | Template/Formatting Spec 中间层 | ✅ 采纳，Spec = overrides 超集 | 复用 Engine 现有参数路径，最小改动 |
| 5 | Rule Resolver 五级优先级 | ✅ 采纳 + 不变量层 + overrides 并入 P1 | 安全约束不可被规则覆盖（§2.3） |
| 6 | Role Mapping | ✅ 采纳，定义三级匹配 | 需确定性、可回显、失败不猜 |
| 7 | Reference 抽象 | ⚠ 修订：Skill 只测量，Agent 诠释 | 代码内主观映射违反确定性（§2.2） |
| 8 | Intent Analyzer 解析 NL | ⚠ 修订为校验/归一化模块 | NL→JSON 是 Agent 职责（§2.1） |
| 9 | Dry Run / Content Guard / Batch / 质检分级 / 渲染验证 | ✅ 已存在，按对照表处置 | §0 对照表 |
| 10 | 包目录重构 | ❌ 推迟到 V2 | churn > 收益（§16） |
| 11 | 编号/节结构策略 | ➕ 新增 §12.1/12.2 | 原方案缺失的关键决策 |
| 12 | 版本命名 | ➕ 统一为 1.5.0 系列（§12.4） | 消除 2.0.0/V2 文档命名冲突 |
| 13 | 十三步验收 | ✅ 原样保留为完成定义 | §19，8/13 已具备 |
| 14 | 不做清单 | ✅ 保留 + 6 条补充 | §18 |
| 15 | §27 项目经验吸收表 | ✅ 思想采纳（format-agent-skill 的多来源统一、docx-editor 的 token 友好摘要） | 具体吸收方式见 §8/§15 |

---

## 结论

V1.5 的本质升级路径（原方案结论保留，措辞校准）：

```text
会改 Word → 会理解 Word → 会理解模板 → 会理解用户要求
→ 会规划排版（Resolver + Spec）→ 会安全执行（不变量 + Guard）
→ 会检查自己做得对不对（Checker + 渲染 + Agent 目检）
```

执行纪律：每 Phase 以"30 项存量测试不改一行全绿"为底线，新增能力以新增
模块与新增测试交付；对照 §0，凡已实现的不重写。
