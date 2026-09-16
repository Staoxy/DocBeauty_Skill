---
name: docbeauty
description: "Word 文档排版美化引擎：在【不修改正文内容】的前提下统一字体字号、段落缩进行距、识别并套用标题层级、美化表格、缩放居中图片、添加页码、生成原生目录，并输出 JSON 处理报告自证内容未变。适用于：课程论文/毕业论文排版、报告格式统一、Word 美化、字体混乱修复、按学校格式要求调整 docx。Beautify or format an existing Word (.docx) document WITHOUT changing its text content: normalize fonts/paragraphs, detect heading levels, format tables/images, add page numbers, insert native TOC, and produce a JSON report proving content integrity. Use when the user wants formatting only; do NOT use when the user asks to rewrite, polish, or generate document content."
---

# DocBeauty — 确定性 Word 排版引擎

## 何时调用本 Skill

- 用户给了 .docx，要求"排版/美化/格式统一/字体统一/加页码/生成目录"，且**不需要改写内容**
- 用户给出学校/单位的格式要求（文字描述或格式要求文档），要求按规范调整现有 Word

## 何时不调用

- 用户要求改写、润色、扩写、生成正文内容（这是 LLM 的职责）
- 输入是 .pdf/.md 等（本 Skill 只处理 Word 系格式）
- 注意：.doc/.wps 在 Windows 且装有 Word/WPS 时会**自动转换为 .docx** 再处理（报告中 `input.converted_from` 可见）；无 COM 环境时返回 UNSUPPORTED_FILE 并提示用户另存为 .docx
- 用户要从零生成一份新文档（用生成型 docx skill）

## 五条默认行为（必须先读懂）

1. **内容保护默认开启**：任何正文文字变化都会导致 exit_code=3 且**不产出文件**。目录、页码域是仅有的白名单插入项。
2. `document_type`/`style` 默认 `auto`；检测不确定时回退 `generic + formal`，不会强行套模板。
3. 已有页码、已有目录、页眉页脚默认**不覆盖**，只补缺失。
4. 标点/符号规范化默认关闭（会改文字）；要开启必须同时 `content_protection=false`。
5. `status != "success"` 时，你必须把 report 里的 warnings/errors 如实转述给用户，**不得宣称成功**；`guard_failed` 必须明确告知"未产出文件"。

## 标准工作流（推荐）

```bash
# 0. （可选）诊断：用户问"这份文档哪里乱"时，只报告不修改
python <skill_dir>/main.py "输入.docx" -c diagnose.json -o output
# diagnose.json: {"task": "diagnose"}  ->  报告 diagnosis.issues 按 high/medium/low/info 分级

# 1. 先 dry-run：不修改任何东西，看计划
python <skill_dir>/main.py "输入.docx" --dry-run -o output

# 2. 向用户展示计划（模板选择、将提升的标题、将插入的目录/页码）

# 3. 正式运行（stdout 是唯一要解析的 JSON 报告）
python <skill_dir>/main.py "输入.docx" -o output
# 退出码: 0 成功 | 2 部分成功 | 3 内容护栏失败(无文件) | 1 错误

# 4. （Windows+Word 可选）渲染级验收：真 Word 更新域、验证目录条目/页码、导出 PDF
python <skill_dir>/main.py "输入.docx" -o output --render-check --pdf
```

- stderr 是人类日志；**stdout 是 report.json**（Agent 解析入口）。
- 输出：`output/<原名>_beautified.docx` + `output/<原名>_report.json`。

## 配置（可选）

不传配置 = 全自动。需要控制时写一个 JSON（schema: `--print-schema` 查看）：

```json
{
  "task": "beautify",
  "document_type": "auto",
  "style": "auto",
  "content_protection": true,
  "operations": "auto",
  "overrides": {
    "font": {"chinese": "仿宋", "size_pt": 12},
    "page": {"margins_cm": {"top": 2.5, "bottom": 2.5, "left": 3.0, "right": 2.5}}
  },
  "options": {"heading_promote": "high_confidence"}
}
```

```bash
python <skill_dir>/main.py "输入.docx" -c config.json -o output
```

常用开关（`options`）：
- `heading_promote`: `off`（不自动提升标题，只报告候选）/ `high_confidence`（默认）/ `all`
- `table_borders`: `auto` / `grid` / `threeline`（三线表）/ `none`
- `toc_replace_manual`: 检测到手工文字目录时是否替换（默认 false，只警告）
- `image_max_width_percent`: 图片最大宽度百分比（默认 100 = 页面可用宽度）

## 你的职责（AI 侧）

1. **读用户的格式要求文档**（学校格式规定 .docx/PDF/网页）：由你阅读并翻译为 `overrides` JSON 传给 Skill——Skill 不读格式要求文档。
2. dry-run 后向用户转述计划：`config.document_type`（检测值+证据）、`operations_detail[0].heading_promote_plan`（将提升的标题段落清单）、候选未提升段落。
3. 运行后解析 report.json 并如实转述：
   - `status` / `exit_code`
   - `statistics`（before/after 段落、标题数）
   - `content_guard.content_changed`（应为 false）与 `quality_checks`
   - `warnings`（修订/批注提示、手工目录提示等）逐条转述
   - `errors` 非空时说明哪些操作失败（partial_success 仍有输出文件）
4. 用户对排版不满意时，调整 `overrides`/`options` 重跑——Skill 是幂等的，重复运行安全。

## 报告字段速查

| 字段 | 含义 |
|---|---|
| `status` | success / partial_success / guard_failed / dry_run / error |
| `config.style_resolved` | 实际使用的模板（academic/formal/clean） |
| `config.document_type` | 自动检测的文档类型 + 置信度 + 证据词 |
| `statistics.reconciliation` | 段落数对账（删除空段/新增目录段） |
| `content_guard` | 内容指纹、diff 明细、完整性计数（超链接/表格/图片/节数） |
| `quality_checks` | Q01–Q16 质量检查结果 |
| `untouched_boundary` | 浮动图片/文本框/公式/脚注：未处理亦未破坏 |
| `render_check` | （--render-check）真 Word 渲染验证：页数、目录条目数与预览、页码是否渲染、PDF 路径 |
| `diagnosis` | （task=diagnose）格式问题分级清单：FONT_CHAOS/INDENT_MISSING/NO_HEADINGS/TABLE_OVERFLOW/TRACKED_CHANGES 等 |

## 已知边界（如实告知用户）

- 只处理**嵌入式（inline）图片**；浮动图片只计数不移动不删除
- 文本框、公式（OMML）、脚注内容不参与格式化，也不会被破坏
- 含未接受修订的文档：会警告建议先接受修订再排版
- .doc/.wps 需先转换为 .docx
