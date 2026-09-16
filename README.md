# DocBeauty — 可被 AI Agent 调用的确定性 Word 排版引擎

> AI 的理解能力 + 程序的确定性 = 可靠的文档自动化；
> 而"可靠"必须被证明——每一次输出都附带一份**内容未被改变的自证报告**。

DocBeauty 在**不修改正文内容**的前提下，对已有 .docx 执行规范化排版：统一字体字号、段落缩进行距、识别并套用标题层级、美化表格、缩放居中图片、添加页码、插入 Word 原生目录，并输出 JSON 处理报告（含内容指纹比对）。

设计文档：[DESIGN_V2.md](DESIGN_V2.md)（含 V1→V2 变更追溯表）。触发入口：[SKILL.md](SKILL.md)。

## 快速开始

```bash
pip install -r requirements.txt

# 全自动（推荐先 dry-run 看计划）
python main.py "论文.docx" --dry-run -o output
python main.py "论文.docx" -o output

# 指定配置
python main.py "论文.docx" -c examples/configs/academic_custom.json -o output

# 批量
python main.py --batch ./docs_dir -o output

# 旧格式（Windows+Word 自动转换 .doc/.wps）
python main.py "旧文档.doc" -o output

# 渲染级验收（真 Word：更新域/验证目录/导出 PDF + 视觉指标/PNG）
python main.py "论文.docx" -o output --render-check --pdf

# V1.5 子命令：analyze / audit / plan / apply / verify / render
python main.py analyze "文档.docx" -o output                          # 诊断（只读）
python main.py audit  "论文.docx" --template "学校模板.docx"          # 模板差异清单
python main.py plan   "论文.docx" --template "学校模板.docx" -o out   # formatting_spec.json
python main.py apply  "论文.docx" --template "学校模板.docx" -o out   # 按模板排版
```

退出码：`0` 成功 · `2` 部分成功（errors 已逐项记录）· `3` **内容护栏失败，不产出文件** · `1` 错误。
stdout 仅输出 report.json；日志走 stderr。

## 核心能力

| 模块 | 能力 |
|---|---|
| 结构检测 | 原生样式/大纲级别/编号模板/启发式打分四级标题识别；封面/摘要/参考文献/图注分区（ExclusionZones） |
| 排版引擎 | 中文字体 eastAsia 四件套、firstLineChars 字符级缩进、1.5 倍行距、样式双轨制（样式定义+清残留直接格式） |
| 安全机制 | 修改前快照→修改后指纹 diff→完整性计数→**失败拒绝落盘**；防双重编号；两表间空段保护 |
| 表格 | 三线表/grid 边框、表头跨页重复、行防拆、合并单元格降级处理、页面宽度适配 |
| 图片 | 仅 inline：等比缩小防溢出、居中、图注跟随（keepNext） |
| 页面 | A4/页边距、页码（检测已有不覆盖、封面-罗马-阿拉伯三段式、WPS 兼容开关） |
| 目录 | 原生 TOC 域 + updateFields 自动刷新 + 已有域/手工目录去重 |
| 质检 | Q01–Q16（对账、计数、标题层级、字体覆盖率、兼容红线）+ 可选真 Word 渲染验收 |
| 旧格式 | .doc/.wps 经 Word COM 自动转 .docx（COM 不可用时明确报错并给出替代路径） |
| **V1.5 模板引擎** | Template Analyzer（规范型/样例型双路径）、Role Mapping、Rule Resolver（P1用户>P2模板>P3参考>P5默认，来源逐项追溯）、Formatting Spec、四种输入模式（AUTO/PROMPT/TEMPLATE/REFERENCE）、Intent 受控词表 |
| **V1.5 验证增强** | audit 差异清单（PASS/ATTENTION）、plan 计划输出、PDF→PNG 视觉指标（空白页/越界/孤行，pymupdf 可选）、渲染 PNG 供 Agent 目检 |

## 内容保护机制

- 修改前提取全文段落+表格文本并指纹（sha256，NFC+空白归一化）
- 处理后重提取比对：目录/页码等白名单新增项登记为 `excluded_generated`
- 超链接/书签/域/表格/图片/节数**只增不减**，否则判失败
- 失败 → 删除临时输出，exit_code=3，报告附 diff 明细（段落号+前后文本）
- 边界诚实：浮动图/文本框/公式/脚注在 `untouched_boundary` 中声明"未处理亦未破坏"

## 开发

```bash
python -m pytest tests/ -v        # 66 项测试矩阵（COM/pymupdf 相关项无环境时自动 skip）
```

测试覆盖：护栏不误报/拦截不落盘、防双重编号、两表合并保护、sectPr 保留、
超链接/强调保留、图片缩放、已有/手工目录、修订警告、幂等性、CLI 错误码、三段式页码。

## 项目结构

```text
main.py            # CLI 入口 + §5.2 固定顺序流水线
config.py          # schema 校验 + overrides 深合并
analyzer.py        # 文档分析（§6）
structure.py       # 区域/标题/图注/列表检测 + ExclusionZones（§7-8）
guard.py           # 内容护栏：快照/指纹/diff/完整性（§17）
formatter.py       # 字体+段落+图注（§10）
headings.py        # 标题样式应用+防双重编号（§10.3）
lists.py tables.py images.py layout.py toc.py cleanup.py checker.py report.py
punctuation.py     # 标点规范化（§10.5，唯一改内容操作，默认关闭）
diagnose.py        # 格式诊断（task=diagnose，只报告不修改）
legacy_convert.py  # .doc/.wps -> .docx（Word COM，可选）
render_check.py    # 渲染级验证：真 Word 更新域/验目录/数页数/导出 PDF（可选）
utils/ooxml.py     # OOXML 工具层（schema 顺序插入、域构造、WPS 兼容）
templates/         # academic / formal / clean 参数模板
schemas/           # 输入/输出 JSON Schema
```
