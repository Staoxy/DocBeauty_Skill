# DocBeauty Skill V1 — AI 可读技术设计方案

- **版本**：V1.0
- **项目类型**：Agent Tool / Skill
- **核心技术**：Python + python-docx + lxml
- **主要输入**：DOCX
- **主要输出**：DOCX + JSON 处理报告

---

## 1. 项目定位

DocBeauty 是一个可以被 AI Agent 调用的 Word 文档自动化处理 Skill。

它的核心目标不是生成 Word 内容，而是：**在尽可能不改变用户原始内容的前提下，自动分析 DOCX 文档结构，并根据用户要求或自动判断结果，对文档进行规范化、美化和排版。**

DocBeauty V1 必须能够独立运行，也必须能够作为 Agent 的一个工具被调用。

## 2. 核心原则

### 2.1 内容保护原则

默认情况下：**只修改文档格式，不修改用户正文内容。**

禁止默认执行：

- 改写句子
- 删除正文
- 增加正文
- 修改事实
- 修改数字
- 修改引用内容
- 修改用户原始标题文字

如果用户明确要求 AI 修改内容，则该任务属于 Agent/LLM 的内容处理职责，而不是 DocBeauty 的默认职责。

### 2.2 AI 负责理解，代码负责执行

LLM/Agent 负责：

- 理解用户需求
- 判断文档类型
- 决定处理模式
- 生成结构化处理参数

DocBeauty 负责：

- 读取 DOCX
- 分析文档
- 执行格式操作
- 生成 DOCX
- 检查输出结果

**禁止让 LLM 直接生成或修改 Word XML。**

## 3. V1 支持的用户场景

V1 重点支持以下场景：

- 课程论文
- 普通报告
- 调查报告
- 实验报告
- 会议纪要
- 普通正式文档
- 普通长文档格式统一

V1 暂不专门支持：

- 复杂政府公文
- 专业出版排版
- 复杂学术期刊模板
- PDF / PPT / Excel / LaTeX
- 复杂简历设计

## 4. 输入

DocBeauty 支持以下输入：

### 4.1 DOCX 文件

必需。例如：`input.docx`

### 4.2 用户处理要求

可选。例如：

> 帮我把这份 Word 排版得正式一点。

或者：

> 按照课程论文格式排版，不要修改正文。

### 4.3 Agent 结构化参数

Agent 可以直接向 Skill 提供 JSON：

```json
{
  "task": "beautify",
  "document_type": "academic",
  "style": "formal",
  "content_protection": true,
  "operations": [
    "normalize_font",
    "normalize_paragraph",
    "format_headings",
    "format_tables",
    "format_images",
    "page_number",
    "toc"
  ]
}
```

## 5. 输出

成功执行后必须生成：

```
output/
├── processed.docx
└── report.json
```

- **processed.docx**：为最终处理后的 Word。
- **report.json**：为机器可读处理报告。

## 6. Skill 核心工作流程

```
INPUT DOCX
    ↓ File Validation
    ↓ Document Analyzer
    ↓ Structure Detection
    ↓ Style Detection
    ↓ Task Configuration
    ↓ Formatting Engine
    ↓ Layout Engine
    ↓ Quality Checker
    ↓ Save DOCX
    ↓ Generate JSON Report
```

## 7. Document Analyzer

负责分析输入 Word。必须识别：

- 段落数量
- 表格数量
- 图片数量
- 标题数量
- 超链接数量
- 页眉
- 页脚
- 文档基本属性

输出结构化数据，例如：

```json
{
  "paragraph_count": 126,
  "table_count": 4,
  "image_count": 7,
  "heading_count": 12,
  "has_header": true,
  "has_footer": false
}
```

## 8. Structure Detection

自动分析文档结构。需要识别：

- 文档标题
- 一级标题 / 二级标题 / 三级标题
- 普通正文
- 列表
- 引用
- 表格
- 图片
- 图片说明
- 参考文献区域

## 9. Heading Detection

优先使用 Word 原生 Heading Style。优先级：`Heading 1 > Heading 2 > Heading 3`。

如果原文没有正确使用 Heading Style，则根据以下信息辅助判断：

- 字号
- 是否加粗
- 段落位置
- 文本长度
- 编号模式
- 前后段落关系

例如：

```
一、研究背景      → Heading 1
（一）研究意义    → Heading 2
1. 国内研究现状   → Heading 3
```

但自动判断存在不确定性时，不得破坏原始内容。

## 10. 默认排版策略

V1 提供以下基础模板。

### 10.1 Academic

适合：课程论文、学术作业、调查报告、实验报告。

默认：

- 页面：A4
- 正文：宋体
- 正文大小：小四
- 标题：黑体
- 正文行距：1.5 倍
- 首行缩进：2 字符
- 段落两端对齐

具体字号和学校格式存在差异，因此所有参数必须允许覆盖。

### 10.2 Formal

适合：正式报告、工作材料、会议纪要、普通正式文档。

特点：清晰层级、统一字体、合理留白、标题突出、表格规范。

### 10.3 Clean

适合：普通资料、学习笔记、长文档。

特点：简洁、易读、页面留白合理。

## 11. Formatting Engine

负责实际修改 Word 格式。V1 必须实现：

### 11.1 字体

统一：中文字体、英文字体、数字字体。允许设置：

```json
{
  "font": {
    "chinese": "宋体",
    "english": "Times New Roman",
    "size": 12
  }
}
```

### 11.2 段落

支持：行距、段前、段后、首行缩进、左右缩进、对齐方式。

### 11.3 标题

支持 Heading 1 / Heading 2 / Heading 3，自动统一：字体、字号、粗细、段前段后、分页行为。

## 12. Page Layout Engine

V1 必须支持：A4 页面、页边距、页码、页眉、页脚、分页控制。

### 12.1 标题孤行处理

避免标题单独留在页底（如「第二章 实验结果」在本页、正文在下一页）。如果可能，应让标题与至少一个正文段落保持在一起。

### 12.2 空白页面控制

检测异常空白段落，可以删除明显无意义的连续空白段落。但是：**不得删除可能具有排版意义的分页。**

## 13. Table Processor

V1 支持基础表格美化。自动处理：

- 表格宽度
- 单元格对齐
- 字体统一
- 字号统一
- 表头突出
- 边框
- 行高
- 页面宽度适配

原则：**不修改表格中的原始数据。**

## 14. Image Processor

V1 支持基础图片处理，包括：

- 图片最大宽度限制
- 页面居中
- 保持比例
- 防止明显超出页面
- 图片与图注尽量保持在一起

不得：修改图片内容、自动裁剪重要区域、删除原图片。

## 15. Caption Detection

自动识别类似以下格式：

```
图1 xxx   图 1 xxx
表1 xxx   表 1 xxx
```

将其识别为 Figure Caption / Table Caption，并进行统一格式。

## 16. TOC

V1 支持自动生成 Word 原生目录。只有在 Heading 1/2/3 结构明确时才自动生成。

目录必须使用 **Word 原生 TOC Field**，不要简单生成普通文本目录。

## 17. Page Number

支持页脚页码：居中、左侧、右侧。默认页脚居中。

**如果用户没有明确要求，不应该覆盖已有页码。**

## 18. Content Protection

默认：

```json
{ "content_protection": true }
```

开启后：

- **允许**：改字体、改字号、改段落、改标题样式、改页面布局、改表格样式、调整图片位置
- **禁止**：修改正文文字、修改数字、修改引用、修改表格数据、修改图片内容

## 19. Quality Checker

生成 Word 后必须执行检查，至少检查：

1. 文件是否成功生成
2. DOCX 是否可以正常打开
3. 段落数量是否异常变化
4. 表格数量是否异常变化
5. 图片数量是否异常变化
6. 标题层级是否存在
7. 是否存在超出页面的表格
8. 是否存在明显异常空白
9. 页码是否成功添加
10. 目录是否成功生成

## 20. 内容变化检测

如果 `content_protection = true`，必须进行基础内容一致性检查。

- 处理前提取：`paragraph_text`、`table_text`
- 处理后再次提取
- 比较：before vs after

如果发现文字内容发生变化：

```json
{
  "content_changed": true,
  "status": "warning"
}
```

Skill 不得假装没有变化。

## 21. JSON 处理报告

输出示例：

```json
{
  "status": "success",
  "input_file": "input.docx",
  "output_file": "processed.docx",
  "document_type": "academic",
  "operations": [
    "font_normalized",
    "paragraph_normalized",
    "headings_formatted",
    "tables_formatted",
    "images_formatted",
    "page_number_added"
  ],
  "statistics": {
    "paragraphs": 126,
    "tables": 4,
    "images": 7,
    "headings": 12
  },
  "warnings": []
}
```

## 22. 错误处理

如果输入不是 DOCX：

```json
{ "status": "error", "code": "UNSUPPORTED_FILE" }
```

如果文件损坏：

```json
{ "status": "error", "code": "INVALID_DOCX" }
```

如果某个操作失败：

```json
{
  "status": "partial_success",
  "warnings": [
    "Table 3 exceeds page width and could not be fully optimized."
  ]
}
```

**不能因为单个表格失败而直接丢失整个输出文件。**

## 23. 项目结构

建议：

```
docbeauty/
├── SKILL.md
├── README.md
├── requirements.txt
├── main.py
├── analyzer.py
├── structure.py
├── formatter.py
├── layout.py
├── tables.py
├── images.py
├── headings.py
├── toc.py
├── checker.py
├── content_guard.py
├── templates/
│   ├── academic.py
│   ├── formal.py
│   └── clean.py
├── schemas/
│   ├── input_schema.json
│   └── output_schema.json
├── examples/
│   ├── input/
│   └── output/
└── tests/
    ├── test_analyzer.py
    ├── test_formatter.py
    ├── test_tables.py
    ├── test_images.py
    └── test_content_guard.py
```

## 24. main.py 职责

main.py 是 Skill 的主要入口。流程：

```
receive request
    ↓ validate input
    ↓ load document
    ↓ analyze
    ↓ build configuration
    ↓ execute operations
    ↓ quality check
    ↓ save output
    ↓ generate report
```

## 25. Agent 调用协议

推荐：

```json
{
  "task": "beautify",
  "input_file": "/path/to/input.docx",
  "document_type": "auto",
  "style": "auto",
  "content_protection": true,
  "operations": "auto"
}
```

Agent 可以指定：

```json
{
  "task": "beautify",
  "input_file": "/path/to/input.docx",
  "document_type": "academic",
  "style": "formal",
  "content_protection": true
}
```

## 26. 自动模式

如果 `document_type = auto`，Skill 根据文档特征进行基础判断。例如：

- 论文关键词 + 标题层级 + 参考文献 → academic
- 会议时间 + 参会人员 + 议题 → meeting
- 实验目的 + 实验步骤 + 实验结果 → experiment

自动判断不确定时：`document_type = generic`，而不是强行套模板。

## 27. Agent 调用示例

用户：「把这个 Word 排版得正式一点，内容不要改。」

Agent 应调用：

```json
{
  "task": "beautify",
  "input_file": "input.docx",
  "document_type": "auto",
  "style": "formal",
  "content_protection": true,
  "operations": "auto"
}
```

Skill 执行：

```
分析 → 识别文档 → 选择 Formal → 格式统一 → 页面优化 → 内容一致性检查 → 生成 Word
```

## 28. V1 明确不做的事情

为了控制项目规模，V1 禁止加入：

- AI 自动改写正文 / 润色 / 扩写 / 缩写
- 自动生成论文
- 自动生成参考文献
- PDF 转换
- PPT / Excel 处理
- OCR
- 复杂图表生成
- 自动改变图片内容
- 复杂出版社排版
- 云端协作
- 用户账号系统

这些属于未来版本。

## 29. V1 成功标准

DocBeauty V1 完成的标准不是「代码能运行」，必须满足：

**基础**

- 能读取 DOCX
- 能生成 DOCX
- 文件可以正常打开

**自动化**

- 能自动识别基础文档结构
- 能自动统一字体
- 能自动统一段落
- 能自动处理标题
- 能处理表格
- 能处理图片
- 能添加页码
- 能生成目录

**安全**

- 默认不改变正文内容
- 能检测内容变化
- 出错时不会静默丢失内容

**Agent**

- Agent 能够通过结构化参数调用 Skill
- Skill 返回结构化 JSON 结果
- Agent 能够获得最终 DOCX 文件

## 30. V1 开发优先级

- **Phase 1**：DOCX 读取 → DOCX 输出 → 基础格式统一
- **Phase 2**：标题识别 → Heading Style → 段落格式
- **Phase 3**：页面布局 → 页码 → 页眉页脚
- **Phase 4**：表格 → 图片 → 图注
- **Phase 5**：目录 → Content Guard → Quality Checker
- **Phase 6**：Agent Interface → JSON Schema → 完整 Skill 封装

## 31. V1 最终架构

```
              AI AGENT
                 │
                 │  Skill Request
                 ↓
           ┌───────────┐
           │ DocBeauty │
           └─────┬─────┘
                 ↓
        Document Analyzer
                 ↓
       Structure Detector
                 ↓
          Configuration
                 ↓
     ┌───────────┼───────────┐
     ↓           ↓           ↓
 Formatter     Layout     Processor
     │           │        ┌──┴──┐
     │           │        ↓     ↓
     │           │      Table  Image
     └───────────┼───────────┘
                 ↓
          Quality Checker
                 ↓
          Content Guard
                 ↓
          output.docx
              +
          report.json
                 ↓
               AGENT
                 ↓
                USER
```

## 32. V1 设计哲学

DocBeauty 不是「一个会写 Word 的 AI」，而应该是：**一个可以被 AI Agent 调用的确定性 Word 文档处理引擎。**

- Agent 负责理解用户。
- DocBeauty 负责执行。

最终形成：

> **AI 的理解能力 + 程序的确定性 = 可靠的文档自动化**

这应该作为 DocBeauty V1 后续开发的核心原则。
