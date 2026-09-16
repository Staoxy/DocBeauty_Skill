# -*- coding: utf-8 -*-
"""DocBeauty V2 常量定义（DESIGN_V2.md §22/附录 B）。"""
import os

SKILL_NAME = "docbeauty"
SKILL_VERSION = "2.0.0"

# ---------------------------------------------------------------------------
# 退出码（§0）
# ---------------------------------------------------------------------------
EXIT_SUCCESS = 0
EXIT_PARTIAL = 2
EXIT_GUARD_FAILED = 3
EXIT_ERROR = 1

# ---------------------------------------------------------------------------
# 错误码（§20）
# ---------------------------------------------------------------------------
FILE_NOT_FOUND = "FILE_NOT_FOUND"
UNSUPPORTED_FILE = "UNSUPPORTED_FILE"
FILE_TOO_LARGE = "FILE_TOO_LARGE"
PASSWORD_PROTECTED = "PASSWORD_PROTECTED"
INVALID_DOCX = "INVALID_DOCX"
EMPTY_DOCUMENT = "EMPTY_DOCUMENT"
CONFIG_INVALID = "CONFIG_INVALID"
OUTPUT_WRITE_FAILED = "OUTPUT_WRITE_FAILED"
GUARD_FAILED = "GUARD_FAILED"
OP_FAILED = "OP_FAILED"

MAX_FILE_BYTES = 100 * 1024 * 1024  # 100MB

# ---------------------------------------------------------------------------
# 操作目录（§3.4）：输入动词 -> 报告完成时态，固定映射
# ---------------------------------------------------------------------------
OPERATION_CATALOG = {
    "page_setup":             "page_setup_done",
    "detect_headings":        "headings_detected",
    "format_headings":        "headings_formatted",
    "normalize_font":         "font_normalized",
    "normalize_paragraph":    "paragraph_normalized",
    "normalize_lists":        "lists_normalized",
    "format_captions":        "captions_formatted",
    "format_tables":          "tables_formatted",
    "format_images":          "images_formatted",
    "page_number":            "page_number_added",
    "header_footer":          "header_footer_done",
    "toc":                    "toc_added",
    "cleanup_blank_paragraphs": "blank_paragraphs_cleaned",
    "punctuation_normalize":  "punctuation_normalized",
    "diagnose":               "diagnosis_done",
}

# operations="auto" 时的执行序（§5.2）
AUTO_OPERATIONS = [
    "page_setup",
    "detect_headings",
    "format_headings",
    "normalize_font",
    "normalize_paragraph",
    "normalize_lists",
    "format_captions",
    "format_tables",
    "format_images",
    "page_number",
    "toc",
    "cleanup_blank_paragraphs",
]

# ---------------------------------------------------------------------------
# 默认 options（§3.4）
# ---------------------------------------------------------------------------
DEFAULT_OPTIONS = {
    "delete_blank_paragraphs": True,
    "heading_promote": "high_confidence",   # off | high_confidence | all
    "strip_style_numbering": True,
    "keep_inline_emphasis": True,
    "page_number_skip_existing": True,
    "table_repeat_header": True,
    "table_borders": "auto",                # auto | grid | threeline | none
    "image_max_width_percent": 100,
    "toc_replace_manual": False,
    "toc_refresh_hint": True,
    "toc_page_break": True,
    "punctuation_normalize": False,
}

# ---------------------------------------------------------------------------
# 中文字号映射（附录 B）
# ---------------------------------------------------------------------------
CN_FONT_SIZE_PT = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5,
}
PT_TO_CN = {round(v, 1): k for k, v in CN_FONT_SIZE_PT.items()}

# twips 换算
CM_TO_TWIP = 567
PT_TO_TWIP = 20
# 半点（half-point, w:sz 的单位）= 0.5pt
def pt_to_half_pt(v: float) -> int:
    return int(round(v * 2))


def pt_to_twip(v: float) -> int:
    return int(round(v * PT_TO_TWIP))


def cm_to_twip(v: float) -> int:
    return int(round(v * CM_TO_TWIP))


# 页码域 instrText（§11.2 / 附录 A：禁止 \* decimal）
PAGE_FIELD_ARABIC = r" PAGE \* arabic \* MERGEFORMAT "
PAGE_FIELD_ROMAN = r" PAGE \* ROMAN \* MERGEFORMAT "
PAGE_FIELD_PLAIN = r" PAGE \* MERGEFORMAT "

# 目录域（§16）
TOC_INSTR = r' TOC \o "1-3" \h \z \u '
TOC_PLACEHOLDER = "（打开文档后按 F9，或右键目录选择“更新域”以生成目录条目）"
TOC_TITLE = "目录"
TOC_HINT = "提示：目录为 Word 原生域，右键目录 → “更新域”可刷新页码。"

# ---------------------------------------------------------------------------
# 编号模板（§8.2）
# ---------------------------------------------------------------------------
HEADING_PATTERNS = {
    "cn_zhang": {
        "l1": r"^第[一二三四五六七八九十百\d]+章",
        "l2": r"^第[一二三四五六七八九十百\d]+节",
        "l3": r"^[一二三四五六七八九十]+、",
        "l4": r"^（[一二三四五六七八九十]+）",
    },
    "cn_gov": {
        "l1": r"^[一二三四五六七八九十]+、",
        "l2": r"^（[一二三四五六七八九十]+）",
        # (?!\d)：小数点后紧跟数字是小数（"0.4x1+1.1x2≤800"），不是编号
        "l3": r"^(\d+)[\.、](?!\d)\s*",
        "l4": r"^（\d+）",
    },
    "numeric": {
        "l1": r"^(\d+)[\.、](?!\d)\s*",
        "l2": r"^\d+\.\d+(?!\.\d)",
        "l3": r"^\d+\.\d+\.\d+(?!\.\d)",
        "l4": r"^\d+\.\d+\.\d+\.\d+",
    },
}
# numeric 层级需要按小数点数细分，特殊处理顺序
NUMERIC_LEVEL_ORDER = ["l4", "l3", "l2", "l1"]

# 图注/表注模式（§15）
CAPTION_PATTERNS = [
    r"^图\s*[\d\-\.．]+",
    r"^表\s*[\d\-\.．]+",
    r"^Figure\s*\d+",
    r"^Table\s*\d+",
    r"^Fig\.\s*\d+",
]

# 区域识别标题（§7.1）
REFERENCE_HEADINGS = ["参考文献", "References", "Bibliography", "REFERENCES"]
APPENDIX_HEADINGS = ["附录", "Appendix", "APPENDIX"]
FRONT_MATTER_HEADINGS = ["摘要", "Abstract", "ABSTRACT", "引言", "Introduction"]
KEYWORD_HEADINGS = ["关键词", "Keywords", "KEYWORDS", "关键字"]

# 文档类型检测证据词（§9.5）
ACADEMIC_EVIDENCE = ["摘要", "关键词", "参考文献", "Abstract", "Keywords", "References", "引言"]
EXPERIMENT_EVIDENCE = ["实验目的", "实验步骤", "实验结果", "实验器材", "实验原理", "实验数据"]
MEETING_EVIDENCE = ["会议时间", "参会人员", "会议议题", "主持人", "记录人", "出席人员"]

SENTENCE_END_PUNCT = "。！？…!?."


def temp_dir_root() -> str:
    root = os.path.join(os.path.expanduser("~"), ".docbeauty", "tmp")
    os.makedirs(root, exist_ok=True)
    return root
