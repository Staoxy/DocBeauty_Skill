# -*- coding: utf-8 -*-
"""Formal 模板（DESIGN_V2.md §9.3）：正式报告/会议纪要/工作材料。"""

TEMPLATE = {
    "name": "formal",
    "page": {
        "size": "A4",
        "margins_cm": {"top": 2.54, "bottom": 2.54, "left": 2.54, "right": 2.54},
    },
    "font": {"chinese": "宋体", "western": "Times New Roman", "size_pt": 12},
    "paragraph": {
        "line_spacing": 1.5,
        "first_line_chars": 2,
        "align": "both",
        "before_pt": 0,
        "after_pt": 0,
    },
    "headings": {
        "h1": {"font_cn": "黑体", "font_western": "Times New Roman",
               "size_pt": 16, "align": "left", "before_pt": 18, "after_pt": 12},
        "h2": {"font_cn": "黑体", "font_western": "Times New Roman",
               "size_pt": 14, "align": "left", "before_pt": 12, "after_pt": 6},
        "h3": {"font_cn": "黑体", "font_western": "Times New Roman",
               "size_pt": 12, "align": "left", "before_pt": 6, "after_pt": 6},
    },
    "captions": {"font_cn": "宋体", "font_western": "Times New Roman",
                 "size_pt": 10.5, "align": "center"},
    "tables": {"borders": "grid", "repeat_header": True,
               "header_bold": True, "header_align": "center",
               "header_shade": "F2F2F2"},
    "page_number": {"position": "footer", "align": "center",
                    "font_cn": "宋体", "font_western": "Times New Roman",
                    "size_pt": 9},
    "toc": {"levels": "1-3"},
}
