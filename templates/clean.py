# -*- coding: utf-8 -*-
"""Clean 模板（DESIGN_V2.md §9.4）：普通资料/学习笔记/长文档。"""

TEMPLATE = {
    "name": "clean",
    "page": {
        "size": "A4",
        "margins_cm": {"top": 2.54, "bottom": 2.54, "left": 2.54, "right": 2.54},
    },
    "font": {"chinese": "等线", "western": "Calibri", "size_pt": 11},
    "paragraph": {
        "line_spacing": 1.3,
        "first_line_chars": 0,          # 无首行缩进
        "align": "both",
        "before_pt": 0,
        "after_pt": 6,
    },
    "headings": {
        "h1": {"font_cn": "微软雅黑", "font_western": "Calibri",
               "size_pt": 14, "align": "left", "before_pt": 12, "after_pt": 6,
               "bold": True},
        "h2": {"font_cn": "微软雅黑", "font_western": "Calibri",
               "size_pt": 12, "align": "left", "before_pt": 10, "after_pt": 5,
               "bold": True},
        "h3": {"font_cn": "等线", "font_western": "Calibri",
               "size_pt": 11, "align": "left", "before_pt": 8, "after_pt": 4,
               "bold": True},
    },
    "captions": {"font_cn": "等线", "font_western": "Calibri",
                 "size_pt": 10, "align": "center"},
    "tables": {"borders": "grid", "repeat_header": True,
               "header_bold": True, "header_align": "center",
               "header_shade": "F5F5F5"},
    "page_number": {"position": "footer", "align": "center",
                    "font_cn": "等线", "font_western": "Calibri",
                    "size_pt": 9},
    "toc": {"levels": "1-3"},
}
