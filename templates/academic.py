# -*- coding: utf-8 -*-
"""Academic 模板（DESIGN_V2.md §9.2）：课程论文/学术作业/调查报告/实验报告。"""

TEMPLATE = {
    "name": "academic",
    "page": {
        "size": "A4",
        "margins_cm": {"top": 2.5, "bottom": 2.5, "left": 3.0, "right": 2.5},
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
               "size_pt": 16, "align": "center", "before_pt": 24, "after_pt": 18},
        "h2": {"font_cn": "黑体", "font_western": "Times New Roman",
               "size_pt": 14, "align": "left", "before_pt": 12, "after_pt": 6},
        "h3": {"font_cn": "黑体", "font_western": "Times New Roman",
               "size_pt": 12, "align": "left", "before_pt": 6, "after_pt": 6},
    },
    "captions": {"font_cn": "宋体", "font_western": "Times New Roman",
                 "size_pt": 10.5, "align": "center"},
    "tables": {"borders": "threeline", "repeat_header": True,
               "header_bold": True, "header_align": "center",
               "header_shade": "F2F2F2"},
    "page_number": {"position": "footer", "align": "center",
                    "font_cn": "宋体", "font_western": "Times New Roman",
                    "size_pt": 9},
    "toc": {"levels": "1-3"},
}
