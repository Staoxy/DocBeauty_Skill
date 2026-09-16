# -*- coding: utf-8 -*-
"""列表策略（DESIGN_V2.md §10.4）：不重排编号、不改编号文字、不动 numId。"""
import re

from docx.oxml.ns import qn

from analyzer import para_text
from utils import ooxml as ox

BULLET_RE = re.compile(r"^[•·▪◦‣\-–]\s+")


def detect_lists(doc, zones) -> int:
    """识别列表段落并登记 skip_indent（须在段落归一之前调用）。

    判据：段落带 numPr，或以项目符号字符开头。
    编号与正文同段混排（"1. xxx 内容…"长段）不视为列表。
    """
    count = 0
    for i, p in enumerate(doc.paragraphs):
        pPr = p._p.find(qn("w:pPr"))
        has_numpr = pPr is not None and pPr.find(qn("w:numPr")) is not None
        text = para_text(p._p).strip()
        bullet = bool(BULLET_RE.match(text))
        if not (has_numpr or bullet):
            continue
        # 混排长段排除：编号开头但超过 60 字符且含句末标点
        if (re.match(r"^\d+[\.、)]?\s*", text) or bullet) and len(text) > 60 \
                and text.endswith(("。", "！", "？", "；")):
            continue
        zones.skip_indent.add(i)
        count += 1
    return count


def apply(doc, zones, effective: dict) -> dict:
    """列表段落只统一行距/段间距；缩进与对齐保持原样（§10.4）。"""
    para = effective["paragraph"]
    changed = 0
    for i in sorted(zones.skip_indent):
        if i in zones.captions or i in zones.skip_all or i in zones.headings:
            continue
        p = doc.paragraphs[i]
        pPr = p._p.find(qn("w:pPr"))
        if pPr is None:
            pPr = ox.insert_ordered(p._p, ox.make_elem("w:pPr"))
        ox.set_line_spacing(pPr, para["line_spacing"])
        ox.set_spacing_before_after(pPr, para.get("before_pt", 0),
                                    para.get("after_pt", 0))
        changed += 1
    return {"list_paragraphs": changed}
