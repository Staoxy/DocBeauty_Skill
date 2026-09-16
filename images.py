# -*- coding: utf-8 -*-
"""Image Processor（DESIGN_V2.md §14）：仅 inline，只缩小不放大，等比缩放。"""
from docx.oxml.ns import qn
from docx.shared import Emu

from layout import usable_width_emu
from utils import ooxml as ox


def _owning_paragraph(inline):
    el = inline
    while el is not None and el.tag != qn("w:p"):
        el = el.getparent()
    return el


def format_images(doc, effective: dict, options: dict) -> dict:
    pct = float(options.get("image_max_width_percent", 100)) / 100.0
    usable = usable_width_emu(doc, effective)
    max_w = int(usable * pct)
    detail = {"inline_images": len(doc.inline_shapes), "scaled": 0, "centered": 0}

    for shape in doc.inline_shapes:
        try:
            w, h = int(shape.width), int(shape.height)
        except Exception:
            continue
        if w > 0 and max_w > 0 and w > max_w:
            factor = max_w / w
            shape.width = Emu(int(w * factor))
            shape.height = Emu(int(h * factor))  # 等比，禁止只改一边
            detail["scaled"] += 1
        # 所在段落居中 + 清首行缩进
        p = _owning_paragraph(shape._inline)
        if p is not None:
            pPr = p.find(qn("w:pPr"))
            if pPr is None:
                pPr = ox.insert_ordered(p, ox.make_elem("w:pPr"))
            ox.set_alignment(pPr, "center")
            ind = ox.get_child(pPr, "w:ind")
            if ind is not None:
                for a in ("w:firstLineChars", "w:firstLine"):
                    if ind.get(qn(a)) is not None:
                        del ind.attrib[qn(a)]
            detail["centered"] += 1
    return detail
