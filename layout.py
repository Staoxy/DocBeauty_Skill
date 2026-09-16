# -*- coding: utf-8 -*-
"""Page Layout Engine（DESIGN_V2.md §11）：页面/页码/页眉。"""
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Emu

import constants as C
from utils import ooxml as ox

A4_W = Cm(21.0)
A4_H = Cm(29.7)


def setup_page(doc, effective: dict) -> dict:
    m = effective["page"]["margins_cm"]
    for section in doc.sections:
        section.page_width = A4_W
        section.page_height = A4_H
        section.top_margin = Cm(m["top"])
        section.bottom_margin = Cm(m["bottom"])
        section.left_margin = Cm(m["left"])
        section.right_margin = Cm(m["right"])
    return {"sections": len(doc.sections)}


def usable_width_emu(doc, effective: dict, section_idx: int = None) -> int:
    """节可用宽度 = 页宽 - 左右边距（EMU）。"""
    m = effective["page"]["margins_cm"]
    try:
        sec = doc.sections[section_idx if section_idx is not None else 0]
        return int(sec.page_width - sec.left_margin - sec.right_margin)
    except Exception:
        return int(A4_W - Cm(m["left"]) - Cm(m["right"]))


def strip_empty_pgnumtype(doc) -> int:
    """附录 A-3：清除空 pgNumType（干扰 WPS）。"""
    n = 0
    for el in doc.element.body.iter(qn("w:sectPr")):
        pg = ox.get_child(el, "w:pgNumType")
        if pg is not None and not pg.attrib:
            el.remove(pg)
            n += 1
    return n


def _footer_has_page(section) -> bool:
    for hf in (section.footer, section.first_page_footer, section.even_page_footer):
        try:
            if hf is None:
                continue
            if "PAGE" in hf.part.element.xml:
                return True
        except Exception:
            continue
    return False


def _add_page_field_to_footer(section, instr: str, west: str, east: str,
                              size_pt: float, align: str = "center"):
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    # 清空可能存在的空 run（保留已有内容不动——只在无 PAGE 域时调用）
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER,
                   "left": WD_ALIGN_PARAGRAPH.LEFT,
                   "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(align, WD_ALIGN_PARAGRAPH.CENTER)
    p._p.append(ox.make_page_field(west, east, size_pt, instr))


def _set_pgnumtype(section, fmt: str = None, start: int = None):
    attrs = {}
    if fmt:
        attrs["w:fmt"] = fmt
    if start is not None:
        attrs["w:start"] = str(start)
    ox.sub(section._sectPr, "w:pgNumType", attrs)


def add_page_numbers(doc, effective: dict, options: dict, analysis: dict,
                     zones) -> dict:
    """页码策略（§11.2）：
    - 已有 PAGE 域且 skip_existing -> 跳过
    - 封面节 + 多节 -> 三段式：封面无页码 / 中间节罗马 / 末节阿拉伯 start=1
    - 其他 -> 全部缺页码的节补页脚居中页码
    """
    pn = effective["page_number"]
    west, east, size = pn["font_western"], pn["font_cn"], pn["size_pt"]
    align = pn.get("align", "center")
    detail = {"sections_numbered": 0, "skipped_existing": False,
              "mode": "simple", "warnings": []}

    existing = analysis.get("existing_page_number_fields") or []
    if existing and options.get("page_number_skip_existing", True):
        detail["skipped_existing"] = True
        return detail

    strip_empty_pgnumtype(doc)
    sections = list(doc.sections)
    has_cover = "cover" in zones.zones and len(sections) >= 2

    if has_cover:
        detail["mode"] = "three_part"
        # 封面节：不设页码（解除对上一节的继承，保持空）
        sections[0].footer.is_linked_to_previous = False
        # 中间节：罗马数字（WPS 需 instrText 双写，附录 A-2）
        for sec in sections[1:-1]:
            if not _footer_has_page(sec):
                _add_page_field_to_footer(sec, C.PAGE_FIELD_ROMAN, west, east, size, align)
                _set_pgnumtype(sec, fmt="upperRoman", start=1)
                detail["sections_numbered"] += 1
        # 末节：阿拉伯 start=1
        last = sections[-1]
        if not _footer_has_page(last):
            _add_page_field_to_footer(last, C.PAGE_FIELD_ARABIC, west, east, size, align)
            _set_pgnumtype(last, fmt="decimal", start=1)
            detail["sections_numbered"] += 1
    else:
        for sec in sections:
            if _footer_has_page(sec):
                continue
            _add_page_field_to_footer(sec, C.PAGE_FIELD_PLAIN, west, east, size, align)
            detail["sections_numbered"] += 1
        if analysis.get("cover_like_first_section") and len(sections) == 1:
            detail["warnings"].append(
                "检测到封面但文档只有一个节，无法实现'封面无页码'，全部页面统一编号")
    return detail


def add_header_text(doc, text: str) -> dict:
    """可选简单文字页眉（§11.3）：只给无页眉的节添加。"""
    added = 0
    for sec in doc.sections:
        try:
            if sec.header is not None and "".join(
                    p.text for p in sec.header.paragraphs).strip():
                continue
        except Exception:
            continue
        sec.header.is_linked_to_previous = False
        p = sec.header.paragraphs[0] if sec.header.paragraphs else sec.header.add_paragraph()
        p.text = text
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        added += 1
    return {"headers_added": added}
