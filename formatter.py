# -*- coding: utf-8 -*-
"""Formatting Engine：字体 + 段落 + 图注（DESIGN_V2.md §10.1/§10.2/§15）。"""
from docx.oxml.ns import qn

from utils import ooxml as ox

_W_R = qn("w:r")
_W_PPR = qn("w:pPr")


def get_or_add_ppr(p_el):
    pPr = p_el.find(_W_PPR)
    if pPr is None:
        pPr = ox.insert_ordered(p_el, ox.make_elem("w:pPr"))
    return pPr


def get_or_add_rpr(r_el):
    rPr = ox.get_child(r_el, "w:rPr")
    if rPr is None:
        rPr = ox.insert_ordered(r_el, ox.make_elem("w:rPr"))
    return rPr


def runs_of(p_el):
    """段落内 run 列表，含超链接内 run，排除文本框/图形内部 run（§2.5）。"""
    excluded = set()
    for tag in ("w:txbxContent", "w:drawing", "w:pict", "w:object"):
        for box in p_el.iter(qn(tag)):
            for r in box.iter(_W_R):
                excluded.add(r)
    return [r for r in p_el.iter(_W_R) if r not in excluded]


def set_para_font(p_el, west: str, east: str, size_pt: float,
                  keep_emphasis: bool = True):
    """段落字体统一（§10.1）。keep_emphasis=True 时只动 rFonts/sz。"""
    for r in runs_of(p_el):
        rPr = get_or_add_rpr(r)
        if not keep_emphasis:
            ox.clear_direct_font(rPr)
        ox.set_run_font(rPr, west, east, size_pt)


# ---------------------------------------------------------------------------
# 双轨制轨道①：样式定义 + docDefaults（D1）
# ---------------------------------------------------------------------------
def update_style_definitions(doc, effective: dict):
    """docDefaults + Normal 样式。Normal 不放缩进（避免泄漏进 TOC/Caption 派生样式）。"""
    west = effective["font"]["western"]
    east = effective["font"]["chinese"]
    size = effective["font"]["size_pt"]
    para = effective["paragraph"]

    # docDefaults
    styles_elem = doc.styles.element
    docDefaults = styles_elem.find(qn("w:docDefaults"))
    if docDefaults is None:
        docDefaults = ox.make_elem("w:docDefaults")
        styles_elem.insert(0, docDefaults)
    rPrDefault = ox.get_child(docDefaults, "w:rPrDefault")
    if rPrDefault is None:
        rPrDefault = ox.insert_ordered(docDefaults, ox.make_elem("w:rPrDefault"))
    rPr = ox.get_child(rPrDefault, "w:rPr")
    if rPr is None:
        rPr = ox.insert_ordered(rPrDefault, ox.make_elem("w:rPr"))
    ox.set_run_font(rPr, west, east, size)

    # Normal
    normal = ox.find_style_by_name(styles_elem, "Normal")
    if normal is not None:
        n_rPr = ox.style_rpr(normal)
        ox.set_run_font(n_rPr, west, east, size)
        n_pPr = ox.style_ppr(normal)
        ox.set_line_spacing(n_pPr, para["line_spacing"])
        ox.set_spacing_before_after(n_pPr, para.get("before_pt", 0),
                                    para.get("after_pt", 0))
        ox.set_alignment(n_pPr, para.get("align", "both"))


# ---------------------------------------------------------------------------
# 字体统一（轨道②：正文 run 级）
# ---------------------------------------------------------------------------
def normalize_fonts(doc, zones, effective: dict, options: dict) -> dict:
    """正文字体统一。跳过：标题（样式接管）；图注（captions 单独处理）。
    skip_all 区（封面/手工目录）只做字体（§7.2）。
    """
    west = effective["font"]["western"]
    east = effective["font"]["chinese"]
    size = effective["font"]["size_pt"]
    keep = options.get("keep_inline_emphasis", True)

    changed = 0
    for i, p in enumerate(doc.paragraphs):
        if i in zones.headings or i in zones.captions or i in zones.cover_ids:
            continue
        if not runs_of(p._p):
            continue
        set_para_font(p._p, west, east, size, keep_emphasis=keep)
        changed += 1
    return {"paragraphs_fonted": changed}


# ---------------------------------------------------------------------------
# 段落统一（§10.2）
# ---------------------------------------------------------------------------
def normalize_paragraphs(doc, zones, effective: dict, options: dict) -> dict:
    para = effective["paragraph"]
    chars = para.get("first_line_chars", 2)
    size = effective["font"]["size_pt"]

    changed = 0
    for i, p in enumerate(doc.paragraphs):
        if i in zones.headings or i in zones.captions or i in zones.skip_all:
            continue
        pPr = get_or_add_ppr(p._p)
        ox.set_line_spacing(pPr, para["line_spacing"])
        ox.set_spacing_before_after(pPr, para.get("before_pt", 0),
                                    para.get("after_pt", 0))
        if i not in zones.skip_indent:
            if chars and chars > 0:
                ox.set_first_line_indent(pPr, chars, size)
            else:
                ind = ox.get_child(pPr, "w:ind")
                if ind is not None:
                    for a in ("w:firstLineChars", "w:firstLine",
                              "w:hangingChars", "w:hanging"):
                        if ind.get(qn(a)) is not None:
                            del ind.attrib[qn(a)]
                    if not ind.attrib:
                        pPr.remove(ind)  # 空属性 w:ind 会覆盖样式缩进，直接移除
            ox.set_alignment(pPr, para.get("align", "both"))
        changed += 1
    return {"paragraphs_normalized": changed}


# ---------------------------------------------------------------------------
# 图注/表注（§15）
# ---------------------------------------------------------------------------
def _next_block_is_table(doc, para_idx: int) -> bool:
    body = doc.element.body
    children = list(body)
    target = doc.paragraphs[para_idx]._p
    try:
        k = children.index(target)
    except ValueError:
        return False
    for el in children[k + 1:]:
        if el.tag == qn("w:tbl"):
            return True
        if el.tag == qn("w:p"):
            return False
    return False


def _para_has_inline_image(p_el) -> bool:
    return bool(p_el.findall(".//" + qn("wp:inline")) or
                p_el.findall(".//" + qn("w:pict")))


def format_captions(doc, zones, effective: dict) -> dict:
    cap = effective["captions"]
    west, east, size = cap["font_western"], cap["font_cn"], cap["size_pt"]
    changed = 0
    for i in sorted(zones.captions):
        p = doc.paragraphs[i]
        set_para_font(p._p, west, east, size)
        pPr = get_or_add_ppr(p._p)
        ox.set_alignment(pPr, cap.get("align", "center"))
        ind = ox.get_child(pPr, "w:ind")
        if ind is not None:
            for a in ("w:firstLineChars", "w:firstLine"):
                if ind.get(qn(a)) is not None:
                    del ind.attrib[qn(a)]
        # 表注在表上方 -> keepNext；图注在图下方 -> 图片段 keepNext（§11.4）
        if _next_block_is_table(doc, i):
            ox.set_keep_with_next(pPr, True)
        elif i > 0 and _para_has_inline_image(doc.paragraphs[i - 1]._p):
            img_pPr = get_or_add_ppr(doc.paragraphs[i - 1]._p)
            ox.set_keep_with_next(img_pPr, True)
        changed += 1
    return {"captions_formatted": changed}
