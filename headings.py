# -*- coding: utf-8 -*-
"""Heading 样式应用（DESIGN_V2.md §10.3，防双重编号是硬性要求）。"""
from docx.oxml.ns import qn

from analyzer import heading_level_of
from utils import ooxml as ox


def _heading_style_names(analysis: dict) -> set:
    """analyzer 检出的绑定编号样式名集合（统一小写比较）。"""
    return {str(n).strip().lower()
            for n in analysis.get("numbering", {}).get("styles_with_numPr", [])}


def apply_style_definitions(doc, effective: dict, options: dict, analysis: dict):
    """轨道①：格式化 Heading 1-3 样式定义 + 按需禁用样式编号（§10.3 步骤1）。"""
    detail = {"styles_formatted": [], "numbering_disabled_on_styles": []}
    bound = _heading_style_names(analysis)
    for level in (1, 2, 3):
        sid = ox.ensure_heading_style(doc, level)
        st = ox.find_style_by_sid(doc.styles.element, sid)
        if st is None:
            continue
        h = effective["headings"][f"h{level}"]
        # rPr
        rPr = ox.style_rpr(st)
        ox.set_run_font(rPr, h["font_western"], h["font_cn"], h["size_pt"])
        if h.get("bold"):
            ox.sub(rPr, "w:b")
        color = ox.get_child(rPr, "w:color")
        if color is not None and (color.get(qn("w:val")) or "").lower() not in ("auto", "000000"):
            color.set(qn("w:val"), "auto")
        # pPr
        pPr = ox.style_ppr(st)
        ox.set_alignment(pPr, h.get("align", "left"))
        ox.set_spacing_before_after(pPr, h.get("before_pt", 0), h.get("after_pt", 6))
        ox.set_keep_with_next(pPr, True)
        ox.set_keep_lines(pPr, True)
        ol = ox.get_child(pPr, "w:outlineLvl")
        if ol is None:
            ox.insert_ordered(pPr, ox.make_elem("w:outlineLvl", {"w:val": str(level - 1)}))
        # 防双重编号：样式绑定编号定义 -> numId=0
        if options.get("strip_style_numbering", True):
            nm = ox.get_child(st, "w:name")
            name = (nm.get(qn("w:val")) or "").strip().lower() if nm is not None else ""
            if name in bound:
                if ox.disable_style_numbering(st):
                    detail["numbering_disabled_on_styles"].append(name)
        detail["styles_formatted"].append(sid)
    return detail


def apply_to_paragraphs(doc, zones, effective: dict, options: dict) -> dict:
    """轨道②：对已采信标题段落执行 §10.3 七步算法。"""
    detail = {"promoted": 0, "reformatted": 0,
              "paragraph_numpr_removed": 0, "cleared_direct": 0}
    promote_mode = options.get("heading_promote", "high_confidence")
    if promote_mode == "off":
        return detail

    for idx, level in sorted(zones.headings.items()):
        p = doc.paragraphs[idx]._p
        sid = ox.ensure_heading_style(doc, level)
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            pPr = ox.insert_ordered(p, ox.make_elem("w:pPr"))

        # 步骤2：移除段落级自动编号（编号文字属于内容，绝不动）
        if ox.remove_para_numbering(pPr):
            detail["paragraph_numpr_removed"] += 1

        # 步骤3：清除段落直接格式（保留 sectPr/rPr）
        for tag in ("w:ind", "w:jc", "w:spacing", "w:pBdr", "w:shd"):
            ox.remove_child(pPr, tag)
        # run 直接格式：清字体/字号/颜色/高亮（保留 b/i/u）
        from formatter import runs_of, get_or_add_rpr
        for r in runs_of(p):
            rPr = get_or_add_rpr(r)
            ox.remove_child(rPr, "w:rFonts")
            ox.remove_child(rPr, "w:sz")
            ox.remove_child(rPr, "w:szCs")
            ox.remove_child(rPr, "w:color")
            ox.remove_child(rPr, "w:highlight")
            detail["cleared_direct"] += 1

        # 步骤4/5：pStyle + keepNext/keepLines
        pStyle = ox.sub(pPr, "w:pStyle", {"w:val": sid})
        pStyle.set(qn("w:val"), sid)
        ox.sub(pPr, "w:keepNext")
        ox.sub(pPr, "w:keepLines")

        # 是否原本就是该级 Heading 样式
        prev = heading_level_of(doc.styles.element, p)
        if prev == level:
            detail["reformatted"] += 1
        else:
            detail["promoted"] += 1
    return detail
