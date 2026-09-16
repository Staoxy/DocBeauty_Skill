# -*- coding: utf-8 -*-
"""Style Analyzer（DESIGN_V15 §5）：Word Style System 的客观画像。

只测量、不诠释（§2.2）：输出 docDefaults、各样式的中英文字体/字号/段落参数、
主题字体、样式使用直方图、编号绑定、以及"声明 vs 实际直接格式"的偏差统计。
"""
import re
from collections import Counter

from docx.oxml.ns import qn

import constants as C
from analyzer import para_text, styles_with_numpr, style_usage
from utils import ooxml as ox

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _half_to_pt(sz_val):
    try:
        return round(int(sz_val) / 2, 1)
    except (TypeError, ValueError):
        return None


def _twips_line_to_multiple(pPr):
    sp = ox.get_child(pPr, "w:spacing")
    if sp is None:
        return None
    line = sp.get(qn("w:line"))
    rule = sp.get(qn("w:lineRule")) or "auto"
    if line is None:
        return None
    try:
        v = int(line)
    except (TypeError, ValueError):
        return None
    if rule == "auto":
        return round(v / 240.0, 3)
    return {"exact_pt": round(v / 20.0, 1)}


def doc_defaults(doc) -> dict:
    """docDefaults rPrDefault：默认中英文字体与字号。"""
    out = {"eastAsia": None, "ascii": None, "size_pt": None}
    styles_elem = doc.styles.element
    dd = styles_elem.find(qn("w:docDefaults"))
    if dd is None:
        return out
    rPr = dd.find(".//" + qn("w:rPr"))
    if rPr is None:
        return out
    rf = ox.get_child(rPr, "w:rFonts")
    if rf is not None:
        out["eastAsia"] = rf.get(qn("w:eastAsia"))
        out["ascii"] = rf.get(qn("w:ascii"))
    sz = ox.get_child(rPr, "w:sz")
    if sz is not None:
        out["size_pt"] = _half_to_pt(sz.get(qn("w:val")))
    return out


def theme_fonts(doc) -> dict:
    """主题字体（theme1.xml 的 major/minor latin+ea）。best effort。"""
    out = {}
    try:
        for part in doc.part.package.parts:
            if str(part.partname).endswith("theme1.xml"):
                root = part.element
                for tag, key in (("majorFont", "major"), ("minorFont", "minor")):
                    node = root.find(f".//{{{_NS}}}{tag}")
                    if node is None:
                        continue
                    latin = node.find(f"{{{_NS}}}latin")
                    ea = node.find(f"{{{_NS}}}ea")
                    out[key] = {
                        "latin": latin.get("typeface") if latin is not None else None,
                        "eastAsia": ea.get("typeface") if ea is not None else None,
                    }
                break
    except Exception:
        pass
    return out


def style_font_para(style_el) -> dict:
    """单个样式元素 → 客观参数（字体/字号/段落）。"""
    info = {"eastAsia": None, "ascii": None, "size_pt": None, "bold": None,
            "line": None, "first_line_chars": None, "align": None,
            "before_pt": None, "after_pt": None, "outline_lvl": None,
            "numbering": None}
    rPr = ox.get_child(style_el, "w:rPr")
    if rPr is not None:
        rf = ox.get_child(rPr, "w:rFonts")
        if rf is not None:
            info["eastAsia"] = rf.get(qn("w:eastAsia"))
            info["ascii"] = rf.get(qn("w:ascii"))
        sz = ox.get_child(rPr, "w:sz")
        if sz is not None:
            info["size_pt"] = _half_to_pt(sz.get(qn("w:val")))
        b = ox.get_child(rPr, "w:b")
        if b is not None:
            info["bold"] = b.get(qn("w:val")) not in ("false", "0", "none")
    pPr = ox.get_child(style_el, "w:pPr")
    if pPr is not None:
        info["line"] = _twips_line_to_multiple(pPr)
        ind = ox.get_child(pPr, "w:ind")
        if ind is not None and ind.get(qn("w:firstLineChars")):
            try:
                info["first_line_chars"] = int(ind.get(qn("w:firstLineChars"))) / 100.0
            except (TypeError, ValueError):
                pass
        jc = ox.get_child(pPr, "w:jc")
        if jc is not None:
            info["align"] = jc.get(qn("w:val"))
        sp = ox.get_child(pPr, "w:spacing")
        if sp is not None:
            for key, attr in (("before_pt", "w:before"), ("after_pt", "w:after")):
                v = sp.get(qn(attr))
                if v is not None:
                    try:
                        info[key] = round(int(v) / 20.0, 1)
                    except (TypeError, ValueError):
                        pass
        ol = ox.get_child(pPr, "w:outlineLvl")
        if ol is not None:
            try:
                info["outline_lvl"] = int(ol.get(qn("w:val")))
            except (TypeError, ValueError):
                pass
        numPr = ox.get_child(pPr, "w:numPr")
        if numPr is not None:
            numId = ox.get_child(numPr, "w:numId")
            info["numbering"] = numId.get(qn("w:val")) if numId is not None else None
    return info


def analyze_styles(doc) -> dict:
    """全量 Style System 画像（DESIGN_V15 §5 输出）。"""
    styles_out = {}
    for st in doc.styles.element.findall(qn("w:style")):
        st_type = st.get(qn("w:type"))
        nm = ox.get_child(st, "w:name")
        name = nm.get(qn("w:val")) if nm is not None else None
        if not name or st_type != "paragraph":
            continue
        styles_out[name] = style_font_para(st)
        styles_out[name]["styleId"] = st.get(qn("w:styleId"))

    # 声明 vs 实际：正文段落的直接格式覆盖率（V1 字号直方图思想的三维扩展）
    actual = {"explicit_font": 0, "explicit_size": 0, "explicit_spacing": 0,
              "total_body_paras": 0}
    body_hist = Counter()
    styles_elem = doc.styles.element
    for p in doc.paragraphs:
        from analyzer import heading_level_of
        if heading_level_of(styles_elem, p._p) in (1, 2, 3):
            continue
        if not para_text(p._p).strip():
            continue
        actual["total_body_paras"] += 1
        pPr = p._p.find(qn("w:pPr"))
        sp = ox.get_child(pPr, "w:spacing") if pPr is not None else None
        if sp is not None and sp.get(qn("w:line")):
            actual["explicit_spacing"] += 1
        for r in p._p.findall(qn("w:r")):
            txt = "".join(t.text or "" for t in r.iter(qn("w:t"))).strip()
            if not txt:
                continue
            rPr = ox.get_child(r, "w:rPr")
            rf = ox.get_child(rPr, "w:rFonts") if rPr is not None else None
            if rf is not None and rf.get(qn("w:eastAsia")):
                actual["explicit_font"] += 1
            sz = ox.get_child(rPr, "w:sz") if rPr is not None else None
            if sz is not None:
                actual["explicit_size"] += 1
                try:
                    body_hist[round(int(sz.get(qn("w:val"))) / 2, 1)] += len(txt)
                except (TypeError, ValueError):
                    pass

    return {
        "doc_defaults": doc_defaults(doc),
        "theme_fonts": theme_fonts(doc),
        "styles": styles_out,
        "style_usage": style_usage(doc),
        "styles_with_numPr": styles_with_numpr(doc),
        "declared_vs_actual": actual,
        "body_size_hist_pt": dict(body_hist),
    }


def style_by_role(sa: dict, role: str):
    """从 Style 画像按规范角色取样式参数（role_mapper 的精确名层，§9）。"""
    names = {
        "body": ["Normal", "正文"],
        "h1": ["heading 1", "标题 1", "Heading 1"],
        "h2": ["heading 2", "标题 2", "Heading 2"],
        "h3": ["heading 3", "标题 3", "Heading 3"],
        "caption": ["Caption", "题注"],
        "title": ["Title", "标题"],
        "toc_title": ["TOC Heading", "TOC 标题", "目录标题"],
    }.get(role, [])
    for name in names:
        if name in sa["styles"]:
            return sa["styles"][name], name
    return None, None
