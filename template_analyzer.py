# -*- coding: utf-8 -*-
"""Template Analyzer（DESIGN_V15 §6）：模板 → Template Specification。

双路径（§6 关键补充）：
- regular：styles.xml 有真正的样式定义 → Style Analyzer 直接提取；
- sample：模板是排好版的示范论文（格式在直接格式化里）→ 复用 structure.detect
  做结构识别，从实际段落提取各角色的有效格式。

只提取不执行；Engine 不支持的维度进 unsupported[]（边界诚实）。
"""
import re
from collections import Counter

from docx.oxml.ns import qn

import analyzer
import constants as C
import structure
import style_analyzer
from utils import ooxml as ox

_EMU_PER_CM = 360000


def _page_size_label(w_emu, h_emu):
    if w_emu is None or h_emu is None:
        return None
    w, h = round(w_emu / _EMU_PER_CM, 1), round(h_emu / _EMU_PER_CM, 1)
    if (w, h) == (21.0, 29.7):
        return "A4"
    if (w, h) == (21.6, 27.9):
        return "Letter"
    return f"{w}x{h}cm"


def _margins_cm(section) -> dict:
    out = {}
    for key, attr in (("top", "top_margin"), ("bottom", "bottom_margin"),
                      ("left", "left_margin"), ("right", "right_margin")):
        v = getattr(section, attr, None)
        out[key] = round(v / _EMU_PER_CM, 2) if v is not None else None
    return out


def _run_fonts_modes(p_el):
    """段落内 run 的字体/字号/加粗众数（sample 路径的有效格式提取）。"""
    east, west, sizes, bolds = Counter(), Counter(), Counter(), Counter()
    for r in p_el.findall(qn("w:r")):
        txt = "".join(t.text or "" for t in r.iter(qn("w:t"))).strip()
        if not txt:
            continue
        rPr = ox.get_child(r, "w:rPr")
        rf = ox.get_child(rPr, "w:rFonts") if rPr is not None else None
        if rf is not None:
            if rf.get(qn("w:eastAsia")):
                east[rf.get(qn("w:eastAsia"))] += len(txt)
            if rf.get(qn("w:ascii")):
                west[rf.get(qn("w:ascii"))] += len(txt)
        sz = ox.get_child(rPr, "w:sz") if rPr is not None else None
        if sz is not None:
            try:
                sizes[round(int(sz.get(qn("w:val"))) / 2, 1)] += len(txt)
            except (TypeError, ValueError):
                pass
        b = ox.get_child(rPr, "w:b") if rPr is not None else None
        if b is not None:
            bolds[1] += len(txt)
    return {
        "eastAsia": east.most_common(1)[0][0] if east else None,
        "ascii": west.most_common(1)[0][0] if west else None,
        "size_pt": sizes.most_common(1)[0][0] if sizes else None,
        "bold": bool(bolds) and bolds[1] > 0,
    }


def _para_effective(p_el, doc_default=None):
    """段落的实际格式：run 众数 + 段落属性（sample 路径）。"""
    fmt = _run_fonts_modes(p_el)
    pPr = p_el.find(qn("w:pPr"))
    jc = ox.get_child(pPr, "w:jc") if pPr is not None else None
    fmt["align"] = jc.get(qn("w:val")) if jc is not None else None
    sp = ox.get_child(pPr, "w:spacing") if pPr is not None else None
    if sp is not None and sp.get(qn("w:line")):
        try:
            v = int(sp.get(qn("w:line")))
            rule = sp.get(qn("w:lineRule")) or "auto"
            fmt["line"] = round(v / 240.0, 3) if rule == "auto" else {"exact_pt": v / 20}
        except (TypeError, ValueError):
            pass
    ind = ox.get_child(pPr, "w:ind") if pPr is not None else None
    if ind is not None:
        flc = ind.get(qn("w:firstLineChars"))
        if flc:
            fmt["first_line_chars"] = int(flc) / 100.0
        elif ind.get(qn("w:firstLine")) and fmt.get("size_pt"):
            try:
                fmt["first_line_chars"] = round(
                    int(ind.get(qn("w:firstLine"))) / (fmt["size_pt"] * 20), 2)
            except (TypeError, ValueError, ZeroDivisionError):
                pass
    if sp is not None:
        for key, attr in (("before_pt", "w:before"), ("after_pt", "w:after")):
            v = sp.get(qn(attr))
            if v is not None:
                try:
                    fmt[key] = round(int(v) / 20.0, 1)
                except (TypeError, ValueError):
                    pass
    # docDefaults 兜底（§2.2 客观测量：无直接格式时继承链的最终落点）
    if doc_default:
        fmt["eastAsia"] = fmt.get("eastAsia") or doc_default.get("eastAsia")
        fmt["ascii"] = fmt.get("ascii") or doc_default.get("ascii")
        fmt["size_pt"] = fmt.get("size_pt") or doc_default.get("size_pt")
    return {k: v for k, v in fmt.items() if v not in (None, False)}


def _style_to_role_params(style_info: dict, dd: dict) -> dict:
    """样式参数 → roles 词表参数（regular 路径）。docDefaults 补缺。"""
    out = {}
    for src, dst in (("eastAsia", "eastAsia"), ("ascii", "ascii"),
                     ("size_pt", "size_pt"), ("bold", "bold"),
                     ("line", "line"), ("first_line_chars", "first_line_chars"),
                     ("align", "align"), ("before_pt", "before_pt"),
                     ("after_pt", "after_pt")):
        v = style_info.get(src)
        if v is not None:
            out[dst] = v
    for key in ("eastAsia", "ascii", "size_pt"):
        if out.get(key) is None and dd.get(key) is not None:
            out[key] = dd[key]
    return out


def _classify_table_borders(tbl) -> dict:
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        return {}
    borders = ox.get_child(tblPr, "w:tblBorders")
    if borders is None:
        return {}
    def val(tag):
        e = ox.get_child(borders, tag)
        return e.get(qn("w:val")) if e is not None else None
    top, bottom, insv, insh = val("w:top"), val("w:bottom"), val("w:insideV"), val("w:insideH")
    if top == "single" and bottom == "single" and insv in (None, "none") \
            and val("w:left") in (None, "none") and val("w:right") in (None, "none"):
        mode = "threeline"
    elif all(v == "single" for v in (top, bottom, val("w:left"), val("w:right"))):
        mode = "grid"
    else:
        mode = "custom"
    # 表头加粗信号
    tr = tbl.find(qn("w:tr"))
    header_bold = False
    if tr is not None:
        for r in tr.iter(qn("w:r")):
            rPr = ox.get_child(r, "w:rPr")
            if rPr is not None and ox.get_child(rPr, "w:b") is not None:
                if "".join(t.text or "" for t in r.iter(qn("w:t"))).strip():
                    header_bold = True
                    break
    return {"borders": mode, "header_bold": header_bold}


def _unsupported_features(doc) -> list:
    out = []
    try:
        cols = doc.sections[0]._sectPr.find(qn("w:cols"))
        if cols is not None and (cols.get(qn("w:num")) or "1") not in ("1", None):
            out.append("columns")
    except Exception:
        pass
    for section in doc.sections:
        for e in section._sectPr.findall(qn("w:headerReference")):
            if e.get(qn("w:type")) in ("even", "first"):
                out.append("odd_even_or_first_page_header")
                break
    return sorted(set(out))


def _footer_page_number(doc) -> dict:
    info = {"present": False}
    for section in doc.sections:
        try:
            footer = section.footer
        except Exception:
            continue
        if footer is None:
            continue
        found_para = None
        for p in footer.paragraphs:
            if "PAGE" in p._p.xml:
                found_para = p
                break
        if found_para is None:
            continue
        info["present"] = True
        jc = ox.get_child(found_para._p.find(qn("w:pPr")), "w:jc")
        info["align"] = jc.get(qn("w:val")) if jc is not None else "left"
        for r in found_para._p.iter(qn("w:r")):
            rPr = ox.get_child(r, "w:rPr")
            sz = ox.get_child(rPr, "w:sz") if rPr is not None else None
            rf = ox.get_child(rPr, "w:rFonts") if rPr is not None else None
            if sz is not None:
                info["size_pt"] = style_analyzer._half_to_pt(sz.get(qn("w:val")))
            if rf is not None and rf.get(qn("w:eastAsia")):
                info["eastAsia"] = rf.get(qn("w:eastAsia"))
        break
    return info


def _header_text(doc) -> dict:
    try:
        header = doc.sections[0].header
    except Exception:
        return {"text": None}
    if header is None:
        return {"text": None}
    text = "\n".join(p.text.strip() for p in header.paragraphs if p.text.strip())
    return {"text": text or None}


def _toc_info(analysis) -> dict:
    if not analysis.get("existing_toc_field"):
        return {}
    return {"levels": "1-3"}  # 常规模板域参数；精确 levels 从 instr 解析的扩展留待 V2


def analyze_template(doc) -> dict:
    """模板 → Template Specification（DESIGN_V15 §6/§7）。"""
    analysis = analyzer.analyze(doc)
    zones = structure.detect(doc, analysis)
    sa = style_analyzer.analyze_styles(doc)
    dd = sa["doc_defaults"]

    heading_used = (analysis["heading_counts"].get("h1", 0)
                    + analysis["heading_counts"].get("h2", 0)
                    + analysis["heading_counts"].get("h3", 0))
    heading_styles_defined = sum(
        1 for v in sa["styles"].values() if v.get("outline_lvl") is not None
        and v["outline_lvl"] <= 2)

    # ---- 路径判定（§6）：标题样式被实际使用 → regular；
    #      注意默认模板总是携带 latent Heading 样式定义，"有定义"不构成判据 ----
    if heading_used >= 1:
        kind, confidence = "regular", min(0.6 + 0.1 * min(heading_used, 4), 0.95)
    else:
        kind = "sample"
        confidence = min(0.5 + 0.1 * min(len(zones.headings), 4), 0.9)

    roles, roles_meta = {}, {}

    if kind == "regular":
        body_info, body_name = style_analyzer.style_by_role(sa, "body")
        if body_info:
            roles["body"] = _style_to_role_params(body_info, dd)
            roles_meta["body"] = body_name
        for lvl, role in ((1, "h1"), (2, "h2"), (3, "h3")):
            # 优先 outlineLvl 匹配（样式名可能是自定义名），再按名
            cand = [(nm, info) for nm, info in sa["styles"].items()
                    if info.get("outline_lvl") == lvl - 1]
            if cand:
                nm, info = cand[0]
                roles[role] = _style_to_role_params(info, dd)
                roles_meta[role] = nm
                continue
            info, nm = style_analyzer.style_by_role(sa, role)
            if info:
                roles[role] = _style_to_role_params(info, dd)
                roles_meta[role] = nm
        cap_info, cap_name = style_analyzer.style_by_role(sa, "caption")
        if cap_info:
            roles["caption"] = _style_to_role_params(cap_info, dd)
            roles_meta["caption"] = cap_name
        title_info, title_name = style_analyzer.style_by_role(sa, "toc_title")
        if title_info:
            roles["toc_title"] = _style_to_role_params(title_info, dd)
            roles_meta["toc_title"] = title_name
        elif "h1" in roles:
            roles["toc_title"] = dict(roles["h1"])
            roles_meta["toc_title"] = "h1(fallback)"
    else:
        # sample 路径：从实际段落提取有效格式
        paras = [p._p for p in doc.paragraphs]
        cover = zones.zones.get("cover")
        title_idx = None
        if cover:
            for i in range(cover[0], cover[1] + 1):
                if analyzer.para_text(paras[i]).strip():
                    title_idx = i
                    break
        if title_idx is None:
            # 退化：前 5 段中字号最大的居中段
            best, best_size = None, 0
            for i in range(min(5, len(paras))):
                fmt = _para_effective(paras[i])
                if fmt.get("size_pt") and fmt["size_pt"] > best_size:
                    best, best_size = i, fmt["size_pt"]
            title_idx = best
        if title_idx is not None:
            roles["title"] = _para_effective(paras[title_idx])
            roles_meta["title"] = f"paragraph {title_idx}"

        by_level = {}
        for i, lv in sorted(zones.headings.items()):
            by_level.setdefault(lv, []).append(i)
        for lv, role in ((1, "h1"), (2, "h2"), (3, "h3")):
            idxs = by_level.get(lv) or []
            fmts = [_para_effective(paras[i]) for i in idxs]
            if not fmts:
                continue
            merged = {}
            for key in ("eastAsia", "ascii", "size_pt", "align", "line",
                        "first_line_chars", "before_pt", "after_pt"):
                vals = [f.get(key) for f in fmts if f.get(key) is not None]
                if vals:
                    merged[key] = Counter(
                        [str(v) for v in vals]).most_common(1)[0][0]
                    if key in ("size_pt", "line", "first_line_chars",
                               "before_pt", "after_pt"):
                        merged[key] = float(merged[key])
                    if key == "first_line_chars" and merged[key] == 0:
                        del merged[key]
            merged.setdefault("bold", True)
            roles[role] = merged
            roles_meta[role] = f"{len(idxs)} paragraphs"

        body_fmts = []
        for i, p in enumerate(paras):
            if i in zones.headings or i in zones.captions or i in zones.skip_all:
                continue
            if not analyzer.para_text(p).strip():
                continue
            body_fmts.append(_para_effective(p, doc_default=dd))
        if body_fmts:
            merged = {}
            for key in ("eastAsia", "ascii", "size_pt", "align", "line",
                        "first_line_chars"):
                vals = [f.get(key) for f in body_fmts if f.get(key) is not None]
                if vals:
                    merged[key] = Counter([str(v) for v in vals]).most_common(1)[0][0]
                    if key in ("size_pt", "line", "first_line_chars"):
                        merged[key] = float(merged[key])
                    if key == "first_line_chars" and merged[key] == 0:
                        del merged[key]
            roles["body"] = merged
            roles_meta["body"] = f"{len(body_fmts)} paragraphs"

        caps = sorted(zones.captions)
        if caps:
            roles["caption"] = _para_effective(paras[caps[0]])
            roles_meta["caption"] = f"{len(caps)} paragraphs"

    # ---- 页面 / 节 / 页眉页脚 ----
    sec0 = doc.sections[0] if doc.sections else None
    page = {"size": _page_size_label(
        sec0.page_width, sec0.page_height) if sec0 else None,
        "margins_cm": _margins_cm(sec0) if sec0 else None}

    sections_out = []
    for i, section in enumerate(doc.sections):
        pg = analyzer.sections_pgnumtype(doc)[i] if i < len(doc.sections) else None
        has_page_field = False
        try:
            has_page_field = any("PAGE" in p._p.xml for p in section.footer.paragraphs)
        except Exception:
            pass
        sections_out.append({"index": i, "page_number": (
            {"format": (pg or {}).get("fmt") or ("decimal" if has_page_field else None),
             "start": (pg or {}).get("start")} if (pg or has_page_field) else None)})

    header_footer = {"header": _header_text(doc), "footer_page_number": _footer_page_number(doc)}

    # ---- 表格 / 编号 / 目录 / 不支持项 ----
    tables = []
    from tables import _top_level_tables
    for tbl in _top_level_tables(doc)[:3]:
        c = _classify_table_borders(tbl)
        if c:
            tables.append(c)
    tables_out = tables[0] if tables else {}

    numbering = {"detected_pattern": zones.numbering_template,
                 "styles_with_numPr": sa["styles_with_numPr"]}

    spec = {
        "template_kind": kind,
        "confidence": round(confidence, 2),
        "page": page,
        "sections": sections_out,
        "header_footer": header_footer,
        "roles": roles,
        "roles_meta": roles_meta,
        "tables": tables_out,
        "numbering": numbering,
        "toc": _toc_info(analysis),
        "unsupported": _unsupported_features(doc),
        "analysis_summary": {
            "paragraph_count": analysis["paragraph_count"],
            "table_count": analysis["table_count"],
            "heading_counts": analysis["heading_counts"],
            "style_usage_top": dict(list(analysis["style_usage_histogram"].items())[:6]),
        },
    }
    return spec
