# -*- coding: utf-8 -*-
"""Document Analyzer（DESIGN_V2.md §6）。"""
import re
from collections import Counter

from docx.oxml.ns import qn

import constants as C
from utils import ooxml as ox

_W_T = qn("w:t")
_CJK_RE = re.compile(r"[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]")


def para_text(p) -> str:
    """XML 级段落文本提取（含超链接内文本，含域结果文本）。"""
    return "".join(t.text or "" for t in p.iter(_W_T))


def all_text(doc) -> str:
    return "\n".join(para_text(p._p if hasattr(p, "_p") else p)
                     for p in doc.paragraphs)


def _footnote_count(doc) -> int:
    try:
        for part in doc.part.package.parts:
            if str(part.partname).endswith("footnotes.xml"):
                return len(part.element.findall(qn("w:footnote"))) - 2  # 去掉分隔符定义
    except Exception:
        pass
    return 0


def _comment_count(doc) -> int:
    try:
        for part in doc.part.package.parts:
            if str(part.partname).endswith("comments.xml"):
                return len(part.element.findall(qn("w:comment")))
    except Exception:
        pass
    return 0


def _hf_parts_with_page_field(doc) -> list:
    """含 PAGE 域的页眉/页脚部件名（§11.2 跳过依据）。"""
    found = []
    seen = set()
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.even_page_header, section.even_page_footer,
                   section.first_page_header, section.first_page_footer):
            try:
                part = hf.part
            except Exception:
                continue
            if part is None or id(part) in seen:
                continue
            seen.add(id(part))
            xml = part.element.xml
            if re.search(r'\bPAGE\b', xml):
                found.append(str(part.partname).split("/")[-1])
    return found


def _has_toc_field(doc) -> bool:
    body_xml = doc.element.body.xml
    if "TOC" not in body_xml:
        return False
    # instrText 或 fldSimple instr 中含 TOC 指令
    for it in doc.element.body.iter(qn("w:instrText")):
        if it.text and re.search(r'\bTOC\b', it.text):
            return True
    for fs in doc.element.body.iter(qn("w:fldSimple")):
        if re.search(r'\bTOC\b', fs.get(qn("w:instr")) or ""):
            return True
    return False


def heading_level_of(doc_styles_elem, p) -> int:
    """返回 1-3（标题级）、9（outlineLvl>=3 的其他大纲级）或 0（非标题）。

    优先级（§8.1）：Heading 样式 > 段落 outlineLvl > （numPr 由 structure 判断）。
    """
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return 0
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is not None:
        sid = pStyle.get(qn("w:val"))
        st = ox.find_style_by_sid(doc_styles_elem, sid) if sid else None
        if st is not None:
            nm = ox.get_child(st, "w:name")
            name = (nm.get(qn("w:val")) or "").strip().lower() if nm is not None else ""
            m = re.match(r"^heading (\d)$", name)
            if m:
                lvl = int(m.group(1))
                return lvl if lvl <= 3 else 9
            m = re.match(r"^标题 (\d)$", name)
            if m:
                lvl = int(m.group(1))
                return lvl if lvl <= 3 else 9
            # 样式定义中的 outlineLvl
            spPr = ox.get_child(st, "w:pPr")
            if spPr is not None:
                ol = spPr.find(qn("w:outlineLvl"))
                if ol is not None:
                    lvl = int(ol.get(qn("w:val"))) + 1
                    return lvl if lvl <= 3 else 9
    ol = pPr.find(qn("w:outlineLvl"))
    if ol is not None:
        lvl = int(ol.get(qn("w:val"))) + 1
        return lvl if lvl <= 3 else 9
    return 0


def style_usage(doc) -> Counter:
    usage = Counter()
    styles_elem = doc.styles.element
    for p in doc.paragraphs:
        pPr = p._p.find(qn("w:pPr"))
        label = "Normal"
        if pPr is not None:
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                sid = pStyle.get(qn("w:val"))
                st = ox.find_style_by_sid(styles_elem, sid) if sid else None
                if st is not None:
                    nm = ox.get_child(st, "w:name")
                    if nm is not None:
                        label = nm.get(qn("w:val"))
                        m = re.match(r"^heading (\d)$", label.strip().lower())
                        if m:
                            label = f"Heading {m.group(1)}"
        usage[label] += 1
    return usage


def styles_with_numpr(doc) -> list:
    """绑定了多级编号定义的样式名（§6 防双重编号输入）。"""
    result = []
    for st in doc.styles.element.findall(qn("w:style")):
        pPr = ox.get_child(st, "w:pPr")
        if pPr is None:
            continue
        numPr = ox.get_child(pPr, "w:numPr")
        if numPr is None:
            continue
        numId = ox.get_child(numPr, "w:numId")
        if numId is not None and numId.get(qn("w:val")) not in (None, "0"):
            nm = ox.get_child(st, "w:name")
            result.append(nm.get(qn("w:val")) if nm is not None else "?")
    return result


def num_ids_in_use(doc) -> list:
    ids = set()
    for numPr in doc.element.body.iter(qn("w:numPr")):
        numId = ox.get_child(numPr, "w:numId")
        if numId is not None:
            v = numId.get(qn("w:val"))
            if v not in (None, "0"):
                ids.add(v)
    return sorted(ids)


def _doc_default_size_pt(doc) -> float:
    """docDefaults 的默认字号（pt），缺省 11。"""
    styles_elem = doc.styles.element
    dd = styles_elem.find(qn("w:docDefaults"))
    if dd is not None:
        for sz in dd.iter(qn("w:sz")):
            try:
                return int(sz.get(qn("w:val"))) / 2
            except (TypeError, ValueError):
                break
    return 11.0


def font_size_histogram(doc) -> dict:
    """按字符数统计字号分布。无显式 sz 的 run 按文档默认字号计入
    （否则大字号封面会淹没正文众数，导致标题打分失真）。"""
    hist = Counter()
    default_pt = round(_doc_default_size_pt(doc), 1)
    for p in doc.paragraphs:
        for r in p._p.findall(qn("w:r")):
            txt = "".join(t.text or "" for t in r.iter(_W_T)).strip()
            if not txt:
                continue
            rPr = ox.get_child(r, "w:rPr")
            sz = ox.get_child(rPr, "w:sz") if rPr is not None else None
            if sz is not None:
                try:
                    hist[round(int(sz.get(qn("w:val"))) / 2, 1)] += len(txt)
                except (TypeError, ValueError):
                    pass
            else:
                hist[default_pt] += len(txt)
    return dict(hist)


def body_font_size_mode(hist: dict):
    """正文字号众数（标题打分基线）。无数据时 12pt。"""
    if not hist:
        return 12.0
    return sorted(hist.items(), key=lambda kv: -kv[1])[0][0]


def cjk_ratio(text: str) -> float:
    if not text:
        return 0.0
    cjk = len(_CJK_RE.findall(text))
    letters = len(re.findall(r"[A-Za-z]", text))
    denom = cjk + letters
    return cjk / denom if denom else 0.0


def sections_pgnumtype(doc) -> list:
    out = []
    for section in doc.sections:
        pgNumType = ox.get_child(section._sectPr, "w:pgNumType")
        if pgNumType is None:
            out.append(None)
        else:
            out.append({
                "fmt": pgNumType.get(qn("w:fmt")),
                "start": pgNumType.get(qn("w:start")),
            })
    return out


def _cover_like(doc) -> bool:
    """首节封面信号：首节段落数 <=10 且存在 >=22pt 居中段落（§6）。"""
    paras = doc.paragraphs
    if not paras:
        return False
    # 首节范围；无分节符的单节文档只看前 5 段
    n = 0
    multi_section = False
    for p in paras:
        n += 1
        pPr = p._p.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
            multi_section = True
            break
    limit = n if multi_section else min(5, n)
    if limit > 12:
        return False
    for p in paras[:limit]:
        centered = False
        big = False
        pPr = p._p.find(qn("w:pPr"))
        if pPr is not None:
            jc = ox.get_child(pPr, "w:jc")
            centered = jc is not None and jc.get(qn("w:val")) == "center"
        for r in p._p.findall(qn("w:r")):
            rPr = ox.get_child(r, "w:rPr")
            if rPr is None:
                continue
            sz = ox.get_child(rPr, "w:sz")
            try:
                if sz is not None and int(sz.get(qn("w:val"))) >= 44:  # 22pt
                    big = True
            except (TypeError, ValueError):
                pass
        if centered and big:
            return True
    return False


def detect_document_type(doc, analysis: dict) -> dict:
    """auto 模式文档类型检测（§9.5）。返回 {value, confidence, evidence}。"""
    text = all_text(doc)
    ev_exp = [w for w in C.EXPERIMENT_EVIDENCE if w in text]
    ev_aca = [w for w in C.ACADEMIC_EVIDENCE if w in text]
    ev_meet = [w for w in C.MEETING_EVIDENCE if w in text]
    if len(ev_exp) >= 2:
        return {"value": "experiment", "confidence": min(0.5 + 0.15 * len(ev_exp), 0.95),
                "evidence": ev_exp}
    if len(ev_aca) >= 2:
        return {"value": "academic", "confidence": min(0.5 + 0.15 * len(ev_aca), 0.95),
                "evidence": ev_aca}
    if len(ev_meet) >= 2:
        return {"value": "meeting", "confidence": min(0.5 + 0.15 * len(ev_meet), 0.95),
                "evidence": ev_meet}
    return {"value": "generic", "confidence": 0.3, "evidence": []}


def analyze(doc) -> dict:
    body = doc.element.body
    styles_elem = doc.styles.element

    heading_counts = Counter()
    for p in doc.paragraphs:
        lvl = heading_level_of(styles_elem, p._p)
        if lvl == 0:
            continue
        key = "h1" if lvl == 1 else "h2" if lvl == 2 else "h3" if lvl == 3 else "outline_lvl_other"
        heading_counts[key] += 1

    text_all = all_text(doc)
    hist = font_size_histogram(doc)

    tracked = {
        "ins": len(body.findall(".//" + qn("w:ins"))),
        "del": len(body.findall(".//" + qn("w:del"))),
    }

    analysis = {
        "paragraph_count": len(doc.paragraphs),
        "table_count": len(body.findall(".//" + qn("w:tbl"))),
        "inline_image_count": len(body.findall(".//" + qn("wp:inline"))),
        "floating_image_count": len(body.findall(".//" + qn("wp:anchor"))),
        "textbox_count": len(body.findall(".//" + qn("w:txbxContent"))),
        "omml_equation_count": len(body.findall(".//" + qn("m:oMath"))),
        "footnote_count": _footnote_count(doc),
        "hyperlink_count": len(body.findall(".//" + qn("w:hyperlink"))),
        "heading_counts": dict(heading_counts),
        "style_usage_histogram": dict(style_usage(doc)),
        "numbering": {
            "styles_with_numPr": styles_with_numpr(doc),
            "num_ids_in_use": num_ids_in_use(doc),
        },
        "existing_toc_field": _has_toc_field(doc),
        "existing_manual_toc": False,  # 由 structure.py 补充
        "existing_page_number_fields": _hf_parts_with_page_field(doc),
        "section_count": len(doc.sections),
        "sections_pgnumtype": sections_pgnumtype(doc),
        "tracked_changes": tracked,
        "comment_count": _comment_count(doc),
        "cjk_char_ratio": round(cjk_ratio(text_all), 3),
        "font_size_histogram_pt": {str(k): v for k, v in hist.items()},
        "body_font_size_mode_pt": body_font_size_mode(hist),
        "cover_like_first_section": _cover_like(doc),
        "has_encryption": False,
        "field_count": (len(body.findall(".//" + qn("w:fldChar"))) +
                        len(body.findall(".//" + qn("w:fldSimple")))),
    }
    return analysis
