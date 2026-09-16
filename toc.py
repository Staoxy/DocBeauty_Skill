# -*- coding: utf-8 -*-
"""TOC（DESIGN_V2.md §16）：Word 原生 TOC 域 + updateFields + 分页。"""
from docx.oxml.ns import qn

import constants as C
from utils import ooxml as ox


def _toc_field_paragraphs(doc):
    """已有 TOC 域所在段落索引列表。"""
    from analyzer import para_text  # noqa: F401
    idxs = []
    for i, p in enumerate(doc.paragraphs):
        hit = False
        for it in p._p.iter(qn("w:instrText")):
            if it.text and "TOC" in it.text:
                hit = True
                break
        if not hit:
            for fs in p._p.iter(qn("w:fldSimple")):
                if "TOC" in (fs.get(qn("w:instr")) or ""):
                    hit = True
                    break
        if hit:
            idxs.append(i)
    return idxs


def toc_conditions(zones, analysis: dict):
    """§16 量化条件。"""
    h = analysis.get("heading_counts", {})
    total = h.get("h1", 0) + h.get("h2", 0) + h.get("h3", 0) + len(
        [1 for i, lv in zones.headings.items()])
    return total, h.get("h1", 0) + len([1 for i, lv in zones.headings.items() if lv == 1])


def add_toc(doc, zones, effective: dict, options: dict, analysis: dict,
            accounting: dict) -> dict:
    detail = {"status": "skipped", "reason": None, "warnings": []}

    # 已有 TOC 域 -> 只补 updateFields（§16 步骤1）
    if analysis.get("existing_toc_field") or _toc_field_paragraphs(doc):
        changed = ox.set_update_fields(doc.settings.element)
        detail["status"] = "updated_existing"
        detail["update_fields_set"] = True
        return detail

    # 手工目录 -> 默认不删除，仅警告（§16 步骤2）
    manual = zones.zones.get("toc_existing_manual")
    if manual and not options.get("toc_replace_manual", False):
        detail["reason"] = "manual_toc_present"
        detail["warnings"].append(
            f"检测到手工文字目录（段落 {manual[0]}~{manual[1]}），默认保留；"
            "删除后重跑可生成原生目录，或设置 options.toc_replace_manual=true")
        return detail

    total_headings, h1_count = toc_conditions(zones, analysis)
    if total_headings < 3 or h1_count < 1:
        detail["reason"] = f"heading structure unclear (headings={total_headings}, h1={h1_count})"
        return detail

    # 幂等：TOC 域所在段（或其前一段）已是"目录"标题（上次运行产出）→ 只补 updateFields
    from analyzer import para_text
    idxs = _toc_field_paragraphs(doc)
    if idxs and any(
            para_text(doc.paragraphs[i]._p).strip() == C.TOC_TITLE
            for i in range(max(0, min(idxs) - 3), min(min(idxs) + 1, len(doc.paragraphs)))):
        ox.set_update_fields(doc.settings.element)
        detail["status"] = "updated_existing"
        return detail

    # 插入位置（§16 步骤3）
    anchor_el = None
    front = zones.zones.get("front_matter")
    if front:
        anchor_idx = front[1] + 1
        if anchor_idx < len(doc.paragraphs):
            anchor_el = doc.paragraphs[anchor_idx]._p
    if anchor_el is None and zones.headings:
        anchor_el = doc.paragraphs[min(zones.headings.keys())]._p
    if anchor_el is None:
        detail["reason"] = "no insertion anchor"
        return detail
    # doc.paragraphs 每次返回新包装对象，必须用底层 XML 元素定位
    anchor_idx = next(i for i, p in enumerate(doc.paragraphs) if p._p is anchor_el)

    pn = effective["page_number"]
    west, east = pn["font_western"], pn["font_cn"]
    body = effective["font"]

    # ① 目录标题（直接格式，不用 Heading 1，避免目录收录自身）
    title_p = doc.paragraphs[anchor_idx].insert_paragraph_before()
    r = ox.run_with_font(C.TOC_TITLE, effective["headings"]["h1"]["font_western"],
                         effective["headings"]["h1"]["font_cn"],
                         effective["headings"]["h1"]["size_pt"], bold=True)
    title_p._p.append(r)
    from formatter import get_or_add_ppr
    ox.set_alignment(get_or_add_ppr(title_p._p), "center")

    # ② TOC 域（三段 fldChar + 占位文本）；levels 可被模板/覆盖指定（§7 toc）
    levels = (effective.get("toc") or {}).get("levels", "1-3")
    instr = f' TOC \\o "{levels}" \\h \\z \\u '
    field_p = doc.paragraphs[anchor_idx + 1].insert_paragraph_before()
    for run in ox.make_toc_field_runs(instr, C.TOC_PLACEHOLDER,
                                      west, east, body["size_pt"]):
        field_p._p.append(run)

    # ③ 提示行 + ④ 分页（提示段落尾部，避免多余空段）
    tail_p = field_p
    if options.get("toc_refresh_hint", True):
        hint_p = doc.paragraphs[anchor_idx + 2].insert_paragraph_before()
        hint_p._p.append(ox.run_with_font(
            C.TOC_HINT, west, east, 9, italic=True, color="808080"))
        tail_p = hint_p
    if options.get("toc_page_break", True):
        tail_p._p.append(ox.make_page_break_run())

    # ⑤ settings.xml updateFields（否则目录打开是空的）
    ox.set_update_fields(doc.settings.element)

    added = 2 + (1 if options.get("toc_refresh_hint", True) else 0)
    accounting["toc_paragraphs_added"] = accounting.get("toc_paragraphs_added", 0) + added
    accounting.setdefault("excluded_generated_texts", []).extend(
        [C.TOC_TITLE, C.TOC_PLACEHOLDER, C.TOC_HINT])
    zones.shift(anchor_idx, added)

    detail["status"] = "added"
    detail["paragraphs_added"] = added
    detail["inserted_before_paragraph"] = anchor_idx + added
    return detail
