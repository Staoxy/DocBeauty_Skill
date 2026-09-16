# -*- coding: utf-8 -*-
"""Quality Checker（DESIGN_V2.md §18）：Q01–Q16。

输入上下文（ctx）：
  before_analysis / after_analysis: analyzer.analyze() 结果
  accounting: {blank_deleted, toc_paragraphs_added, ...}
  requested: {"toc": bool, "page_number": bool, "toc_requested": bool, ...}
  zones_headings_after: list[int] 输出文档的标题层级序列
"""
from docx.oxml.ns import qn

from analyzer import heading_level_of, para_text
from layout import usable_width_emu
from structure import is_blank_paragraph
from tables import _grid_cols, _top_level_tables, _EMU_PER_TWIP
from utils import ooxml as ox

_W_T = qn("w:t")


def _heading_levels(doc) -> list:
    styles_elem = doc.styles.element
    levels = []
    for p in doc.paragraphs:
        lvl = heading_level_of(styles_elem, p._p)
        if lvl in (1, 2, 3):
            levels.append(lvl)
    return levels


def _body_run_font_ratio(doc):
    total = east = 0
    for p in doc.paragraphs:
        for r in p._p.findall(qn("w:r")):
            txt = "".join(t.text or "" for t in r.iter(_W_T)).strip()
            if not txt:
                continue
            rPr = ox.get_child(r, "w:rPr")
            if rPr is None:
                continue
            rf = ox.get_child(rPr, "w:rFonts")
            total += 1
            if rf is not None and rf.get(qn("w:eastAsia")):
                east += 1
    return (east / total) if total else 1.0, total


def _indent_ratio(doc):
    total = ok = 0
    styles_elem = doc.styles.element
    for i, p in enumerate(doc.paragraphs):
        if not para_text(p._p).strip():
            continue
        lvl = heading_level_of(styles_elem, p._p)
        if lvl in (1, 2, 3):
            continue
        pPr = p._p.find(qn("w:pPr"))
        ind = ox.get_child(pPr, "w:ind") if pPr is not None else None
        if ind is not None and ind.get(qn("w:firstLineChars")) == "200":
            ok += 1
        total += 1
    return (ok / total) if total else 1.0, total


def _spacing_ratio(doc):
    from collections import Counter
    hist = Counter()
    for p in doc.paragraphs:
        pPr = p._p.find(qn("w:pPr"))
        sp = ox.get_child(pPr, "w:spacing") if pPr is not None else None
        if sp is not None and sp.get(qn("w:line")):
            hist[sp.get(qn("w:line"))] += 1
    total = sum(hist.values())
    if not total:
        return 1.0
    return hist.most_common(1)[0][1] / total


def _blank_clusters(doc) -> int:
    worst = cur = 0
    for p in doc.paragraphs:
        if is_blank_paragraph(p._p):
            cur += 1
            worst = max(worst, cur)
        else:
            cur = 0
    return worst


def run_checks(doc, ctx: dict) -> list:
    """返回 [{id, status: pass|warn|fail|info|skip, detail}]。"""
    results = []

    def add(qid, ok, detail, level="warn"):
        status = "pass" if ok else ("fail" if level == "error" else "warn" if level == "warn" else "info")
        results.append({"id": qid, "status": status, "detail": detail})

    ba, aa = ctx.get("before_analysis", {}), ctx.get("after_analysis", {})
    acc = ctx.get("accounting", {})
    req = ctx.get("requested", {})

    # Q03/Q04/Q05/Q06/Q07 计数类
    for qid, key, level, cmp in (
        ("Q03", "table_count", "error", "eq"),
        ("Q04", "inline_image_count", "error", "eq"),
        ("Q05", "floating_image_count", "error", "eq"),
        ("Q06", "hyperlink_count", "warn", "ge"),
    ):
        b, a = ba.get(key, 0), aa.get(key, 0)
        ok = (a == b) if cmp == "eq" else (a >= b)
        add(qid, ok, f"{key}: {b} -> {a}", level)

    # Q07 域/书签不减（用 guard integrity 更可靠，此处独立校验）
    b_fields = ba.get("field_count", 0)
    a_fields = _count_fields(doc)
    add("Q07", a_fields >= b_fields,
        f"fields: {b_fields} -> {a_fields}", "warn")

    # Q02 段落对账（§18 公式）
    b_paras = ba.get("paragraph_count", 0)
    a_paras = aa.get("paragraph_count", 0)
    expected = b_paras - acc.get("blank_deleted", 0) + acc.get("toc_paragraphs_added", 0) \
        - acc.get("manual_toc_removed", 0)
    add("Q02", a_paras == expected,
        f"paragraphs: {b_paras} -> {a_paras} (expected {expected})", "error")

    # Q08 标题层级
    levels = _heading_levels(doc)
    if levels:
        h1_ok = 1 in levels
        jumps = [(i, a, b) for i, (a, b) in enumerate(zip(levels, levels[1:]))
                 if b - a > 1]
        add("Q08", h1_ok and not jumps,
            f"headings={len(levels)}, h1_present={h1_ok}, jumps={jumps[:3]}", "warn")
    else:
        add("Q08", not req.get("format_headings"),
            "no headings detected/applied", "warn")

    # Q09 表格宽度（逐表 vs 可用宽度）
    usable = usable_width_emu(doc, ctx.get("effective", {})) // _EMU_PER_TWIP
    overflow = []
    for i, tbl in enumerate(_top_level_tables(doc)):
        total = sum(_grid_cols(tbl))
        if total > usable * 1.02:  # 2% 容差
            overflow.append(i)
    add("Q09", not overflow,
        f"tables exceeding usable width: {overflow}" if overflow else
        f"all {len(_top_level_tables(doc))} tables fit (usable={usable} dxa)", "warn")

    # Q10 残留空段簇（>2）
    worst = _blank_clusters(doc)
    add("Q10", worst <= 2, f"max consecutive blank paragraphs: {worst}", "warn")

    # Q11 页码
    if req.get("page_number"):
        has_page = bool(aa.get("existing_page_number_fields"))
        add("Q11", has_page, "PAGE field present in headers/footers", "warn")
    else:
        results.append({"id": "Q11", "status": "skip", "detail": "not requested"})

    # Q12 目录
    if req.get("toc"):
        has_field = _has_toc_field(doc)
        uf = ox.get_child(doc.settings.element, "w:updateFields")
        uf_on = uf is not None and uf.get(qn("w:val")) in (None, "true", "1", "on")
        add("Q12", has_field and uf_on,
            f"toc_field={has_field}, updateFields={uf_on}", "warn")
    else:
        results.append({"id": "Q12", "status": "skip", "detail": "not requested"})

    # Q13 字体覆盖
    ratio, total = _body_run_font_ratio(doc)
    add("Q13", ratio >= 0.95, f"eastAsia set on {ratio:.1%} of {total} body runs", "warn")

    # Q14 缩进一致（clean 模板跳过）
    if ctx.get("effective", {}).get("paragraph", {}).get("first_line_chars", 2):
        ir, n = _indent_ratio(doc)
        add("Q14", ir >= 0.90, f"firstLineChars=200 on {ir:.1%} of {n} body paragraphs", "info")
    else:
        results.append({"id": "Q14", "status": "skip", "detail": "clean template: no indent"})

    # Q15 行距一致
    sr = _spacing_ratio(doc)
    add("Q15", sr >= 0.90, f"dominant line spacing on {sr:.1%} of paragraphs", "info")

    # Q16 兼容红线（附录 A）
    bad_switch = [it.text for it in doc.element.body.iter(qn("w:instrText"))
                  if it.text and r"\* decimal" in it.text]
    for hf_part in _hf_parts(doc):
        for it in hf_part.iter(qn("w:instrText")):
            if it.text and r"\* decimal" in it.text:
                bad_switch.append(it.text)
    empty_pg = 0
    for el in doc.element.body.iter(qn("w:pgNumType")):
        if not el.attrib:
            empty_pg += 1
    sec_ok = ba.get("section_count") == aa.get("section_count")
    add("Q16", not bad_switch and not empty_pg and sec_ok,
        f"decimal_switch={bad_switch[:1]}, empty_pgNumType={empty_pg}, "
        f"sections {ba.get('section_count')}->{aa.get('section_count')}", "warn")

    return results


def _count_fields(doc):
    body = doc.element.body
    return (len(body.findall(".//" + qn("w:fldChar"))) +
            len(body.findall(".//" + qn("w:fldSimple"))))


def _has_toc_field(doc):
    for it in doc.element.body.iter(qn("w:instrText")):
        if it.text and "TOC" in it.text:
            return True
    for fs in doc.element.body.iter(qn("w:fldSimple")):
        if "TOC" in (fs.get(qn("w:instr")) or ""):
            return True
    return False


def _hf_parts(doc):
    parts = []
    seen = set()
    for section in doc.sections:
        for hf in (section.header, section.footer):
            try:
                part = hf.part
            except Exception:
                continue
            if part is not None and id(part) not in seen:
                seen.add(id(part))
                parts.append(part.element)
    return parts
