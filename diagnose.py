# -*- coding: utf-8 -*-
"""格式诊断（DESIGN_V2.md §3.4 task=diagnose）：只诊断、不修改。

输出结构化诊断项（diagnosis.issues），每项含 severity / evidence，
供 Agent 向用户展示"这份文档哪里乱"以及 beautify 将修复什么。
借鉴 document-format-skills 的诊断思路，但不改任何内容。
"""
import re
from collections import Counter

from docx.oxml.ns import qn

from analyzer import heading_level_of, para_text
from layout import usable_width_emu
from structure import is_blank_paragraph, para_is_caption
from tables import _EMU_PER_TWIP, _grid_cols, _top_level_tables
from utils import ooxml as ox

_W_R = qn("w:r")


def _body_font_profile(doc):
    """正文 run 的字体/字号分布（排除标题样式段落）。"""
    styles_elem = doc.styles.element
    fonts, sizes = Counter(), Counter()
    for p in doc.paragraphs:
        if heading_level_of(styles_elem, p._p) in (1, 2, 3):
            continue
        if para_is_caption(para_text(p._p)):
            continue
        for r in p._p.findall(_W_R):
            txt = "".join(t.text or "" for t in r.iter(qn("w:t"))).strip()
            if not txt:
                continue
            rPr = ox.get_child(r, "w:rPr")
            rf = ox.get_child(rPr, "w:rFonts") if rPr is not None else None
            # eastAsia 优先；缺失时回退 ascii（python-docx 的 font.name 只设 ascii，
            # 只设西文字体的中文 run 是常见混乱源，诊断必须能看到它）
            east = rf.get(qn("w:eastAsia")) if rf is not None else None
            font = east or (rf.get(qn("w:ascii")) if rf is not None else None)
            fonts[font or "(继承样式)"] += len(txt)
            sz = ox.get_child(rPr, "w:sz") if rPr is not None else None
            if sz is not None:
                try:
                    sizes[round(int(sz.get(qn("w:val"))) / 2, 1)] += len(txt)
                except (TypeError, ValueError):
                    pass
            else:
                sizes["(默认)"] += len(txt)
    return fonts, sizes


def _indent_stats(doc):
    total = missing = 0
    styles_elem = doc.styles.element
    for p in doc.paragraphs:
        text = para_text(p._p).strip()
        if not text or heading_level_of(styles_elem, p._p) in (1, 2, 3):
            continue
        if para_is_caption(text) or text.startswith(("一、", "二、", "三、", "四、", "五、")):
            continue
        pPr = p._p.find(qn("w:pPr"))
        ind = ox.get_child(pPr, "w:ind") if pPr is not None else None
        has = ind is not None and ind.get(qn("w:firstLineChars")) not in (None, "0")
        has = has or (ind is not None and ind.get(qn("w:firstLine")) not in (None, "0"))
        total += 1
        if not has:
            missing += 1
    return total, missing


def _spacing_stats(doc):
    hist = Counter()
    for p in doc.paragraphs:
        pPr = p._p.find(qn("w:pPr"))
        sp = ox.get_child(pPr, "w:spacing") if pPr is not None else None
        if sp is not None and sp.get(qn("w:line")):
            rule = sp.get(qn("w:lineRule")) or "auto"
            hist[f"{sp.get(qn('w:line'))}/{rule}"] += 1
    return hist


def diagnose(doc, analysis: dict, zones) -> dict:
    """生成诊断结果。返回 {issues: [...], summary: {...}}。"""
    issues = []

    def add(severity, code, message, evidence=None):
        issues.append({"severity": severity, "code": code,
                       "message": message, "evidence": evidence or []})

    fonts, sizes = _body_font_profile(doc)

    # ---- 字体混乱 ----
    explicit = {k: v for k, v in fonts.items() if k != "(继承样式)"}
    if len(explicit) > 1:
        total = sum(explicit.values())
        dominant = max(explicit.values())
        add("high" if dominant / total < 0.8 else "medium", "FONT_CHAOS",
            f"正文混用 {len(explicit)} 种中文字体",
            [f"{k}: {v} 字符" for k, v in
             sorted(explicit.items(), key=lambda kv: -kv[1])[:5]])

    # ---- 字号混乱（同段落群多字号，且不是标题字号梯度）----
    explicit_sizes = {k: v for k, v in sizes.items() if isinstance(k, float)}
    if len(explicit_sizes) > 3:
        add("medium", "SIZE_CHAOS", f"正文出现 {len(explicit_sizes)} 种字号",
            [f"{k}pt: {v} 字符" for k, v in
             sorted(explicit_sizes.items(), key=lambda kv: -kv[1])[:5]])

    # ---- 缩进缺失 ----
    total, missing = _indent_stats(doc)
    if total and missing / total >= 0.3:
        add("medium" if missing / total < 0.7 else "high", "INDENT_MISSING",
            f"{missing}/{total} 个正文段落缺少首行缩进")

    # ---- 行距混乱 ----
    spacing = _spacing_stats(doc)
    if len(spacing) > 1:
        add("medium", "SPACING_MIX", f"混用 {len(spacing)} 种行距设置",
            [f"{k}: {v} 段" for k, v in spacing.most_common(4)])

    # ---- 标题体系 ----
    style_headings = analysis.get("heading_counts", {})
    style_total = (style_headings.get("h1", 0) + style_headings.get("h2", 0)
                   + style_headings.get("h3", 0))
    promoted = len(zones.headings)
    if style_total == 0 and promoted == 0:
        add("high", "NO_HEADINGS", "未检测到任何标题样式或可用标题结构，无法生成目录")
    elif promoted > style_total:
        add("info", "MANUAL_HEADINGS",
            f"{promoted - style_total} 个手工编号标题将被提升为 Heading 样式",
            [c["text"] for c in zones.candidates[:5]])
    if zones.numbering_template:
        add("info", "NUMBERING_PATTERN",
            f"检测到编号模式: {zones.numbering_template}")

    # ---- 表格 ----
    try:
        usable = usable_width_emu(doc, {}) // _EMU_PER_TWIP
    except Exception:
        usable = 0
    for i, tbl in enumerate(_top_level_tables(doc)):
        cols = _grid_cols(tbl)
        if usable and sum(cols) > usable * 1.02:
            add("high", "TABLE_OVERFLOW", f"表格 {i} 超出页面可用宽度",
                [f"列宽合计 {sum(cols)} dxa > 可用 {usable} dxa"])
        if not tbl.findall(".//" + qn("w:tblHeader")):
            add("low", "TABLE_NO_HEADER_REPEAT", f"表格 {i} 未设置表头跨页重复")

    # ---- 页码/页面 ----
    if not analysis.get("existing_page_number_fields"):
        add("medium", "NO_PAGE_NUMBER", "未检测到页码域")
    for i, sec in enumerate(doc.sections):
        w = sec.page_width
        if w and abs(w - 11906) > 60:  # 非 A4 宽（±1mm）
            add("medium", "PAGE_SIZE", f"节 {i} 页面宽度非 A4",
                [f"width={w} EMU"])
            break

    # ---- 目录 ----
    if analysis.get("existing_toc_field"):
        add("info", "TOC_FIELD_EXISTS", "已存在 TOC 域（将只确保打开时自动更新）")
    elif zones.zones.get("toc_existing_manual"):
        rng = zones.zones["toc_existing_manual"]
        add("medium", "MANUAL_TOC", f"检测到手工文字目录（段落 {rng[0]}~{rng[1]}），默认保留")
    elif style_total + promoted >= 3:
        add("info", "TOC_CAN_ADD", "标题结构满足条件，可生成原生目录")

    # ---- 边界对象 / 协作痕迹 ----
    if analysis.get("floating_image_count"):
        add("info", "FLOATING_IMAGES",
            f"{analysis['floating_image_count']} 个浮动图片不参与处理（保留原样）")
    if analysis.get("textbox_count"):
        add("info", "TEXTBOXES", f"{analysis['textbox_count']} 个文本框不参与处理（保留原样）")
    tc = analysis.get("tracked_changes", {})
    if tc.get("ins") or tc.get("del"):
        add("medium", "TRACKED_CHANGES",
            f"存在未接受修订（ins={tc['ins']}, del={tc['del']}），建议先接受修订再排版")
    if analysis.get("comment_count"):
        add("low", "COMMENTS", f"包含 {analysis['comment_count']} 条批注（保留原样）")

    # ---- 空段 ----
    blanks = sum(1 for p in doc.paragraphs if is_blank_paragraph(p._p))
    if blanks >= 3:
        add("low", "BLANK_PARAGRAPHS", f"{blanks} 个空白段落（清理将只收敛连续空段）")

    severity_rank = {"high": 0, "medium": 1, "low": 2, "info": 3}
    issues.sort(key=lambda x: severity_rank.get(x["severity"], 9))
    summary = {
        "high": sum(1 for i in issues if i["severity"] == "high"),
        "medium": sum(1 for i in issues if i["severity"] == "medium"),
        "low": sum(1 for i in issues if i["severity"] == "low"),
        "info": sum(1 for i in issues if i["severity"] == "info"),
    }
    return {"issues": issues, "summary": summary}
