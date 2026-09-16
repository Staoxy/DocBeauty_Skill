# -*- coding: utf-8 -*-
"""Audit（DESIGN_V15 §13）：目标文档 vs 模板的差异清单，只读不改。

输出 ✓/✗ 检查表（PASS / ATTENTION），供 Agent 展示"apply 会改什么"。
"""
import analyzer
import style_analyzer
import template_analyzer


def _fmt_role(info: dict, keys=("eastAsia", "ascii", "size_pt")) -> dict:
    return {k: info.get(k) for k in keys if info.get(k) is not None}


def run_audit(target_doc, template_spec: dict) -> dict:
    ta = analyzer.analyze(target_doc)
    tsa = style_analyzer.analyze_styles(target_doc)
    checks = {}

    sec0 = target_doc.sections[0] if target_doc.sections else None

    # ---- page_size ----
    t_size = template_analyzer._page_size_label(
        sec0.page_width, sec0.page_height) if sec0 else None
    tpl_size = (template_spec.get("page") or {}).get("size")
    checks["page_size"] = {"target": t_size, "template": tpl_size,
                           "match": t_size == tpl_size}

    # ---- margins ----
    t_m = template_analyzer._margins_cm(sec0) if sec0 else {}
    p_m = (template_spec.get("page") or {}).get("margins_cm") or {}
    if p_m:
        diffs = [f"{k}: {t_m.get(k)} vs {v}" for k, v in p_m.items()
                 if t_m.get(k) is not None and abs(t_m[k] - v) > 0.05]
        checks["margins"] = {"match": not diffs, "detail": diffs or None}
    else:
        checks["margins"] = {"match": True, "detail": "template unspecified"}

    # ---- body_font（模板指定的字段逐一比对）----
    body_tpl = (template_spec.get("roles") or {}).get("body") or {}
    if body_tpl:
        normal = tsa["styles"].get("Normal") or {}
        dd = tsa["doc_defaults"]
        t_body = {"eastAsia": normal.get("eastAsia") or dd.get("eastAsia"),
                  "ascii": normal.get("ascii") or dd.get("ascii"),
                  "size_pt": normal.get("size_pt") or dd.get("size_pt")}
        diffs = [f"{k}: {t_body.get(k)} vs {v}" for k, v in body_tpl.items()
                 if t_body.get(k) != v]
        checks["body_font"] = {"target": t_body, "template": body_tpl,
                               "match": not diffs, "detail": diffs or None}
    else:
        checks["body_font"] = {"match": True, "detail": "template unspecified"}

    # ---- headings h1 ----
    h1_tpl = (template_spec.get("roles") or {}).get("h1") or {}
    if h1_tpl:
        h1_info, _ = style_analyzer.style_by_role(tsa, "h1")
        t_h1 = _fmt_role(h1_info or {})
        diffs = [f"{k}: {t_h1.get(k)} vs {v}" for k, v in h1_tpl.items()
                 if k in t_h1 and t_h1.get(k) != v]
        checks["heading_h1"] = {"target": t_h1, "template": h1_tpl,
                                "match": not diffs, "detail": diffs or None}
    else:
        checks["heading_h1"] = {"match": True, "detail": "template unspecified"}

    # ---- table_borders ----
    tpl_tables = template_spec.get("tables") or {}
    if tpl_tables.get("borders"):
        from tables import _top_level_tables
        tbls = _top_level_tables(target_doc)
        if tbls:
            actual = template_analyzer._classify_table_borders(tbls[0])
            match = actual.get("borders") == tpl_tables["borders"]
            checks["table_borders"] = {
                "target": actual.get("borders"), "template": tpl_tables["borders"],
                "match": match}
        else:
            checks["table_borders"] = {"match": True, "detail": "target has no tables"}
    else:
        checks["table_borders"] = {"match": True, "detail": "template unspecified"}

    # ---- page_number ----
    tpl_pn = (template_spec.get("header_footer") or {}).get("footer_page_number") or {}
    if tpl_pn.get("present"):
        t_has = bool(ta.get("existing_page_number_fields"))
        checks["page_number"] = {"target": t_has, "template": True, "match": t_has}
    else:
        checks["page_number"] = {"match": True, "detail": "template unspecified"}

    # ---- header_footer 文字 ----
    tpl_header = (template_spec.get("header_footer") or {}).get("header") or {}
    if tpl_header.get("text"):
        try:
            t_header = "\n".join(p.text.strip() for p in
                                 target_doc.sections[0].header.paragraphs
                                 if p.text.strip()) or None
        except Exception:
            t_header = None
        checks["header_footer"] = {"target": t_header, "template": tpl_header["text"],
                                   "match": t_header == tpl_header["text"]}
    else:
        checks["header_footer"] = {"match": True, "detail": "template unspecified"}

    all_match = all(c.get("match", True) for c in checks.values())
    return {
        "result": "PASS" if all_match else "ATTENTION",
        "template_kind": template_spec.get("template_kind"),
        "checks": checks,
    }
