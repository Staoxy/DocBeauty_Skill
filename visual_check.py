# -*- coding: utf-8 -*-
"""Visual Check（DESIGN_V15 §14）：PDF 栅格化 + 客观版面指标。

分工（§2.2）：代码只算客观指标（墨水覆盖率/内容框/孤行），PNG 交给 Agent 目检。
不做审美评分——结果只有 PASS/WARNING/ERROR 级别的 issue 列表。

指标（§14）：
- blank          空白页：墨水覆盖率 < 0.5%                     ERROR
- sparse         大面积空白：正文页内容高度占比 < 30%（首页除外） WARNING
- out_of_margin  水平越界：内容框越过左右边距（垂直方向由 Word
                 自动分页截断，不检查——局限如实声明）            ERROR
- orphan_heading 标题孤行：标题行位于页面底部 10%                WARNING
"""
import os
from collections import Counter

import constants as C

_PT_PER_CM = 28.3465
_DPI = 72
_BLANK_INK = 0.005
_SPARSE_RATIO = 0.3


def _ink_coverage(pix) -> float:
    """非白像素占比（灰度近似：RGB 任一通道 < 245 记为着墨）。"""
    samples = pix.samples
    n = pix.width * pix.height
    if not n:
        return 0.0
    ink = 0
    step = pix.n
    data = samples
    # 每 4 像素抽样一次（72dpi 下已足够稳定，速度提升 4 倍）
    for i in range(0, n, 4):
        base = i * step
        if data[base] < 245 or data[base + 1] < 245 or data[base + 2] < 245:
            ink += 1
    return ink / (n / 4)


def _span_lines(page):
    d = page.get_text("dict")
    lines = []
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            text = "".join(s.get("text", "") for s in spans).strip()
            if not text:
                continue
            max_size = max(s.get("size", 0) for s in spans)
            bold = any("Bold" in (s.get("font") or "") for s in spans)
            lines.append({"bbox": line["bbox"], "text": text,
                          "size": max_size, "bold": bold})
    return lines


def check_pdf(pdf_path: str, margins_cm: dict = None,
              png_dir: str = None) -> dict:
    """返回 {available, pages:[{page,ink,content_height_ratio,png,flags}], issues, summary}。"""
    try:
        import pymupdf
    except ImportError:
        return {"available": False, "reason": "pymupdf not installed"}
    try:
        doc = pymupdf.open(pdf_path)
    except Exception as e:
        return {"available": False, "reason": f"cannot open pdf: {e}"}

    if png_dir:
        os.makedirs(png_dir, exist_ok=True)

    # 正文大小众数（孤行判定的字号基线）
    size_hist = Counter()
    all_lines = []
    for page in doc:
        lines = _span_lines(page)
        all_lines.append(lines)
        for ln in lines:
            size_hist[ln["size"]] += len(ln["text"])
    body_size = size_hist.most_common(1)[0][0] if size_hist else None

    margin_pt = None
    if margins_cm:
        try:
            margin_pt = {
                "left": margins_cm.get("left", 2.54) * _PT_PER_CM,
                "right": margins_cm.get("right", 2.54) * _PT_PER_CM,
            }
        except Exception:
            margin_pt = None

    pages, issues = [], []
    for pno, page in enumerate(doc):
        rect = page.rect
        flags = []

        pix = page.get_pixmap(dpi=_DPI)
        png_path = None
        if png_dir:
            png_path = os.path.join(png_dir, f"page_{pno + 1:03d}.png")
            pix.save(png_path)

        ink = _ink_coverage(pix)

        lines = all_lines[pno]
        content_h = 0.0
        x0 = x1 = y0 = y1 = None
        boxes = [ln["bbox"] for ln in lines]
        # 内容框必须合并矢量绘图/图片（越界表格/图形是真实风险）
        try:
            for dr in page.get_drawings():
                boxes.append(tuple(dr["rect"]))
        except Exception:
            pass
        if boxes:
            y0 = min(b[1] for b in boxes)
            y1 = max(b[3] for b in boxes)
            x0 = min(b[0] for b in boxes)
            x1 = max(b[2] for b in boxes)
            content_h = (y1 - y0) / rect.height if rect.height else 0.0

        if ink < _BLANK_INK:
            flags.append("blank")
            issues.append({"page": pno + 1, "severity": "error",
                           "code": "blank", "detail": f"ink={ink:.4f}"})
        elif lines and content_h < _SPARSE_RATIO and pno > 0:
            flags.append("sparse")
            issues.append({"page": pno + 1, "severity": "warning",
                           "code": "sparse",
                           "detail": f"content_height_ratio={content_h:.2f}"})

        if margin_pt and x0 is not None and rect.width:
            tol = 2.0
            if x0 < margin_pt["left"] - tol or x1 > rect.width - margin_pt["right"] + tol:
                flags.append("out_of_margin")
                issues.append({"page": pno + 1, "severity": "error",
                               "code": "out_of_margin",
                               "detail": f"x0={x0:.0f}, x1={x1:.0f}, "
                                         f"page_w={rect.width:.0f}"})

        # 标题孤行：页面最后一行像标题（字号 >= 正文+1 或加粗）且位于底部 10%
        if lines and body_size:
            last = max(lines, key=lambda ln: ln["bbox"][3])
            looks_heading = (last["size"] >= body_size + 1 or last["bold"]) \
                and len(last["text"]) <= 40 \
                and not last["text"].endswith(tuple(C.SENTENCE_END_PUNCT))
            if looks_heading and last["bbox"][3] > rect.height * 0.9:
                flags.append("orphan_heading")
                issues.append({"page": pno + 1, "severity": "warning",
                               "code": "orphan_heading",
                               "detail": last["text"][:30]})

        pages.append({"page": pno + 1, "ink": round(ink, 4),
                      "content_height_ratio": round(content_h, 3),
                      "png": png_path, "flags": flags})

    errors = sum(1 for i in issues if i["severity"] == "error")
    warns = sum(1 for i in issues if i["severity"] == "warning")
    return {
        "available": True,
        "engine": "pymupdf",
        "page_count": len(pages),
        "pages": pages,
        "issues": issues,
        "summary": {"errors": errors, "warnings": warns,
                    "result": "PASS" if not errors and not warns else
                              ("WARNING" if not errors else "ERROR")},
    }
