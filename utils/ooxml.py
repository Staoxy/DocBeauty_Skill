# -*- coding: utf-8 -*-
"""OOXML 操作工具层（DESIGN_V2.md 附录 C 的实现）。

安全守则（§22）：
- 所有 XML 元素经本模块创建，命名空间一律经 qn()；
- 元素插入遵守 OOXML schema 顺序（TAG_SEQ）。
"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# ---------------------------------------------------------------------------
# 常用父元素的子元素 schema 顺序（插入排序用，只列会用到的关键段）
# ---------------------------------------------------------------------------
TAG_SEQ = {
    # 键/值统一使用本地名（无前缀），与 _local() 输出一致
    "p": ["pPr"],
    "r": ["rPr"],
    "tbl": ["tblPr", "tblGrid", "tr"],
    "tr": ["trPr", "tc"],
    "tc": ["tcPr", "p"],
    "pPr": [
        "pStyle", "keepNext", "keepLines", "pageBreakBefore",
        "framePr", "widowControl", "numPr", "suppressLineNumbers",
        "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku",
        "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
        "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
        "spacing", "ind", "contextualSpacing", "mirrorIndents",
        "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
        "rPr", "sectPr", "pPrChange",
    ],
    "rPr": [
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps",
        "smallCaps", "strike", "dstrike", "outline", "shadow",
        "emboss", "imprint", "noProof", "snapToGrid", "vanish",
        "webHidden", "color", "spacing", "w", "kern", "position",
        "sz", "szCs", "highlight", "u", "effect", "bdr",
        "shd", "fitText", "vertAlign", "rtl", "cs", "em",
        "lang", "eastAsianLayout", "specVanish", "oMath",
    ],
    "tblPr": [
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
        "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
        "tblCellSpacing", "tblInd", "tblBorders", "shd",
        "tblLayout", "tblCellMar", "tblLook", "tblCaption",
        "tblDescription",
    ],
    "trPr": [
        "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore",
        "wAfter", "cantSplit", "trHeight", "tblHeader",
        "tblCellSpacing", "jc", "hidden", "ins", "del",
        "trPrChange",
    ],
    "tcPr": [
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge",
        "tcBorders", "shd", "noWrap", "tcMar", "textDirection",
        "tcFitText", "vAlign", "hideMark", "headers",
    ],
    "tblBorders": [
        "top", "left", "start", "bottom", "right", "end",
        "insideH", "insideV",
    ],
    "tcBorders": [
        "top", "left", "start", "bottom", "right", "end",
        "insideH", "insideV", "tl2br", "tr2bl",
    ],
    "sectPr": [
        "headerReference", "footerReference", "footnotePr",
        "endnotePr", "type", "pgSz", "pgMar", "paperSrc",
        "pgBorders", "lnNumType", "pgNumType", "cols", "formProt",
        "vAlign", "noEndnote", "titlePg", "textDirection", "bidi",
        "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
    ],
    "style": [
        "name", "aliases", "basedOn", "next", "link",
        "autoRedefine", "hidden", "uiPriority", "semiHidden",
        "unhideWhenUsed", "qFormat", "locked", "personal",
        "personalCompose", "personalReply", "cs", "rsid",
        "pPr", "rPr", "tblPr", "trPr", "tcPr",
    ],
}


def make_elem(tag: str, attrs: dict = None):
    """创建带属性的 OOXML 元素。"""
    e = OxmlElement(tag)
    if attrs:
        for k, v in attrs.items():
            e.set(qn(k) if ":" in k else k, str(v))
    return e


def _local(tag: str) -> str:
    """统一取本地名：'w:jc' -> 'jc'，'{ns}jc' -> 'jc'。"""
    if "}" in tag:
        return tag.split("}")[-1]
    return tag.split(":")[-1]


def insert_ordered(parent, child):
    """按 schema 顺序把 child 插入 parent（TAG_SEQ）。

    未知标签视为排在已知标签之后（如 w:p 中的 w:r 排在 w:pPr 后），
    保证 pPr/rPr/trPr/tcPr 始终位于父元素首位（schema 要求）。
    """
    ptag = _local(parent.tag)
    seq = TAG_SEQ.get(ptag)
    if not seq:
        parent.append(child)
        return child
    ctag = _local(child.tag)
    idx = seq.index(ctag) if ctag in seq else len(seq)
    for existing in parent:
        etag = _local(existing.tag)
        eidx = seq.index(etag) if etag in seq else len(seq)
        if eidx > idx:
            existing.addprevious(child)
            return child
    parent.append(child)
    return child


def sub(parent, tag: str, attrs: dict = None):
    """创建子元素并按 schema 顺序插入。若已存在同名子元素则返回它（单例语义）。"""
    ptag = _local(parent.tag)
    seq = TAG_SEQ.get(ptag, [])
    ctag = _local(tag)
    # 单例元素：已存在则复用
    if ctag in seq:
        for existing in parent:
            if _local(existing.tag) == ctag:
                if attrs:
                    for k, v in attrs.items():
                        existing.set(qn(k) if ":" in k else k, str(v))
                return existing
    return insert_ordered(parent, make_elem(tag, attrs))


def get_child(parent, tag: str):
    if parent is None:
        return None
    ctag = _local(tag)
    for existing in parent:
        if _local(existing.tag) == ctag:
            return existing
    return None


def remove_child(parent, tag: str):
    ctag = _local(tag)
    removed = False
    for existing in list(parent):
        if _local(existing.tag) == ctag:
            parent.remove(existing)
            removed = True
    return removed


# ---------------------------------------------------------------------------
# run 级：字体（附录 C-1）
# ---------------------------------------------------------------------------
THEME_FONT_ATTRS = [
    "w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme",
]


def set_run_font(rPr, west: str, east: str, size_pt: float = None):
    """设置 rFonts 四件套（ascii/hAnsi/eastAsia/cs）并清除主题字体引用。"""
    rFonts = sub(rPr, "w:rFonts")
    for a in THEME_FONT_ATTRS:
        if rFonts.get(qn(a)) is not None:
            del rFonts.attrib[qn(a)]
    rFonts.set(qn("w:ascii"), west)
    rFonts.set(qn("w:hAnsi"), west)
    rFonts.set(qn("w:eastAsia"), east)
    rFonts.set(qn("w:cs"), west)
    if size_pt is not None:
        set_run_size(rPr, size_pt)


def set_run_size(rPr, size_pt: float):
    half = int(round(size_pt * 2))
    sz = sub(rPr, "w:sz", {"w:val": str(half)})
    sz.set(qn("w:val"), str(half))
    szCs = sub(rPr, "w:szCs")
    szCs.set(qn("w:val"), str(half))


def clear_direct_font(rPr):
    """清除 run 残留直接字体格式（字体名/字号/主题引用），保留强调属性。

    keep_inline_emphasis=True（默认）时保留：b/bCs/i/iCs/u/color/vertAlign/
    highlight/strike/dstrike。本函数只清 rFonts/sz/szCs。
    """
    remove_child(rPr, "w:rFonts")
    remove_child(rPr, "w:sz")
    remove_child(rPr, "w:szCs")


# ---------------------------------------------------------------------------
# 段落级：缩进 / 行距 / 对齐 / 孤行控制（附录 C-2/C-3/C-4）
# ---------------------------------------------------------------------------
def pt_to_twip(v: float) -> int:
    return int(round(v * 20))


def set_first_line_indent(pPr, chars: int = 2, body_size_pt: float = 12):
    """首行缩进 N 字符：firstLineChars 优先 + twips 回退（字号变化不漂移）。"""
    ind = sub(pPr, "w:ind")
    ind.set(qn("w:firstLineChars"), str(int(chars * 100)))
    ind.set(qn("w:firstLine"), str(pt_to_twip(chars * body_size_pt)))
    # 悬挂缩进与首行缩进互斥
    if ind.get(qn("w:hangingChars")) is not None:
        del ind.attrib[qn("w:hangingChars")]
    if ind.get(qn("w:hanging")) is not None:
        del ind.attrib[qn("w:hanging")]


def set_line_spacing(pPr, multiple: float):
    """多倍行距：1.5 -> line=360 lineRule=auto。"""
    spacing = sub(pPr, "w:spacing")
    spacing.set(qn("w:line"), str(int(round(multiple * 240))))
    spacing.set(qn("w:lineRule"), "auto")
    # 清除固定/最小行距规则残留
    if spacing.get(qn("w:lineRule")) == "exact":
        spacing.set(qn("w:lineRule"), "auto")


def set_spacing_before_after(pPr, before_pt: float, after_pt: float):
    spacing = sub(pPr, "w:spacing")
    spacing.set(qn("w:before"), str(pt_to_twip(before_pt)))
    spacing.set(qn("w:after"), str(pt_to_twip(after_pt)))
    # 不用 beforeLines/afterLines（渲染器支持差）
    for a in ("w:beforeLines", "w:afterLines"):
        if spacing.get(qn(a)) is not None:
            del spacing.attrib[qn(a)]


ALIGN_VAL = {
    "left": "left", "center": "center", "right": "right",
    "both": "both", "justify": "both", "distribute": "distribute",
}


def set_alignment(pPr, align: str):
    jc = sub(pPr, "w:jc")
    jc.set(qn("w:val"), ALIGN_VAL.get(align, "both"))


def set_keep_with_next(pPr, keep: bool = True):
    if keep:
        sub(pPr, "w:keepNext")
    else:
        remove_child(pPr, "w:keepNext")


def set_keep_lines(pPr, keep: bool = True):
    if keep:
        sub(pPr, "w:keepLines")
    else:
        remove_child(pPr, "w:keepLines")


def remove_para_numbering(pPr):
    """移除段落级自动编号 numPr（编号文字属于内容，绝不动）。"""
    return remove_child(pPr, "w:numPr")


# ---------------------------------------------------------------------------
# 域（附录 C-5/C-6）
# ---------------------------------------------------------------------------
def make_run(text: str = None):
    r = make_elem("w:r")
    if text is not None:
        t = make_elem("w:t")
        t.text = text
        t.set(_XML_SPACE, "preserve")
        r.append(t)
    return r


def run_with_font(text: str, west: str, east: str, size_pt: float,
                  bold=False, italic=False, color=None):
    r = make_run(text)
    rPr = insert_ordered(r, make_elem("w:rPr"))
    set_run_font(rPr, west, east, size_pt)
    if bold:
        sub(rPr, "w:b")
    if italic:
        sub(rPr, "w:i")
    if color:
        sub(rPr, "w:color", {"w:val": color})
    return r


def fld_char_run(fld_type: str):
    r = make_run()
    r.append(make_elem("w:fldChar", {"w:fldCharType": fld_type}))
    return r


def instr_text_run(instr: str):
    r = make_run()
    it = make_elem("w:instrText")
    it.text = instr
    it.set(_XML_SPACE, "preserve")
    r.append(it)
    return r


def make_page_field(west: str, east: str, size_pt: float,
                    instr: str = " PAGE \\* MERGEFORMAT ") -> "lxml element":
    """页码域（fldSimple）。instr 绝不允许 \\* decimal（附录 A-1）。"""
    fld = make_elem("w:fldSimple", {"w:instr": instr})
    fld.append(run_with_font("1", west, east, size_pt))
    return fld


def make_toc_field_runs(instr: str, placeholder: str, west: str, east: str,
                        size_pt: float):
    """TOC 域三段结构（begin / instrText / separate / 占位 / end）的 run 列表。"""
    runs = [
        fld_char_run("begin"),
        instr_text_run(instr),
        fld_char_run("separate"),
        run_with_font(placeholder, west, east, size_pt, italic=True,
                      color="808080"),
        fld_char_run("end"),
    ]
    return runs


def make_page_break_run():
    r = make_run()
    r.append(make_elem("w:br", {"w:type": "page"}))
    return r


def set_update_fields(settings_elem) -> bool:
    """settings.xml 写入 updateFields=true（§16 步骤 5）。返回是否有变化。"""
    existing = get_child(settings_elem, "w:updateFields")
    if existing is not None:
        if existing.get(qn("w:val")) in (None, "true", "1", "on"):
            existing.set(qn("w:val"), "true")
            return False
        existing.set(qn("w:val"), "true")
        return True
    e = make_elem("w:updateFields", {"w:val": "true"})
    settings_elem.insert(0, e)
    return True


# ---------------------------------------------------------------------------
# 样式（§10.3 / D1 双轨制）
# ---------------------------------------------------------------------------
BUILTIN_HEADING_NAMES = {
    1: "heading 1", 2: "heading 2", 3: "heading 3",
}


def find_style_by_name(styles_elem, name: str):
    """按 w:name 查找样式（不区分大小写）。"""
    low = name.strip().lower()
    for st in styles_elem.findall(qn("w:style")):
        nm = get_child(st, "w:name")
        if nm is not None and (nm.get(qn("w:val")) or "").strip().lower() == low:
            return st
    return None


def _unique_style_id(styles_elem, base: str) -> str:
    """返回未被占用的 styleId（base / base2 / base3 ...）。"""
    if find_style_by_sid(styles_elem, base) is None:
        return base
    n = 2
    while find_style_by_sid(styles_elem, f"{base}{n}") is not None:
        n += 1
    return f"{base}{n}"


def ensure_heading_style(doc, level: int):
    """确保 Heading N 样式存在，返回 styleId。

    兼容 w:name="heading N"（含中文 Word 文档）与 styleId="HeadingN"/"N"。
    不存在时创建：basedOn Normal + outlineLvl。
    """
    from docx.oxml.ns import qn as _qn
    styles_elem = doc.styles.element
    st = find_style_by_name(styles_elem, BUILTIN_HEADING_NAMES[level])
    if st is not None:
        sid = st.get(_qn("w:styleId"))
        if sid:
            return sid
        sid = f"Heading{level}"
        st.set(_qn("w:styleId"), sid)
        return sid
    # 创建
    sid = f"Heading{level}"
    if find_style_by_sid(styles_elem, sid) is None:
        # styleId 为纯数字（"1"/"2"/"3"）的旧文档/WPS 文档可复用，但**必须**
        # 确认该 id 真的是同名标题样式。否则会撞上 WPS 里 styleId="1" 的
        # 「列表段落1」之类，标题段落被套成列表样式（WPS 生成文档实测必现）。
        alt = find_style_by_sid(styles_elem, str(level))
        alt_name = alt.find(_qn("w:name")) if alt is not None else None
        if alt_name is not None and \
                (alt_name.get(_qn("w:val")) or "").strip().lower() == BUILTIN_HEADING_NAMES[level]:
            sid = str(level)
        else:
            sid = _unique_style_id(styles_elem, sid)
    st = make_elem("w:style", {
        "w:type": "paragraph", "w:styleId": sid,
    })
    insert_ordered(st, make_elem("w:name", {"w:val": BUILTIN_HEADING_NAMES[level]}))
    insert_ordered(st, make_elem("w:basedOn", {"w:val": _normal_style_id(doc)}))
    insert_ordered(st, make_elem("w:next", {"w:val": _normal_style_id(doc)}))
    insert_ordered(st, make_elem("w:qFormat"))
    pPr = insert_ordered(st, make_elem("w:pPr"))
    insert_ordered(pPr, make_elem("w:outlineLvl", {"w:val": str(level - 1)}))
    styles_elem.append(st)
    return sid


def find_style_by_sid(styles_elem, sid: str):
    from docx.oxml.ns import qn as _qn
    for st in styles_elem.findall(_qn("w:style")):
        if st.get(_qn("w:styleId")) == sid:
            return st
    return None


def _normal_style_id(doc) -> str:
    st = find_style_by_name(doc.styles.element, "Normal")
    return st.get(qn("w:styleId")) if st is not None else "Normal"


def style_ppr(st):
    pPr = get_child(st, "w:pPr")
    return pPr if pPr is not None else insert_ordered(st, make_elem("w:pPr"))


def style_rpr(st):
    rPr = get_child(st, "w:rPr")
    return rPr if rPr is not None else insert_ordered(st, make_elem("w:rPr"))


def disable_style_numbering(st) -> bool:
    """在样式 pPr 写 numPr numId=0 禁用继承编号（防双重编号，附录 C-9）。"""
    pPr = style_ppr(st)
    existing = get_child(pPr, "w:numPr")
    if existing is not None:
        # 已经禁用？
        numId = get_child(existing, "w:numId")
        if numId is not None and numId.get(qn("w:val")) == "0":
            return False
        pPr.remove(existing)
    numPr = insert_ordered(pPr, make_elem("w:numPr"))
    insert_ordered(numPr, make_elem("w:ilvl", {"w:val": "0"}))
    insert_ordered(numPr, make_elem("w:numId", {"w:val": "0"}))
    return True


# ---------------------------------------------------------------------------
# 表格（§13）
# ---------------------------------------------------------------------------
BORDER_PT_TO_EIGHTH = lambda pt: str(int(round(pt * 8)))  # 1.5pt -> 12


def set_tbl_width(tblPr, dxa: int):
    tblW = sub(tblPr, "w:tblW", {"w:w": str(dxa), "w:type": "dxa"})
    tblW.set(qn("w:w"), str(dxa))
    tblW.set(qn("w:type"), "dxa")


def set_tbl_layout_fixed(tblPr):
    sub(tblPr, "w:tblLayout", {"w:type": "fixed"})


def _border(parent, tag, val, sz_eighth=None, color="auto"):
    attrs = {"w:val": val}
    if val != "none":
        attrs["w:sz"] = str(sz_eighth or 4)
        attrs["w:space"] = "0"
        attrs["w:color"] = color
    sub(parent, tag, attrs)


def set_tbl_borders(tblPr, mode: str):
    """设置表级边框。threeline: 顶/底 1.5pt + 表头行底线（行级另行处理）。"""
    if mode == "none":
        remove_child(tblPr, "w:tblBorders")
        return
    borders = sub(tblPr, "w:tblBorders")
    for tag in ("w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV"):
        remove_child(borders, tag)
    if mode == "threeline":
        _border(borders, "w:top", "single", 12)
        _border(borders, "w:bottom", "single", 12)
        for tag in ("w:left", "w:right", "w:insideH", "w:insideV"):
            _border(borders, tag, "none")
    elif mode == "grid":
        for tag in ("w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV"):
            _border(borders, tag, "single", 4)


def set_row_cant_split(tr):
    trPr = get_child(tr, "w:trPr")
    if trPr is None:
        trPr = insert_ordered(tr, make_elem("w:trPr"))
    sub(trPr, "w:cantSplit")


def set_row_header_repeat(tr):
    trPr = get_child(tr, "w:trPr")
    if trPr is None:
        trPr = insert_ordered(tr, make_elem("w:trPr"))
    sub(trPr, "w:tblHeader")


def set_cell_valign_center(tc):
    tcPr = get_child(tc, "w:tcPr")
    if tcPr is None:
        tcPr = insert_ordered(tc, make_elem("w:tcPr"))
    sub(tcPr, "w:vAlign", {"w:val": "center"})


def set_cell_width(tc, dxa: int):
    tcPr = get_child(tc, "w:tcPr")
    if tcPr is None:
        tcPr = insert_ordered(tc, make_elem("w:tcPr"))
    tcW = sub(tcPr, "w:tcW", {"w:w": str(dxa), "w:type": "dxa"})
    tcW.set(qn("w:w"), str(dxa))
    tcW.set(qn("w:type"), "dxa")


def set_cell_bottom_border(tc, sz_eighth=6):
    """表头行底线（三线表的栏目线）。"""
    tcPr = get_child(tc, "w:tcPr")
    if tcPr is None:
        tcPr = insert_ordered(tc, make_elem("w:tcPr"))
    borders = sub(tcPr, "w:tcBorders")
    _border(borders, "w:bottom", "single", sz_eighth)


def has_merged_cells(tbl) -> bool:
    return (tbl.findall(".//" + qn("w:gridSpan")) or
            tbl.findall(".//" + qn("w:vMerge"))) is not None and bool(
        tbl.findall(".//" + qn("w:gridSpan")) or
        tbl.findall(".//" + qn("w:vMerge")))


# ---------------------------------------------------------------------------
# 节/页面（§11）
# ---------------------------------------------------------------------------
def iter_body_block_items(document):
    """按文档顺序产出 body 的直接子元素。"""
    body = document.element.body
    yield from body


def element_section_index(document):
    """返回两个 list：
    - elements: body 直接子元素列表
    - sec_idx: 与 elements 等长，每个元素所属节索引（0-based）。
    段落含 pPr/sectPr 表示该节结束；body 尾部的 sectPr 属于最后一节。
    """
    elements = list(document.element.body)
    sec_idx = []
    current = 0
    for el in elements:
        sec_idx.append(current)
        if el.tag == qn("w:p"):
            pPr = el.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                current += 1
        elif el.tag == qn("w:sectPr"):
            pass  # 最后一节的 sectPr
    return elements, sec_idx
