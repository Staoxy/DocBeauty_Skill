# -*- coding: utf-8 -*-
"""程序化 DOCX fixture 生成器（DESIGN_V2.md §24：不提交二进制样例）。"""
import base64
import io

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

# 1x1 红色像素 PNG
PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")


def _p(doc, text="", size=None, bold=False, align=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    if size:
        r.font.size = Pt(size)
    if bold:
        r.bold = True
    if align:
        p.alignment = align
    return p


def _body(doc, n=3, prefix="正文内容"):
    for k in range(n):
        _p(doc, f"{prefix}{k + 1}。这是一段用于测试的正文文字，包含足够长度的中英文 mix 内容 data {k + 1}，"
                f"用于验证字体、缩进与行距的统一效果。")


def basic_manual_headings_doc(path):
    """手工编号标题 + 正文 + 连续空段 + 图注（T-guard / T-reconcile / 幂等）。"""
    doc = Document()
    _p(doc, "课程论文标题", size=22, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph("")
    doc.add_paragraph("")
    _p(doc, "一、绪论", size=16, bold=True)
    _body(doc, 2)
    _p(doc, "（一）研究背景", size=14, bold=True)
    _body(doc, 2)
    _p(doc, "1. 国内研究现状", size=12, bold=True)
    _body(doc, 1)
    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph("")
    _p(doc, "二、方法", size=16, bold=True)
    _body(doc, 1)
    _p(doc, "图1 实验流程示意图")
    _p(doc, "参考文献")
    _p(doc, "[1] 张三. 论文写作研究[J]. 某学报, 2023.")
    _p(doc, "[2] Li S. Paper writing[J]. Journal, 2022.")
    doc.save(path)
    return path


def emphasis_doc(path):
    """粗/斜/下划线/颜色强调 run（T-emphasis）。"""
    doc = Document()
    p = doc.add_paragraph()
    r1 = p.add_run("普通文字")
    r2 = p.add_run("加粗强调")
    r2.bold = True
    r3 = p.add_run("斜体强调")
    r3.italic = True
    r4 = p.add_run("下划线强调")
    r4.underline = True
    r5 = p.add_run("红色文字")
    r5.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    _p(doc, "一、标题", bold=True, size=16)
    _body(doc, 1)
    doc.save(path)
    return path


def hyperlink_doc(path):
    """含超链接段落（T-hyperlink）。"""
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("前文内容。")
    _add_hyperlink(p, "https://example.com", "这是一个超链接文本")
    p.add_run("。后文内容。")
    _p(doc, "一、标题", bold=True, size=16)
    _body(doc, 1)
    doc.save(path)
    return path


def _add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def two_tables_doc(path):
    """两表夹空段（T-table-merge）。"""
    doc = Document()
    _p(doc, "一、数据", bold=True, size=16)
    t1 = doc.add_table(rows=2, cols=2)
    t1.cell(0, 0).text = "表A列1"
    t1.cell(0, 1).text = "表A列2"
    t1.cell(1, 0).text = "1"
    t1.cell(1, 1).text = "2"
    doc.add_paragraph("")  # 两表之间的空段：删除会导致表格合并
    t2 = doc.add_table(rows=1, cols=2)
    t2.cell(0, 0).text = "表B列1"
    t2.cell(0, 1).text = "表B列2"
    _body(doc, 1)
    doc.save(path)
    return path


def sectpr_blank_doc(path):
    """多节文档：节尾空段带 sectPr（T-sectpr）。"""
    doc = Document()
    _p(doc, "封面标题", size=22, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_section()  # 产生带 sectPr 的空段
    _p(doc, "一、正文标题", bold=True, size=16)
    _body(doc, 2)
    doc.save(path)
    return path


def image_doc(path):
    """超宽 inline 图片（T-image）。"""
    doc = Document()
    _p(doc, "一、图片章节", bold=True, size=16)
    _body(doc, 1)
    doc.add_picture(io.BytesIO(PNG_1PX), width=Inches(9))  # 超出 A4 可用宽度
    _p(doc, "图1 测试图片")
    _body(doc, 1)
    doc.save(path)
    return path


def existing_toc_field_doc(path):
    """已有 TOC 域（§16 步骤1：不插入第二个域）。"""
    doc = Document()
    p = doc.add_paragraph()
    r1 = OxmlElement("w:r")
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "begin")
    r1.append(fc)
    r2 = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.text = ' TOC \\o "1-3" \\h \\z \\u '
    r2.append(it)
    r3 = OxmlElement("w:r")
    fc2 = OxmlElement("w:fldChar")
    fc2.set(qn("w:fldCharType"), "end")
    r3.append(fc2)
    for r in (r1, r2, r3):
        p._p.append(r)
    doc.add_heading("一、已有标题", level=1)
    _body(doc, 2)
    doc.save(path)
    return path


def manual_toc_doc(path):
    """手工文字目录（T-manual-toc：保留 + 警告）。"""
    doc = Document()
    _p(doc, "目录", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _p(doc, "一、绪论..........1")
    _p(doc, "二、方法..........5")
    _p(doc, "三、结果..........9")
    _p(doc, "四、结论..........12")
    _p(doc, "一、绪论", bold=True, size=16)
    _body(doc, 2)
    _p(doc, "二、方法", bold=True, size=16)
    _body(doc, 2)
    doc.save(path)
    return path


def numbered_heading_style_doc(path):
    """Heading 1 样式绑定编号 + 手工编号文字（T-double-numbering）。"""
    doc = Document()
    h1 = doc.styles["Heading 1"].element
    pPr = h1.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        h1.append(pPr)
    numPr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    numId = OxmlElement("w:numId")
    numId.set(qn("w:val"), "1")
    numPr.append(ilvl)
    numPr.append(numId)
    pPr.append(numPr)
    _p(doc, "一、手工编号标题", bold=True, size=16)  # 手工编号（内容，不可动）
    _body(doc, 2)
    doc.save(path)
    return path


def tracked_changes_doc(path):
    """含未接受修订（w:ins）与批注引用（T-tracking）。"""
    doc = Document()
    _p(doc, "一、标题", bold=True, size=16)
    p = doc.add_paragraph()
    r = p.add_run("原始文字。")
    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"), "1")
    ins.set(qn("w:author"), "导师")
    ins.set(qn("w:date"), "2024-01-01T00:00:00Z")
    r2 = OxmlElement("w:r")
    t2 = OxmlElement("w:t")
    t2.text = "插入的修订文字。"
    r2.append(t2)
    ins.append(r2)
    p._p.append(ins)
    _body(doc, 1)
    doc.save(path)
    return path


def textbox_doc(path):
    """含文本框（边界对象：未处理亦未破坏）。"""
    doc = Document()
    _p(doc, "一、标题", bold=True, size=16)
    p = doc.add_paragraph()
    r = OxmlElement("w:r")
    pict = OxmlElement("w:pict")
    from docx.oxml import parse_xml
    V = "urn:schemas-microsoft-com:vml"
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    pict.append(parse_xml(
        f'<v:shape xmlns:v="{V}" xmlns:w="{W}" style="width:100pt;height:40pt">'
        f'<v:textbox><w:txbxContent><w:p><w:r><w:t>文本框内部文字</w:t></w:r></w:p>'
        f'</w:txbxContent></v:textbox></v:shape>'))
    r.append(pict)
    p._p.append(r)
    _body(doc, 1)
    doc.save(path)
    return path


def empty_doc(path):
    Document().save(path)
    return path


def corrupt_doc(path):
    with open(path, "wb") as f:
        f.write(b"this is not a zip file at all")
    return path


# ---------------------------------------------------------------------------
# V1.5 模板 fixtures（DESIGN_V15 Phase 1）
# ---------------------------------------------------------------------------
def _style_font(doc, style_name, east=None, ascii_name=None, size_pt=None):
    """设置样式的中英文字体与字号（python-docx 只设 ascii，必须补 eastAsia）。"""
    st = doc.styles[style_name].element
    rPr = st.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        st.append(rPr)
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    if east:
        rf.set(qn("w:eastAsia"), east)
    if ascii_name:
        rf.set(qn("w:ascii"), ascii_name)
        rf.set(qn("w:hAnsi"), ascii_name)
    if size_pt:
        sz = rPr.find(qn("w:sz"))
        if sz is None:
            sz = OxmlElement("w:sz")
            rPr.append(sz)
        sz.set(qn("w:val"), str(int(size_pt * 2)))


def _run_east(run, east, size_pt=None):
    rPr = run._p.find(qn("w:rPr")) if False else run._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    rf.set(qn("w:eastAsia"), east)
    if size_pt:
        run.font.size = Pt(size_pt)


def _footer_page_field(section, east="宋体", size_pt=9):
    p = section.footer.paragraphs[0]
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), r" PAGE \* MERGEFORMAT ")
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:eastAsia"), east)
    rPr.append(rf)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size_pt * 2)))
    rPr.append(sz)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    p._p.append(fld)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def regular_template_doc(path):
    """规范型模板：styles.xml 有真样式定义（DESIGN_V15 §6 regular 路径）。"""
    doc = Document()
    for sec in doc.sections:
        sec.page_width = Cm(21.0); sec.page_height = Cm(29.7)
        sec.top_margin = Cm(2.5); sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(3.0); sec.right_margin = Cm(2.5)
    _style_font(doc, "Normal", east="仿宋", ascii_name="Georgia", size_pt=12)
    _style_font(doc, "Heading 1", east="黑体", ascii_name="Georgia", size_pt=16)
    _style_font(doc, "Heading 2", east="黑体", ascii_name="Georgia", size_pt=14)
    doc.add_heading("示范论文标题", level=0)
    doc.add_heading("一、绪论", level=1)
    doc.add_paragraph("这是模板正文示范段落，展示正文字体与段落格式的要求。")
    doc.add_heading("（一）研究背景", level=2)
    doc.add_paragraph("第二段示范正文，用于稳定样式使用统计。")
    sec = doc.sections[0]
    sec.header.paragraphs[0].text = "XX大学课程论文"
    _footer_page_field(sec)
    doc.save(path)
    return path


def sample_template_doc(path):
    """样例型模板：格式全在直接格式化里（DESIGN_V15 §6 sample 路径）。"""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0); sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.5); sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(3.0); sec.right_margin = Cm(2.5)
    t = doc.add_paragraph()
    tr = t.add_run("示范论文标题")
    tr.bold = True
    _run_east(tr, "黑体", 22)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1 = doc.add_paragraph()
    h1r = h1.add_run("一、绪论")
    _run_east(h1r, "黑体", 16)
    h1r.bold = True
    b1 = doc.add_paragraph()
    b1r = b1.add_run("模板正文示范段落，展示正文格式要求，包含足够长度用于众数统计。")
    _run_east(b1r, "宋体", 12)
    b1.paragraph_format.line_spacing = 1.5
    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLineChars"), "200")
    ind.set(qn("w:firstLine"), "480")
    b1._p.get_or_add_pPr().append(ind)
    h2 = doc.add_paragraph()
    h2r = h2.add_run("（一）研究背景")
    _run_east(h2r, "黑体", 14)
    b2 = doc.add_paragraph()
    b2r = b2.add_run("第二段示范正文，同样应用正文格式，保证众数稳定。")
    _run_east(b2r, "宋体", 12)
    b2.paragraph_format.line_spacing = 1.5
    ind2 = OxmlElement("w:ind")
    ind2.set(qn("w:firstLineChars"), "200")
    ind2.set(qn("w:firstLine"), "480")
    b2._p.get_or_add_pPr().append(ind2)
    cap = doc.add_paragraph()
    capr = cap.add_run("图1 示范图题")
    _run_east(capr, "宋体", 10.5)
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sec.header.paragraphs[0].text = "XX大学课程论文"
    _footer_page_field(sec)
    doc.save(path)
    return path


def messy_paper_doc(path):
    """乱格式目标文档（从 test_realistic_messy_paper 抽取复用）。"""
    doc = Document()
    t = doc.add_paragraph()
    tr = t.add_run("基于深度学习的图像识别研究")
    tr.font.size = Pt(22); tr.bold = True
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1 = doc.add_paragraph()
    h1r = h1.add_run("一、绪论"); h1r.font.size = Pt(15); h1r.font.name = "Calibri"
    b1 = doc.add_paragraph()
    b1r = b1.add_run("近年来，深度学习在图像识别领域取得了显著进展，本文对此进行了系统研究与分析。")
    b1r.font.name = "Arial"
    b1.paragraph_format.line_spacing = 1.0
    b2 = doc.add_paragraph()
    b2.add_run("相关研究见")
    _add_hyperlink(b2, "https://example.org/paper", "文献综述")
    b2.add_run("。本文提出的方法在多个数据集上验证了有效性。")
    h2 = doc.add_paragraph()
    h2r = h2.add_run("（一）研究背景")
    h2r.font.size = Pt(13)
    h2r.bold = True
    b3 = doc.add_paragraph()
    b3.add_run("图像识别技术的发展可以追溯到上世纪六十年代，经历了多个重要阶段。")
    tbl = doc.add_table(rows=3, cols=3)
    data = [["方法", "准确率", "耗时"], ["CNN", "92.5%", "3.2h"], ["本文方法", "95.1%", "2.8h"]]
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            tbl.cell(i, j).text = val
    cap = doc.add_paragraph()
    cap.add_run("图1 方法对比结果")
    b4 = doc.add_paragraph()
    b4.add_run("实验结果表明，本文方法在保持效率的同时显著提升了识别精度。")
    h1b = doc.add_paragraph()
    h1br = h1b.add_run("二、结论")
    h1br.font.size = Pt(15)
    h1br.font.name = "Calibri"
    b5 = doc.add_paragraph()
    b5.add_run("本文提出了面向图像识别的改进方法，实验验证了其有效性与实用性。")
    doc.save(path)
    return path
