# -*- coding: utf-8 -*-
"""程序化 DOCX fixture 生成器（DESIGN_V2.md §24：不提交二进制样例）。"""
import base64
import io

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

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
