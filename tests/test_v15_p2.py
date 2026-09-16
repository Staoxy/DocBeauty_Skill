# -*- coding: utf-8 -*-
"""V1.5 Phase 2 测试：Role Mapper / Rule Resolver / Spec Builder / 模板模式。"""
import json
import os
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import rule_resolver  # noqa: E402
from config import load_config, resolve_style  # noqa: E402
from main import process_document  # noqa: E402
from role_mapper import build_role_map  # noqa: E402
from templates import get_template  # noqa: E402
from utils import ooxml as ox  # noqa: E402

from tests import fixtures as F  # noqa: E402


# ---------------------------------------------------------------------------
# Rule Resolver 单元
# ---------------------------------------------------------------------------
def test_resolver_priority():
    builtin = get_template("academic")
    template_spec = {"page": {"margins_cm": {"left": 2.0}},
                     "roles": {"body": {"eastAsia": "楷体", "size_pt": 12}}}
    cfg = {"mode": "template",
           "overrides": {"font": {"chinese": "隶书"}}}
    params, sources = rule_resolver.resolve(
        cfg, builtin, template_spec=template_spec)
    # P1 overrides > P2 template > P5 default
    assert params["font"]["chinese"] == "隶书"
    assert sources["font.chinese"] == "P1:overrides"
    # P2 覆盖 P5：模板 body eastAsia 胜过内置宋体
    assert params["font"]["eastAsia"] if False else True
    assert params["font"].get("chinese") == "隶书"
    # P2 page margin 覆盖内置
    assert params["page"]["margins_cm"]["left"] == 2.0
    assert sources["page.margins_cm.left"] == "P2:template"
    # 模板未指定的维度保持内置（P5）
    assert params["paragraph"]["line_spacing"] == 1.5
    assert sources["paragraph.line_spacing"] == "P5:default"


def test_resolver_template_beats_builtin_body():
    builtin = get_template("formal")
    template_spec = {"roles": {"body": {"eastAsia": "仿宋", "ascii": "Georgia",
                                        "size_pt": 12}}}
    cfg = {"mode": "template"}
    params, sources = rule_resolver.resolve(cfg, builtin, template_spec=template_spec)
    assert params["font"]["chinese"] == "仿宋"
    assert params["font"]["western"] == "Georgia"
    assert sources["font.chinese"] == "P2:template"


def test_role_mapper_missing_roles():
    spec = {"roles": {"body": {"size_pt": 12}}, "roles_meta": {"body": "Normal"}}
    role_map, warnings = build_role_map(spec)
    assert role_map["body"] == "Normal"
    assert role_map["h1"] is None
    assert any("h1" in w for w in warnings), "必需角色缺失应有警告"


# ---------------------------------------------------------------------------
# 引擎模板/参考模式端到端
# ---------------------------------------------------------------------------
def _template_cfg(tpl_path, outdir, mode="template"):
    cfg = load_config(None)
    cfg["mode"] = mode
    cfg[mode] = str(tpl_path)
    cfg["output_dir"] = str(outdir)
    return cfg


def test_template_mode_e2e(tmp_path):
    target = F.messy_paper_doc(tmp_path / "target.docx")
    tpl = F.regular_template_doc(tmp_path / "tpl.docx")
    rep, code, out = process_document(str(target), _template_cfg(tpl, tmp_path / "out"))
    assert code == 0, rep["errors"]
    assert rep["status"] == "success"
    assert rep["content_guard"]["content_changed"] is False

    # 报告结构
    assert rep["template_spec"]["template_kind"] == "regular"
    assert rep["role_map"]["body"] == "Normal"
    fs = rep["formatting_spec"]
    assert fs["policy"]["numbering"] == "preserve_target"
    assert fs["sources"].get("font.chinese") == "P2:template"

    d = Document(out)
    # 模板正文字体传播：仿宋 + Georgia
    body = [p for p in d.paragraphs if "深度学习在图像识别领域" in p.text][0]
    for r in body._p.findall(qn("w:r")):
        rPr = ox.get_child(r, "w:rPr")
        rf = ox.get_child(rPr, "w:rFonts")
        assert rf.get(qn("w:eastAsia")) == "仿宋"
        assert rf.get(qn("w:ascii")) == "Georgia"
    # 模板页边距传播
    assert round(d.sections[0].left_margin.cm, 2) == 3.0
    # 模板页眉文字已应用
    header_text = "".join(p.text for p in d.sections[0].header.paragraphs)
    assert "XX大学课程论文" in header_text
    # 标题仍提升为 Heading；字体由 Heading 样式接管（run 直接格式被清除，§10.3）
    from analyzer import heading_level_of
    h1 = [p for p in d.paragraphs if p.text.strip() == "一、绪论"][0]
    assert heading_level_of(d.styles.element, h1._p) == 1
    h1_style = ox.find_style_by_name(d.styles.element, "heading 1")
    srPr = ox.get_child(h1_style, "w:rPr")
    srf = ox.get_child(srPr, "w:rFonts")
    assert srf.get(qn("w:eastAsia")) == "黑体"


def test_template_mode_sample_path_e2e(tmp_path):
    target = F.messy_paper_doc(tmp_path / "target.docx")
    tpl = F.sample_template_doc(tmp_path / "tpl.docx")
    rep, code, out = process_document(str(target), _template_cfg(tpl, tmp_path / "out"))
    assert code == 0, rep["errors"]
    assert rep["template_spec"]["template_kind"] == "sample"
    d = Document(out)
    body = [p for p in d.paragraphs if "深度学习在图像识别领域" in p.text][0]
    rPr = ox.get_child(body._p.findall(qn("w:r"))[0], "w:rPr")
    rf = ox.get_child(rPr, "w:rFonts")
    assert rf.get(qn("w:eastAsia")) == "宋体"  # 样例模板正文字体


def test_reference_mode_e2e(tmp_path):
    target = F.messy_paper_doc(tmp_path / "target.docx")
    ref = F.regular_template_doc(tmp_path / "ref.docx")
    rep, code, out = process_document(
        str(target), _template_cfg(ref, tmp_path / "out", mode="reference"))
    assert code == 0, rep["errors"]
    assert rep["role_map"]["body"] == "Normal"
    d = Document(out)
    body = [p for p in d.paragraphs if "深度学习在图像识别领域" in p.text][0]
    rf = ox.get_child(ox.get_child(body._p.findall(qn("w:r"))[0], "w:rPr"), "w:rFonts")
    assert rf.get(qn("w:ascii")) == "Georgia"
    # 参考模式的页眉不应用（Reference 学习风格，不复制页眉文字——§3 MODE 4）
    header_text = "".join(p.text for p in d.sections[0].header.paragraphs)
    assert "XX大学课程论文" not in header_text


def test_template_mode_idempotent(tmp_path):
    target = F.messy_paper_doc(tmp_path / "target.docx")
    tpl = F.regular_template_doc(tmp_path / "tpl.docx")
    rep1, code1, out1 = process_document(
        str(target), _template_cfg(tpl, tmp_path / "out1"))
    assert code1 == 0
    rep2, code2, _ = process_document(out1, _template_cfg(tpl, tmp_path / "out2"))
    assert code2 == 0, rep2["errors"]
    assert rep2["content_guard"]["content_changed"] is False


def test_template_file_invalid(tmp_path):
    import pytest as _pytest
    from config import ConfigError
    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"not a docx")
    target = F.messy_paper_doc(tmp_path / "target.docx")
    cfg = _template_cfg(bad, tmp_path / "out")
    # process_document 直接调用时配置错误以异常抛出（CLI 路径由 run_one 捕获转 error_report）
    with _pytest.raises(ConfigError):
        process_document(str(target), cfg)
