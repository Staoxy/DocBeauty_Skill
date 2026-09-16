# -*- coding: utf-8 -*-
"""V1.5 Phase 4 测试：audit 差异报告 + plan 子命令。"""
import json
import os
import sys

from config import load_config  # noqa: E402
from main import main as cli_main, process_document  # noqa: E402

from tests import fixtures as F  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def test_audit_detects_mismatch(tmp_path, capsys):
    target = F.messy_paper_doc(tmp_path / "target.docx")
    tpl = F.regular_template_doc(tmp_path / "tpl.docx")
    code = cli_main(["audit", str(target), "--template", str(tpl),
                     "-o", str(tmp_path / "out")])
    rep = json.loads(capsys.readouterr().out)
    assert code == 0
    assert rep["audit"]["result"] == "ATTENTION"
    checks = rep["audit"]["checks"]
    # 乱格式目标是 Letter、正文 Arial、无页码、无页眉 → 与模板不一致
    assert checks["page_size"]["match"] is False
    assert checks["body_font"]["match"] is False
    assert checks["page_number"]["match"] is False
    assert checks["header_footer"]["match"] is False
    # 模板无表格 → 表格检查跳过为 match
    assert checks["table_borders"]["match"] is True


def test_audit_pass_after_apply(tmp_path):
    """audit(排版前)=ATTENTION → apply → audit(排版后)=PASS（闭环）。"""
    import subprocess
    target = F.messy_paper_doc(tmp_path / "target.docx")
    tpl = F.regular_template_doc(tmp_path / "tpl.docx")

    def _audit_stdout(path):
        r = subprocess.run([sys.executable, "main.py", "audit", str(path),
                            "--template", str(tpl), "-o", str(tmp_path / "a")],
                           capture_output=True, encoding="utf-8", cwd=ROOT)
        return json.loads(r.stdout)

    before = _audit_stdout(target)
    assert before["audit"]["result"] == "ATTENTION"

    cfg = load_config(None)
    cfg["mode"] = "template"
    cfg["template"] = str(tpl)
    cfg["output_dir"] = str(tmp_path / "out")
    rep, code, out = process_document(str(target), cfg)
    assert code == 0, rep["errors"]

    after = _audit_stdout(out)
    assert after["audit"]["result"] == "PASS", after["audit"]["checks"]
    for cid in ("page_size", "body_font", "heading_h1", "page_number",
                "header_footer"):
        assert after["audit"]["checks"][cid]["match"] is True, cid


def test_plan_writes_formatting_spec(tmp_path, capsys):
    target = F.messy_paper_doc(tmp_path / "target.docx")
    tpl = F.regular_template_doc(tmp_path / "tpl.docx")
    code = cli_main(["plan", str(target), "--template", str(tpl),
                     "-o", str(tmp_path / "out")])
    out = capsys.readouterr().out
    assert code == 0
    rep = json.loads(out)
    assert rep["status"] == "dry_run"
    assert rep["role_map"]["body"] == "Normal"
    # formatting_spec.json 已生成
    spec_path = tmp_path / "out" / "formatting_spec.json"
    assert spec_path.is_file()
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    assert spec["mode"] == "template"
    assert spec["params"]["font"]["chinese"] == "仿宋"
    assert spec["sources"]["font.chinese"] == "P2:template"
    assert spec["policy"]["numbering"] == "preserve_target"
    # plan 不产出 docx
    assert not (tmp_path / "out" / "target_beautified.docx").exists()


def test_plan_with_intent(tmp_path, capsys):
    target = F.messy_paper_doc(tmp_path / "target.docx")
    intent = json.dumps({"style": "formal", "preferences": {"line_density": "loose"}})
    code = cli_main(["plan", str(target), "--intent", intent,
                     "-o", str(tmp_path / "out")])
    out = capsys.readouterr().out
    assert code == 0
    rep = json.loads(out)
    assert rep["config"]["style_resolved"] == "formal"
    spec = json.loads((tmp_path / "out" / "formatting_spec.json").read_text(encoding="utf-8"))
    assert spec["params"]["paragraph"]["line_spacing"] == 1.75
    assert spec["sources"]["paragraph.line_spacing"] == "P1:intent"
