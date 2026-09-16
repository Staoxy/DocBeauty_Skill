# -*- coding: utf-8 -*-
"""JSON 报告组装与写出（DESIGN_V2.md §19）。"""
import json
import os
import platform
import sys
import time

import constants as C
from importlib.metadata import version as _pkg_version


def _versions():
    out = {"skill": C.SKILL_VERSION, "python": platform.python_version()}
    for pkg in ("python-docx", "lxml", "jsonschema"):
        try:
            out[pkg.replace("-", "_")] = _pkg_version(pkg)
        except Exception:
            out[pkg.replace("-", "_")] = "unknown"
    return out


def build_report(status: str, exit_code: int, sections: dict) -> dict:
    rep = {
        "status": status,
        "exit_code": exit_code,
        "version": _versions(),
        "input": sections.get("input", {}),
        "output": sections.get("output", {}),
        "config": sections.get("config", {}),
        "analysis": sections.get("analysis", {}),
        "zones": sections.get("zones", {}),
        "statistics": sections.get("statistics", {}),
        "operations_detail": sections.get("operations_detail", []),
        "content_guard": sections.get("content_guard", {}),
        "quality_checks": sections.get("quality_checks", []),
        "diagnosis": sections.get("diagnosis", {}),
        "render_check": sections.get("render_check", {}),
        "template_spec": sections.get("template_spec", {}),
        "role_map": sections.get("role_map", {}),
        "formatting_spec": sections.get("formatting_spec", {}),
        "intent_resolved": sections.get("intent_resolved", {}),
        "untouched_boundary": sections.get("untouched_boundary", {}),
        "warnings": sections.get("warnings", []),
        "errors": sections.get("errors", []),
        "timing": sections.get("timing", {}),
    }
    return rep


def write_report(report: dict, path: str):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)


def emit_stdout(report: dict):
    """stdout 仅输出报告 JSON（§21.1）。"""
    json.dump(report, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")
