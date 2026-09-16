# -*- coding: utf-8 -*-
"""配置加载 / schema 校验 / overrides 深合并 / options 门控（DESIGN_V2.md §3.4/§9.7）。"""
import copy
import json
import os

import jsonschema

from constants import DEFAULT_OPTIONS, OPERATION_CATALOG
from templates import get_template

_SCHEMA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "schemas", "input_schema.json")


class ConfigError(Exception):
    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message


def load_schema() -> dict:
    with open(_SCHEMA_PATH, encoding="utf-8") as f:
        return json.load(f)


def load_config(source) -> dict:
    """source: dict | json 文件路径 | None（全默认）。返回原始配置（未合并模板）。"""
    if source is None:
        cfg = {}
    elif isinstance(source, dict):
        cfg = copy.deepcopy(source)
    elif isinstance(source, str):
        if not os.path.isfile(source):
            raise ConfigError("FILE_NOT_FOUND", f"config file not found: {source}")
        with open(source, encoding="utf-8") as f:
            cfg = json.load(f)
    else:
        raise ConfigError("CONFIG_INVALID", f"unsupported config source: {type(source)}")
    cfg.setdefault("task", "beautify")
    cfg.setdefault("document_type", "auto")
    cfg.setdefault("style", "auto")
    cfg.setdefault("content_protection", True)
    cfg.setdefault("operations", "auto")
    cfg.setdefault("format_file", None)
    cfg.setdefault("overrides", {})
    cfg.setdefault("options", {})
    cfg.setdefault("mode", "auto")
    cfg.setdefault("template", None)
    cfg.setdefault("reference", None)
    cfg.setdefault("intent", None)
    # 模式校验（DESIGN_V15 §3）
    if cfg["mode"] == "template" and not cfg["template"]:
        raise ConfigError("CONFIG_INVALID", "mode=template requires 'template' path")
    if cfg["mode"] == "reference" and not cfg["reference"]:
        raise ConfigError("CONFIG_INVALID", "mode=reference requires 'reference' path")
    # 校验
    try:
        jsonschema.validate(cfg, load_schema())
    except jsonschema.ValidationError as e:
        raise ConfigError("CONFIG_INVALID", f"config schema validation failed: {e.message}")
    # 未知操作名
    if isinstance(cfg["operations"], list):
        unknown = [op for op in cfg["operations"] if op not in OPERATION_CATALOG]
        if unknown:
            raise ConfigError("CONFIG_INVALID",
                              f"unknown operations: {unknown}; valid: {sorted(OPERATION_CATALOG)}")
    # 门控：punctuation_normalize 必须 content_protection=false（§10.5）
    opts = dict(DEFAULT_OPTIONS)
    opts.update(cfg["options"])
    if opts.get("punctuation_normalize") and cfg["content_protection"]:
        raise ConfigError(
            "CONFIG_INVALID",
            "options.punctuation_normalize=true requires content_protection=false "
            "(punctuation normalization modifies text content)")
    cfg["options"] = opts
    return cfg


def deep_merge(base: dict, override: dict) -> dict:
    """深合并：dict 递归、标量覆盖。返回新对象，不修改入参。"""
    out = copy.deepcopy(base)
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out


def resolve_style(cfg: dict, document_type: str) -> str:
    """style=auto 时按 §9.5 映射；显式指定时直接用。"""
    if cfg["style"] != "auto":
        return cfg["style"]
    return {"academic": "academic", "experiment": "academic"}.get(document_type, "formal")


def build_effective(cfg: dict, document_type: str) -> dict:
    """模板 + overrides 深合并 -> 生效参数（含回显用 overrides_applied）。"""
    style = resolve_style(cfg, document_type)
    tpl = copy.deepcopy(get_template(style))
    effective = deep_merge(tpl, cfg.get("overrides") or {})
    effective["_style_resolved"] = style
    return effective


def output_paths(cfg: dict, input_path: str, outdir_default: str = "./output"):
    """输出路径（§4）：保留原文件名 + _beautified 后缀。"""
    outdir = cfg.get("output_dir") or outdir_default
    os.makedirs(outdir, exist_ok=True)
    base = os.path.splitext(os.path.basename(input_path))[0]
    docx_path = os.path.join(outdir, f"{base}_beautified.docx")
    report_path = os.path.join(outdir, f"{base}_report.json")
    return docx_path, report_path
