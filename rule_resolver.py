# -*- coding: utf-8 -*-
"""Rule Resolver（DESIGN_V15 §10）：多来源排版参数的优先级解析。

层级（高覆盖低）：
  P1 用户显式：overrides（精确）+ intent 映射参数（语义）
  P2 模板规则（Template Spec.roles / page / header_footer / tables）
  P3 参考文档（Style Profile，来自 analyze_template 的客观测量）
  P4 目标文档现状（只影响 policy，不产生参数——见 §12.3）
  P5 内置默认（academic/formal/clean）

不变量层（内容保护）不在此处——它凌驾于一切优先级（§2.3），由 guard 强制。
每个最终叶子参数记录来源，报告可回溯"这个值是谁定的"。
"""
import copy


PRIORITY = {"P1:overrides": 4, "P1:intent": 3, "P2:template": 2,
            "P3:reference": 1, "P5:default": 0}


def _merge_layer(base: dict, layer: dict, source: str, sources: dict, prefix: str = ""):
    """layer 的叶子覆盖进 base，并记录每个叶子的最终来源。"""
    for k, v in (layer or {}).items():
        path = f"{prefix}.{k}" if prefix else str(k)
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            _merge_layer(base[k], v, source, sources, path)
        else:
            base[k] = copy.deepcopy(v)
            sources[path] = source


def _spec_to_layers(spec: dict, source: str) -> dict:
    """Template Spec → 参数层（与 effective config 同形状）。"""
    if not spec:
        return {}
    layer = {}
    if spec.get("page"):
        layer["page"] = {k: v for k, v in spec["page"].items() if v is not None}
    roles = spec.get("roles", {})
    mapped = {}
    if roles.get("body"):
        body = dict(roles["body"])
        font = {}
        for key in ("eastAsia", "ascii", "size_pt"):
            if body.get(key) is not None:
                font[{"eastAsia": "chinese", "ascii": "western",
                       "size_pt": "size_pt"}[key]] = body[key]
        if font:
            layer["font"] = font
        para = {}
        for key in ("line", "first_line_chars", "align", "before_pt", "after_pt"):
            if body.get(key) is not None:
                para[key] = body[key]
        if para:
            layer["paragraph"] = para
    for role, key in (("h1", "h1"), ("h2", "h2"), ("h3", "h3")):
        if roles.get(role):
            mapped[key] = _role_params_to_heading(roles[role])
    if mapped:
        layer["headings"] = mapped
    if roles.get("caption"):
        cap = {}
        for key, dst in (("eastAsia", "font_cn"), ("ascii", "font_western"),
                         ("size_pt", "size_pt"), ("align", "align")):
            if roles["caption"].get(key) is not None:
                cap[dst] = roles["caption"][key]
        if cap:
            layer["captions"] = cap
    if spec.get("tables"):
        layer["tables"] = {k: v for k, v in spec["tables"].items() if v is not None}
    fpn = (spec.get("header_footer") or {}).get("footer_page_number") or {}
    if fpn.get("present"):
        pn = {}
        for key, dst in (("align", "align"), ("size_pt", "size_pt"),
                         ("eastAsia", "font_cn")):
            if fpn.get(key) is not None:
                pn[dst] = fpn[key]
        if pn:
            layer["page_number"] = pn
    if spec.get("toc"):
        layer["toc"] = spec["toc"]
    del source
    return layer


def _role_params_to_heading(params: dict) -> dict:
    out = {}
    for key, dst in (("eastAsia", "font_cn"), ("ascii", "font_western"),
                     ("size_pt", "size_pt"), ("bold", "bold"),
                     ("align", "align"), ("before_pt", "before_pt"),
                     ("after_pt", "after_pt")):
        if params.get(key) is not None:
            out[dst] = params[key]
    return out


def resolve(cfg: dict, builtin_params: dict, template_spec=None,
            reference_profile=None, intent_params=None) -> tuple:
    """返回 (params, sources)。应用顺序：P5 → P3 → P2 → P1（高覆盖低）。"""
    params = copy.deepcopy(builtin_params)
    sources = {}

    # P5 内置默认
    def _record_default(d, prefix=""):
        for k, v in d.items():
            path = f"{prefix}.{k}" if prefix else str(k)
            if isinstance(v, dict):
                _record_default(v, path)
            else:
                sources.setdefault(path, "P5:default")
    _record_default(params)

    # P3 参考文档
    ref_layer = _spec_to_layers(reference_profile, "P3:reference") \
        if cfg.get("mode") == "reference" else {}
    if ref_layer:
        _merge_layer(params, ref_layer, "P3:reference", sources)

    # P2 模板
    tpl_layer = _spec_to_layers(template_spec, "P2:template") \
        if cfg.get("mode") == "template" else {}
    if tpl_layer:
        _merge_layer(params, tpl_layer, "P2:template", sources)

    # P1 用户显式：intent 映射参数（语义）先、overrides（精确）后
    if intent_params:
        _merge_layer(params, intent_params, "P1:intent", sources)
    if cfg.get("overrides"):
        _merge_layer(params, cfg["overrides"], "P1:overrides", sources)

    return params, sources
