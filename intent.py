# -*- coding: utf-8 -*-
"""Intent 校验/归一化（DESIGN_V15 §4）。

Agent 把自然语言翻译成受控 Intent JSON；本模块只做校验与归一化
（不是 NL 解析——NL 理解是 Agent 的职责，§2.1）。未知键降级为 warning，
枚举外值回退默认。preferences → 参数的映射表是确定性代码，映射结果
进报告 config.intent_resolved 回显。
"""
import copy

import jsonschema

INTENT_SCHEMA = {
    "$schema": "http://json-schema.org/draft-07/schema#",
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "style": {"type": "string", "enum": ["academic", "formal", "clean"]},
        "document_type": {"type": "string",
                          "enum": ["auto", "academic", "experiment", "report",
                                   "meeting", "generic"]},
        "content_protection": {"type": "boolean"},
        "preferences": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "heading_emphasis": {"type": "string",
                                     "enum": ["default", "strong", "subtle"]},
                "table_density": {"type": "string",
                                  "enum": ["compact", "default", "comfortable"]},
                "body_consistency": {"type": "boolean"},
                "caption_align": {"type": "string", "enum": ["center", "left"]},
                "line_density": {"type": "string",
                                 "enum": ["compact", "default", "loose"]},
            },
        },
        "notes": {"type": "string"},
    },
}

# preferences → 参数映射（确定性；只映射 Engine 真实支持的旋钮）
# heading_emphasis: 标题段前段后间距倍率
_HEADING_EMPHASIS = {
    "strong": 1.5,
    "default": 1.0,
    "subtle": 0.6,
}
# line_density: 行距
_LINE_DENSITY = {
    "compact": 1.15,
    "default": None,      # 不干预（保持模板/默认值）
    "loose": 1.75,
}
# table_density: 单元格边距 dxa（tables.format_tables 应用 tblCellMar）
_TABLE_DENSITY = {
    "compact": 60,
    "default": None,
    "comfortable": 160,
}


def _scale_spacing(headings: dict, factor: float) -> None:
    for h in headings.values():
        if h.get("before_pt") is not None:
            h["before_pt"] = round(h["before_pt"] * factor, 1)
        if h.get("after_pt") is not None:
            h["after_pt"] = round(h["after_pt"] * factor, 1)


def validate_intent(intent) -> tuple:
    """返回 (normalized, warnings)。非法输入不抛异常，降级 + warning。"""
    warnings = []
    normalized = {"style": None, "document_type": None,
                  "content_protection": None, "preferences": {}}
    if intent is None:
        return normalized, warnings
    if not isinstance(intent, dict):
        return normalized, [f"intent: 期望对象，得到 {type(intent).__name__}，已忽略"]

    # 未知键
    for k in intent:
        if k not in INTENT_SCHEMA["properties"]:
            warnings.append(f"intent: 未知键 '{k}' 已忽略（受控词表见 schemas/intent_schema.json）")

    for key in ("style", "document_type", "content_protection"):
        if key in intent:
            allowed = INTENT_SCHEMA["properties"][key].get("enum")
            v = intent[key]
            if allowed is not None and v not in allowed:
                warnings.append(f"intent.{key}='{v}' 不在受控枚举内，已忽略")
            else:
                normalized[key] = v

    prefs = intent.get("preferences") or {}
    if not isinstance(prefs, dict):
        warnings.append("intent.preferences: 期望对象，已忽略")
        prefs = {}
    allowed_prefs = INTENT_SCHEMA["properties"]["preferences"]["properties"]
    for k, v in prefs.items():
        if k not in allowed_prefs:
            warnings.append(f"intent.preferences.{k}: 未知偏好已忽略")
            continue
        enum = allowed_prefs[k].get("enum")
        if enum is not None and v not in enum:
            warnings.append(f"intent.preferences.{k}='{v}' 不在受控枚举内，已忽略")
            continue
        normalized["preferences"][k] = v
    return normalized, warnings


def intent_to_params(normalized: dict, builtin_params: dict) -> tuple:
    """preferences → 参数层（P1:intent）。返回 (params_layer, resolved_echo)。"""
    prefs = normalized.get("preferences") or {}
    layer, echo = {}, {}

    factor = _HEADING_EMPHASIS.get(prefs.get("heading_emphasis"))
    if factor and factor != 1.0:
        headings = copy.deepcopy(builtin_params.get("headings") or {})
        _scale_spacing(headings, factor)
        layer["headings"] = headings
        echo["heading_emphasis"] = {
            "value": prefs["heading_emphasis"],
            "effect": "标题段前/段后间距 ×%.1f" % factor}

    ld = _LINE_DENSITY.get(prefs.get("line_density"))
    if ld is not None:
        layer["paragraph"] = {"line_spacing": ld}
        echo["line_density"] = {"value": prefs["line_density"],
                                "effect": f"行距 → {ld}"}

    td = _TABLE_DENSITY.get(prefs.get("table_density"))
    if td is not None:
        layer["tables"] = {"cell_margin_dxa": td}
        echo["table_density"] = {"value": prefs["table_density"],
                                 "effect": f"单元格边距 → {td} dxa"}

    ca = prefs.get("caption_align")
    if ca:
        layer["captions"] = {"align": ca}
        echo["caption_align"] = {"value": ca, "effect": f"题注对齐 → {ca}"}

    if prefs.get("body_consistency") is False:
        echo["body_consistency"] = {
            "value": False,
            "effect": "提示：正文一致性由 operations 控制，此偏好仅回显"}
    return layer, echo
