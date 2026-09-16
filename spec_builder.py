# -*- coding: utf-8 -*-
"""Spec Builder（DESIGN_V15 §11）：Resolver 输出 → Formatting Spec。

Formatting Spec 是现有 effective 参数（= V1 overrides 词表）的超集：
params 同名同形，另附 role_map / sources / policy / unsupported / warnings。
Engine 只需接受 params 作为 effective，执行路径不变。
"""
from role_mapper import build_role_map
import rule_resolver


def _sections_policy(template_spec, target_analysis) -> tuple:
    """§12.2 节结构兼容性：模板多节页码 vs 目标单节 → 不套用 + 警告。"""
    warnings = []
    policy = "preserve_target"
    if not template_spec:
        return policy, warnings
    tpl_numbered = [s for s in template_spec.get("sections", [])
                    if s.get("page_number")]
    target_sections = (target_analysis or {}).get("section_count", 1)
    if len(tpl_numbered) > 1 and target_sections < len(tpl_numbered):
        policy = "incompatible"
        warnings.append(
            f"spec_builder: 模板为 {len(tpl_numbered)} 节页码结构，目标文档仅 "
            f"{target_sections} 节，三段式页码不套用（不自动造节，§12.2）")
    return policy, warnings


def build_formatting_spec(cfg: dict, builtin_params: dict, template_spec=None,
                          reference_profile=None, target_analysis=None,
                          intent_params=None) -> dict:
    params, sources = rule_resolver.resolve(
        cfg, builtin_params, template_spec=template_spec,
        reference_profile=reference_profile, intent_params=intent_params)

    role_map, map_warnings = build_role_map(template_spec or reference_profile,
                                            target_analysis)

    policy = {
        "numbering": "preserve_target",          # §12.1 编号定义不迁移
        "existing_page_number": "skip" if (target_analysis or {}).get(
            "existing_page_number_fields") else "add",
        "existing_toc": "update_only" if (target_analysis or {}).get(
            "existing_toc_field") else "add",
        "header_footer": "preserve_if_present",  # §12.3
    }
    sec_policy, sec_warnings = _sections_policy(template_spec, target_analysis)
    policy["sections"] = sec_policy

    warnings = list(map_warnings) + list(sec_warnings)
    if template_spec and template_spec.get("unsupported"):
        warnings.append(
            f"spec_builder: 模板含 Engine 不支持的特性，将跳过并保持目标现状: "
            f"{template_spec['unsupported']}")

    return {
        "mode": cfg.get("mode", "auto"),
        "params": params,
        "role_map": role_map,
        "sources": sources,
        "policy": policy,
        "unsupported": list((template_spec or {}).get("unsupported", [])),
        "warnings": warnings,
    }
