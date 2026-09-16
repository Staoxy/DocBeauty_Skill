# -*- coding: utf-8 -*-
"""Role Mapping（DESIGN_V15 §9）：模板角色 ↔ 目标角色。

V1.5 中两侧角色都已归一化到规范词表（body/h1/h2/h3/caption/title/toc_title），
本模块负责：确认模板提供了哪些角色、来源是什么、缺失哪些（缺失 → 回退目标
现状/默认，绝不猜测）。
"""

CANONICAL_ROLES = ("body", "h1", "h2", "h3", "caption", "title", "toc_title")

# 内置同义词表（精确名层用；样式名小写比较）
ROLE_SYNONYMS = {
    "body": ["normal", "正文", "body text", "正文文本"],
    "h1": ["heading 1", "标题 1", "一级标题", "章标题", "chapter", "chapter title"],
    "h2": ["heading 2", "标题 2", "二级标题", "节标题"],
    "h3": ["heading 3", "标题 3", "三级标题", "小节标题"],
    "caption": ["caption", "题注", "图题", "表题", "图注"],
    "title": ["title", "标题", "文档标题", "题目"],
    "toc_title": ["toc heading", "toc 标题", "目录标题"],
}

REQUIRED_ROLES = ("body", "h1")


def build_role_map(template_spec, target_analysis=None) -> tuple:
    """返回 (role_map, warnings)。

    role_map: {role: 来源描述}——模板提供该角色的参数；None 表示模板未提供。
    """
    warnings = []
    roles = (template_spec or {}).get("roles", {})
    meta = (template_spec or {}).get("roles_meta", {})
    role_map = {}
    for role in CANONICAL_ROLES:
        if role in roles and roles[role]:
            role_map[role] = meta.get(role, "template")
        else:
            role_map[role] = None
            if role in REQUIRED_ROLES:
                warnings.append(
                    f"role_mapper: 模板未提供角色 '{role}' 的格式，将回退目标现状/默认模板")
    return role_map, warnings
