# -*- coding: utf-8 -*-
"""空白段落清理（DESIGN_V2.md §12）：硬排除清单 + 连续空白收敛为 1 个。"""
from docx.oxml.ns import qn

from structure import is_blank_paragraph


def _adjacent_table(el, direction: str):
    """沿方向跳过其他空段后，第一个非空元素是否为表格。"""
    cur = el.getprevious() if direction == "prev" else el.getnext()
    while cur is not None and cur.tag == qn("w:p"):
        if not is_blank_paragraph(cur):
            break
        cur = cur.getprevious() if direction == "prev" else cur.getnext()
    return cur is not None and cur.tag == qn("w:tbl")


def _between_tables(el) -> bool:
    """两个表格之间：删除空段会导致 Word 把两表合并（硬排除 3）。"""
    return _adjacent_table(el, "prev") and _adjacent_table(el, "next")


def _is_protected(doc, el, skip_all_ids: set) -> bool:
    """硬排除：封面/目录区（skip_all）、两表之间、文档首段。
    其余硬排除项（sectPr/分页符/域/书签/图形）已由 is_blank_paragraph 覆盖。
    """
    if id(el) in skip_all_ids:
        return True
    if _between_tables(el):
        return True
    paras = doc.paragraphs
    return bool(paras) and el is paras[0]._p


def run(doc, zones, options: dict, accounting: dict) -> dict:
    """连续 >=2 个空白段收敛为 1 个（保守策略）。每次删除登记进 accounting。"""
    if not options.get("delete_blank_paragraphs", True):
        return {"deleted": 0, "disabled": True}

    body = doc.element.body
    skip_all_ids = {doc.paragraphs[i]._p for i in zones.skip_all
                    if i < len(doc.paragraphs)}
    children = list(body)

    # 收集连续空白段 runs
    runs, current = [], []
    for el in children:
        if el.tag == qn("w:p") and is_blank_paragraph(el):
            current.append(el)
        else:
            if current:
                runs.append(current)
            current = []
    if current:
        runs.append(current)

    deleted = 0
    for blank_run in runs:
        if len(blank_run) < 2:
            continue
        deletable = [el for el in blank_run if not _is_protected(doc, el, skip_all_ids)]
        # 收敛为 1 个：保留 deletable[0]，删除其余
        for el in deletable[1:]:
            body.remove(el)
            deleted += 1

    accounting["blank_deleted"] = accounting.get("blank_deleted", 0) + deleted
    return {"deleted": deleted}
