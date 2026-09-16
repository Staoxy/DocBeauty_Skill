# -*- coding: utf-8 -*-
"""Content Guard（DESIGN_V2.md §17）：快照 / 指纹 / diff / 完整性计数。

三段式（D3）：修改前 snapshot() -> 修改后 snapshot() -> verify()。
护栏失败 -> 调用方必须删除临时输出、不产出文件（§17.5）。
"""
import difflib
import hashlib
import re
import unicodedata

from docx.oxml.ns import qn

_W_T = qn("w:t")
_WS_RE = re.compile(r"[\s\u3000]+")


def normalize(text: str) -> str:
    """NFC + 连续空白（含全角空格）折叠 + strip（§17.1）。"""
    if not text:
        return ""
    return _WS_RE.sub(" ", unicodedata.normalize("NFC", text)).strip()


def _sha(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _cell_texts(tbl):
    return [normalize("".join(t.text or "" for t in tc.iter(_W_T)))
            for tc in tbl.iter(qn("w:tc"))]


def _counts(doc):
    body = doc.element.body
    return {
        "hyperlinks": len(body.findall(".//" + qn("w:hyperlink"))),
        "bookmarks": len(body.findall(".//" + qn("w:bookmarkStart"))),
        "fields": (len(body.findall(".//" + qn("w:fldChar"))) +
                   len(body.findall(".//" + qn("w:fldSimple")))),
        "tables": len(body.findall(".//" + qn("w:tbl"))),
        "inline_images": len(body.findall(".//" + qn("wp:inline"))),
        "floating_images": len(body.findall(".//" + qn("wp:anchor"))),
        "sections": len(doc.sections),
    }


def _expected_exact():
    """必须精确相等的计数（§17.4 + Q03/Q04/Q05/Q16）。"""
    return {"tables", "inline_images", "floating_images", "sections"}


def snapshot(doc) -> dict:
    """修改前/后各调用一次。页眉页脚/脚注/文本框/批注不进快照（§17.1）。"""
    from analyzer import para_text
    body = doc.element.body
    para_texts = [normalize(para_text(p._p)) for p in doc.paragraphs]
    tables = [_cell_texts(tbl) for tbl in body.findall(".//" + qn("w:tbl"))]
    return {
        "para_texts": para_texts,
        "tables": tables,
        "para_fingerprint": _sha("\n".join(para_texts)),
        "table_fingerprints": [_sha("\x01".join(cells)) for cells in tables],
        "counts": _counts(doc),
    }


def _apply_allowed_changes(after_texts, punctuation_changes):
    """把标点等预期变更替换回 before 文本，使序列可比较（§17.3）。"""
    texts = list(after_texts)
    for before_t, after_t in punctuation_changes or []:
        b, a = normalize(before_t), normalize(after_t)
        if not a:
            continue
        for i, t in enumerate(texts):
            if t == a:
                texts[i] = b
                break
        else:
            # 段内局部替换（整段不等但包含改后文本）
            for i, t in enumerate(texts):
                if a in t:
                    texts[i] = t.replace(a, b)
                    break
    return texts


def verify(before: dict, after: dict, accounting: dict = None) -> dict:
    """护栏校验（§17.2–17.4）。

    accounting:
      excluded_generated_texts: 允许新增的文本（目录标题/占位/提示行）
      punctuation_changes: [(before, after)] 预期内容变更对
    """
    accounting = accounting or {}
    allowed_new = {normalize(t) for t in accounting.get("excluded_generated_texts", [])}

    # ---- 完整性计数（只增不减 / 精确相等） ----
    integrity, integrity_failures = {}, []
    for name, bv in before["counts"].items():
        av = after["counts"].get(name, 0)
        if name in _expected_exact():
            ok = av == bv
        else:
            ok = av >= bv
        integrity[name] = {"before": bv, "after": av, "ok": ok}
        if not ok:
            integrity_failures.append(f"{name}: {bv} -> {av}")

    # ---- 段落 diff（空白段折叠；插入仅允许白名单文本） ----
    before_body = [t for t in before["para_texts"] if t]
    after_body_raw = [t for t in after["para_texts"] if t]
    after_body = _apply_allowed_changes(after_body_raw,
                                        accounting.get("punctuation_changes"))

    diffs = []
    matcher = difflib.SequenceMatcher(a=before_body, b=after_body, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag == "insert":
            bad = [t for t in after_body[j1:j2] if t not in allowed_new]
            if not bad:
                continue
            diffs.append({"kind": "insert", "after": bad[:5]})
        elif tag == "delete":
            diffs.append({"kind": "delete", "before": before_body[i1:i2][:5]})
        else:  # replace
            diffs.append({"kind": "replace",
                          "before": before_body[i1:i2][:5],
                          "after": after_body[j1:j2][:5]})

    # ---- 表格 diff ----
    if len(before["tables"]) == len(after["tables"]):
        for ti, (b_cells, a_cells) in enumerate(zip(before["tables"], after["tables"])):
            if b_cells != a_cells:
                for ci, (b, a) in enumerate(zip(b_cells, a_cells)):
                    if b != a:
                        diffs.append({"kind": "table_cell", "table": ti,
                                      "cell": ci, "before": b[:40], "after": a[:40]})
                        if len([d for d in diffs if d["kind"] == "table_cell"]) >= 5:
                            break
    else:
        integrity_failures.append(
            f"tables: {len(before['tables'])} -> {len(after['tables'])}")

    content_changed = bool(diffs) or bool(integrity_failures)
    return {
        "content_changed": content_changed,
        "diffs": diffs[:20],
        "integrity": integrity,
        "integrity_failures": integrity_failures,
        "ok": not content_changed,
    }
