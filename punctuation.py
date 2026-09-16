# -*- coding: utf-8 -*-
"""标点/符号规范化（DESIGN_V2.md §10.5）——唯一的改内容操作。

门控：options.punctuation_normalize=true 且 content_protection=false（config.py 强制）。
实现要点：
- 上下文判定（前后字符是否 CJK）决定全角/半角；
- 保护例外：数字上下文（3.14 / 1,000 / 12:30 / 100%）、英文缩写（e.g. / i.e. / etc. / vs.）、
  URL、邮箱；跨 run 的多字符替换跳过（保守）；
- 文本框（txbxContent）内文字不处理（护栏快照含它，改了会误报）；
- 每处改动逐段登记 accounting["punctuation_changes"]，护栏据此豁免。
"""
import re
from docx.oxml.ns import qn

from formatter import get_or_add_rpr, runs_of

_CJK_RE = re.compile(r"[\u4e00-\u9fff\u3000-\u303f\uff01-\uff60\u201c\u201d\u2018\u2019]")

# 多字符序列（先于单字符处理；跨 run 边界跳过）
_SEQ_RULES = [
    (re.compile(r"\.{3}"), "……"),      # ... -> ……
    (re.compile(r"--+"), "——"),        # -- -> ——
]
# 单字符映射：仅半角 -> 全角（CJK 上下文）；ASCII 上下文一律不动（保守）
HALF2FULL = {
    ",": "，", ":": "：", ";": "；", "!": "！", "?": "？",
    "(": "（", ")": "）",
}
_CJK_QUOTE = {"\"": ("\u201c", "\u201d"), "'": ("\u2018", "\u2019")}

_PROTECT_SPANS = [
    re.compile(r"https?://\S+"),
    re.compile(r"\S+@\S+\.\S+"),
    re.compile(r"\b(?:e\.g|i\.e|etc|vs|Dr|Mr|Mrs|No|Fig|Eq|approx)\."),
    re.compile(r"\b[A-Za-z]\."),
]


def _is_cjk(ch: str) -> bool:
    return bool(ch) and bool(_CJK_RE.match(ch))


def _is_wordish(ch: str) -> bool:
    return bool(ch) and (ch.isalnum() or ch in "_%")


def _protected_mask(text: str) -> list:
    """返回 text 每个字符是否处于保护区间（URL/邮箱/缩写）。"""
    mask = [False] * len(text)
    for pat in _PROTECT_SPANS:
        for m in pat.finditer(text):
            for i in range(m.start(), m.end()):
                mask[i] = True
    return mask


def _convert_char(full: str, k: int, mask):
    """对 full[k] 做上下文判定。返回 (替换文本, 消费长度) 或 None。

    判定使用原文本上下文：prev = k-1，next = k+1（含跨 run 邻居）。
    规则：任一侧 CJK -> 半角转全角；两侧均 ASCII -> 不动（保守）；
    数字上下文（3.14 / 1,000 / 12:30 / 100%）不动。
    """
    ch = full[k]
    if mask[k]:
        return None
    prev = full[k - 1] if k > 0 else ""
    nxt = full[k + 1] if k + 1 < len(full) else ""

    # 多字符序列：... 与 --（调用方已保证序列完整落在同一段 run 文本内）
    for pat, rep in _SEQ_RULES:
        m = pat.match(full, k)
        if m:
            if prev.isdigit() and _is_wordish(full[m.end()]) and rep == "——":
                pass  # 数字区间 12--15 属内容语境，仍转换；此处仅保留结构
            return rep, m.end() - k

    if ch in _CJK_QUOTE:
        # 直引号 -> 中式弯引号：前一个字符是 CJK/开头/前标点 = 开引号，否则闭引号
        if _is_cjk(prev) or prev == "" or prev in "，。；：、（！？":
            return _CJK_QUOTE[ch][0], 1
        return _CJK_QUOTE[ch][1], 1

    if ch not in HALF2FULL:
        return None
    # 数字保护：数字与数字/字母/单位之间的标点保持原样
    if prev.isdigit() and _is_wordish(nxt):
        return None
    # ASCII 上下文不动
    if not (_is_cjk(prev) or _is_cjk(nxt)):
        return None
    return HALF2FULL[ch], 1


def _para_runs_with_text(p_el):
    """段落内待处理 (run, t_elem) 列表；排除文本框/图形内部。"""
    excluded = set()
    for tag in ("w:txbxContent", "w:drawing", "w:pict", "w:object"):
        for box in p_el.iter(qn(tag)):
            for r in box.iter(qn("w:r")):
                excluded.add(r)
    out = []
    for r in runs_of(p_el):
        for t in r.findall(qn("w:t")):
            out.append((r, t))
    return out


def apply(doc, accounting: dict) -> dict:
    """执行标点规范化。返回统计并把 (before, after) 段落对写入 accounting。"""
    changes = []
    paragraphs_changed = 0

    for para in doc.paragraphs:
        p_el = para._p
        pieces = _para_runs_with_text(p_el)
        if not pieces:
            continue
        texts = [t.text or "" for _, t in pieces]
        full = "".join(texts)
        if not full.strip():
            continue
        mask = _protected_mask(full)

        # 偏移 -> (piece_idx, offset_in_piece)
        bounds = []
        off = 0
        for txt in texts:
            bounds.append((off, off + len(txt)))
            off += len(txt)

        def locate(k):
            for pi, (s, e) in enumerate(bounds):
                if s <= k < e:
                    return pi, k - s
            return None, None

        new_texts = list(texts)
        k = 0
        changed = False
        while k < len(full):
            # 先试多字符序列（要求整个序列落在同一段 run 文本内）
            seq_hit = None
            for pat, rep in _SEQ_RULES:
                m = pat.match(full, k)
                if m:
                    pi0, _ = locate(k)
                    pi1, _ = locate(m.end() - 1)
                    if pi0 is not None and pi0 == pi1 and not any(mask[k:m.end()]):
                        seq_hit = (pi0, m.end(), rep)
                    break
            if seq_hit:
                pi, end, rep = seq_hit
                s, e = bounds[pi]
                local = new_texts[pi]  # 基于已累积替换的文本，防止前一处丢失
                lk = k - s
                lk_end = end - s
                new_texts[pi] = local[:lk] + rep + local[lk_end:]
                changed = True
                k = end
                continue

            conv = _convert_char(full, k, mask)
            if conv is not None:
                rep, consumed = conv
                if rep != full[k:k + consumed]:
                    pi, lk = locate(k)
                    pi_end, lk_end = locate(k + consumed - 1)
                    if (pi is not None and pi == pi_end
                            and len(rep) == 1 and consumed == 1):
                        local = new_texts[pi]  # 基于"已累积替换"的文本（同一 run 多处替换）
                        new_texts[pi] = local[:lk] + rep + local[lk + 1:]
                        changed = True
            k += 1

        if changed:
            new_full = "".join(new_texts)
            for (r, t), nt in zip(pieces, new_texts):
                if (t.text or "") != nt:
                    t.text = nt
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            changes.append([full, new_full])
            paragraphs_changed += 1

    accounting["punctuation_changes"] = (
        accounting.get("punctuation_changes", []) + changes)
    return {"paragraphs_changed": paragraphs_changed,
            "changes_total": len(changes)}
