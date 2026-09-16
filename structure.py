# -*- coding: utf-8 -*-
"""Structure Detection + ExclusionZones（DESIGN_V2.md §7/§8/§15，D2 一级公民）。"""
import re
from dataclasses import dataclass, field

from docx.oxml.ns import qn

import constants as C
from analyzer import para_text
from utils import ooxml as ox

_W_T = qn("w:t")

CAPTION_RES = [re.compile(p) for p in C.CAPTION_PATTERNS]
REF_RES = [re.compile(h + r"\s*$") for h in C.REFERENCE_HEADINGS]
APPENDIX_RES = [re.compile("^" + h) for h in C.APPENDIX_HEADINGS]
FRONT_RES = [re.compile(h + r"\s*$") for h in C.FRONT_MATTER_HEADINGS]
DOT_LEADER_RE = re.compile(r"[.\u2026]{3,}\s*\d+\s*$")


def is_blank_paragraph(p_el) -> bool:
    """空白段定义（§12）：无文字、无域、无图片、无分页符、无书签。"""
    if "".join(t.text or "" for t in p_el.iter(_W_T)).strip():
        return False
    for tag in ("w:drawing", "w:pict", "w:object", "w:fldChar", "w:fldSimple",
                "w:instrText", "w:bookmarkStart", "w:commentRangeStart"):
        if p_el.findall(".//" + qn(tag)):
            return False
    for br in p_el.findall(".//" + qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return False
    pPr = p_el.find(qn("w:pPr"))
    if pPr is not None:
        if pPr.find(qn("w:sectPr")) is not None:
            return False
        pbb = pPr.find(qn("w:pageBreakBefore"))
        if pbb is not None and pbb.get(qn("w:val")) != "false":
            return False
    return True


def para_is_caption(text: str) -> bool:
    t = text.strip()
    return bool(t) and any(r.match(t) for r in CAPTION_RES)


def para_bold(p_el) -> bool:
    runs = p_el.findall(qn("w:r"))
    if not runs:
        return False
    any_text = False
    for r in runs:
        txt = "".join(t.text or "" for t in r.iter(_W_T)).strip()
        if not txt:
            continue
        any_text = True
        rPr = ox.get_child(r, "w:rPr")
        b = ox.get_child(rPr, "w:b") if rPr is not None else None
        if b is None or b.get(qn("w:val")) in ("false", "0", "none"):
            return False
    return any_text


def para_max_size(p_el):
    sizes = []
    for r in p_el.findall(qn("w:r")):
        rPr = ox.get_child(r, "w:rPr")
        if rPr is None:
            continue
        sz = ox.get_child(rPr, "w:sz")
        try:
            if sz is not None:
                sizes.append(int(sz.get(qn("w:val"))) / 2)
        except (TypeError, ValueError):
            pass
    return max(sizes) if sizes else None


def para_has_drawing_or_field(p_el) -> bool:
    for tag in ("w:drawing", "w:pict", "w:fldChar", "w:fldSimple"):
        if p_el.findall(".//" + qn(tag)):
            return True
    return False


# ---------------------------------------------------------------------------
# 编号模板自动检测（§8.2）
# ---------------------------------------------------------------------------
def _match_numeric_level(text: str):
    """numeric 模板按小数点数定级。"""
    t = text.strip()
    m = re.match(r"^(\d+(?:\.\d+)+)[\.、]?\s*", t)
    if m:
        dots = m.group(1).count(".")
        lvl = min(dots + 1, 4)  # 1 -> l1, 1.1 -> l2 ...
        return lvl
    m = re.match(r"^(\d+)[\.、]\s*", t)
    if m:
        return 1
    return None


def _match_template_level(tpl_name: str, text: str):
    t = text.strip()
    if tpl_name == "numeric":
        return _match_numeric_level(t)
    for lvl in ("l4", "l3", "l2", "l1"):
        if re.match(C.HEADING_PATTERNS[tpl_name][lvl], t):
            return int(lvl[1])
    return None


def detect_numbering_template(paragraphs):
    """返回 (best_tpl_name|None, {idx: level})。

    评分 = 命中数 - 2*层级跳跃数；>=3 才采纳（§8.2）。
    """
    best_name, best_map, best_score = None, {}, 0
    for tpl_name in C.HEADING_PATTERNS:
        hits = {}
        for i, p in enumerate(paragraphs):
            text = para_text(p).strip()
            if not text or len(text) > 60:
                continue
            lvl = _match_template_level(tpl_name, text)
            if lvl:
                hits[i] = lvl
        if len(hits) < 2:
            continue
        seq = list(hits.values())
        # 跳级 = 向下超过一级（如 1->3）；回退到高级（3->1 新章节）是正常层级
        jumps = sum(1 for a, b in zip(seq, seq[1:]) if b - a > 1)
        # 序列从 1 级开始 +1 偏好：同一文本常同时命中多个模板
        # （如"一、"= cn_zhang.l3 = cn_gov.l1），自然层级应从顶级开始
        start_bonus = 1 if seq[0] == 1 else 0
        score = len(hits) - 2 * jumps + start_bonus
        if score > best_score:
            best_name, best_map, best_score = tpl_name, hits, score
    if best_name is None:
        return None, {}
    # V2 只应用 Heading 1-3（§11.3），第 4 级候选截断
    clipped = {k: min(v, 3) for k, v in best_map.items()}
    seq = list(clipped.values())
    no_downward_jump = not any(b - a > 1 for a, b in zip(seq, seq[1:]))
    # 采纳条件 A：评分达标（>=3 命中减跳级惩罚）
    if best_score >= 3:
        return best_name, clipped
    # 采纳条件 B（放宽）：两级序列（一、/（一））是最常见的中文论文结构——
    # 序列从 1 级开始且无向下跳级时采纳；防误报靠 §8.3 的一票否决兜底
    if len(clipped) >= 2 and seq[0] == 1 and no_downward_jump:
        return best_name, clipped
    return None, {}


# ---------------------------------------------------------------------------
# 标题检测（§8.3 打分）
# ---------------------------------------------------------------------------
TOC_TITLE_NAMES = {"目录", "目 录", "contents", "CONTENTS"}


def score_heading_candidate(p_el, text, numbering_map_idx, body_mode_pt):
    """返回 (score, veto)。veto=True 表示绝不可能是标题。"""
    t = text.strip()
    if not t:
        return 0, True
    if para_is_caption(t):
        return 0, True
    if t.lower() in TOC_TITLE_NAMES:
        return 0, True  # 目录标题进 §16 专用格式，不提升为 Heading（避免目录收录自身）
    if para_has_drawing_or_field(p_el):
        return 0, True
    if t.endswith(C.SENTENCE_END_PUNCT):
        return -4, True
    if t.endswith(("：", ":")) and len(t) <= 12:
        return 0, True
    score = 0
    if numbering_map_idx is not None:
        score += 3
    if para_bold(p_el):
        score += 2
    mx = para_max_size(p_el)
    if mx is not None and body_mode_pt and mx >= body_mode_pt + 1:
        score += 2
    if len(t) <= 40:
        score += 1
    if not t.endswith(C.SENTENCE_END_PUNCT):
        score += 1
    return score, False


# ---------------------------------------------------------------------------
# ExclusionZones（§7.2）
# ---------------------------------------------------------------------------
@dataclass
class ExclusionZones:
    skip_indent: set = field(default_factory=set)     # 段落 idx：跳过缩进/对齐
    skip_all: set = field(default_factory=set)        # 段落 idx：跳过除字体外全部归一
    cover_ids: set = field(default_factory=set)       # 段落 idx：封面（整区排除，含字体，§7.1）
    captions: set = field(default_factory=set)        # 段落 idx：图注/表注
    headings: dict = field(default_factory=dict)      # 段落 idx -> level（已采信）
    candidates: list = field(default_factory=list)    # 低置信度候选（只报告）
    zones: dict = field(default_factory=dict)         # 区域 -> [start, end]
    numbering_hits: dict = field(default_factory=dict)  # idx -> level（编号模板命中）
    numbering_template: str = None
    notes: list = field(default_factory=list)

    def to_report(self):
        return {
            "zones": self.zones,
            "numbering_template": self.numbering_template,
            "candidates": self.candidates[:50],
            "notes": self.notes,
        }

    def shift(self, from_idx: int, delta: int):
        """在 from_idx 处插入 delta 个段落后，平移所有索引集合（供 TOC 插入后调用）。"""
        def sh(s):
            return {i + delta if i >= from_idx else i for i in s}
        self.skip_indent = sh(self.skip_indent)
        self.skip_all = sh(self.skip_all)
        self.captions = sh(self.captions)
        self.headings = {i + delta if i >= from_idx else i: lv
                         for i, lv in self.headings.items()}
        self.numbering_hits = {i + delta if i >= from_idx else i: lv
                               for i, lv in self.numbering_hits.items()}
        self.zones = {
            name: [s + delta if s >= from_idx else s, e + delta if e >= from_idx else e]
            for name, (s, e) in self.zones.items()
        }


def _find_zone_end(paragraphs, texts, start, accepted_headings):
    """区域结束 = 下一个已采信标题前。"""
    for j in range(start + 1, len(paragraphs)):
        if j in accepted_headings:
            return j - 1
    return len(paragraphs) - 1


def detect(doc, analysis: dict) -> ExclusionZones:
    from analyzer import heading_level_of
    styles_elem = doc.styles.element
    paragraphs = [p._p for p in doc.paragraphs]
    texts = [para_text(p) for p in paragraphs]
    n = len(paragraphs)
    zones = ExclusionZones()
    body_mode = analysis.get("body_font_size_mode_pt", 12.0)

    # 1) 图注/表注（先于标题，一票否决）
    for i, t in enumerate(texts):
        if para_is_caption(t):
            zones.captions.add(i)
            zones.skip_indent.add(i)

    # 2) 编号模板
    tpl_name, num_hits = detect_numbering_template(paragraphs)
    zones.numbering_template = tpl_name
    zones.numbering_hits = num_hits

    # 3) 标题采信：样式级（§8.1 优先级 1/2）
    style_based = set()
    for i, p in enumerate(paragraphs):
        lvl = heading_level_of(styles_elem, p)
        if lvl in (1, 2, 3):
            zones.headings[i] = lvl
            style_based.add(i)
        elif lvl == 9:
            zones.notes.append(f"段落 {i} 使用了 Heading 4+ 或其他大纲级别，未采信")

    # 4) 启发式打分（仅未采信段落）
    for i, p in enumerate(paragraphs):
        if i in zones.headings or i in zones.captions:
            continue
        if not texts[i].strip():
            continue
        lvl = num_hits.get(i)
        score, veto = score_heading_candidate(p, texts[i], lvl, body_mode)
        if veto or score < 3:
            continue
        item = {"paragraph": i, "text": texts[i].strip()[:40],
                "score": score, "suggested_level": lvl or 2}
        if score >= 5:
            zones.headings[i] = lvl or 2
        else:
            zones.candidates.append(item)

    # 5) 参考文献区（§7.1）
    for i, t in enumerate(texts):
        ts = t.strip()
        if any(r.match(ts) for r in REF_RES) and len(ts) <= 10:
            end = _find_zone_end(paragraphs, texts, i, zones.headings)
            zones.zones["references"] = [i, end]
            for j in range(i + 1, end + 1):
                zones.skip_indent.add(j)
            break

    # 6) 附录区
    for i, t in enumerate(texts):
        ts = t.strip()
        if any(r.match(ts) for r in APPENDIX_RES) and len(ts) <= 12:
            end = _find_zone_end(paragraphs, texts, i, zones.headings)
            zones.zones["appendix"] = [i, end]
            break

    # 7) 前置区（摘要等，决定目录插入位置；正文段仍参与归一）
    for i, t in enumerate(texts):
        ts = t.strip()
        if any(r.match(ts) for r in FRONT_RES) and len(ts) <= 10:
            end = _find_zone_end(paragraphs, texts, i, zones.headings)
            zones.zones["front_matter"] = [i, end]
            break

    # 8) 已有目录区：TOC 域或手工目录（§16）
    toc_start = None
    for i, p in enumerate(paragraphs):
        for it in p.iter(qn("w:instrText")):
            if it.text and re.search(r"\bTOC\b", it.text):
                toc_start = i
                break
        if toc_start is not None:
            break
    manual = False
    if toc_start is None:
        for i, t in enumerate(texts):
            if t.strip() in ("目录", "目 录", "Contents", "CONTENTS"):
                # 后续 >=3 个点线页码段落 → 手工目录
                cnt = 0
                for j in range(i + 1, min(i + 30, n)):
                    if DOT_LEADER_RE.search(texts[j]):
                        cnt += 1
                    elif texts[j].strip():
                        break
                if cnt >= 3:
                    toc_start, manual = i, True
                    zones.zones["toc_existing_manual"] = [i, i + cnt]
                    zones.notes.append(
                        f"检测到手工文字目录（段落 {i}~{i+cnt}），默认保留不删除")
                    break
    if toc_start is not None and manual:
        for j in range(zones.zones["toc_existing_manual"][0],
                       zones.zones["toc_existing_manual"][1] + 1):
            zones.skip_all.add(j)

    # 9) 封面区（§7.1：整区排除）
    cover_end = None
    if analysis.get("cover_like_first_section"):
        for i, p in enumerate(paragraphs):
            pPr = p.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                cover_end = i
                break
        if cover_end is None:
            # 单节文档：只把开头的大字号居中段当封面
            for i in range(min(3, n)):
                p = paragraphs[i]
                jc = None
                pPr = p.find(qn("w:pPr"))
                if pPr is not None:
                    jce = ox.get_child(pPr, "w:jc")
                    jc = jce.get(qn("w:val")) if jce is not None else None
                mx = para_max_size(p)
                if jc == "center" and mx and mx >= 18:
                    cover_end = i
            if cover_end is not None and len(texts) > 1 and not texts[0].strip():
                pass
        if cover_end is not None:
            zones.zones["cover"] = [0, cover_end]
            for j in range(0, cover_end + 1):
                zones.skip_all.add(j)
                zones.cover_ids.add(j)

    # 10) 封面/手工目录区内的启发式标题撤销（§7.1：封面不参与标题提升，
    #     防止大字号封面标题被误提升为 Heading 进目录）
    for i in list(zones.headings):
        if i in zones.skip_all and i not in style_based:
            del zones.headings[i]
    zones.candidates = [c for c in zones.candidates
                        if c["paragraph"] not in zones.skip_all]

    return zones
