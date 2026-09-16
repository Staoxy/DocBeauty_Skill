# -*- coding: utf-8 -*-
"""Table Processor（DESIGN_V2.md §13）。原则：不修改单元格文字内容。"""
from docx.oxml.ns import qn

from analyzer import para_text
from layout import usable_width_emu
from utils import ooxml as ox

_EMU_PER_TWIP = 635  # 1 twip = 635 EMU


def _tbl_pr(tbl):
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = ox.make_elem("w:tblPr")
        tbl.insert(0, tblPr)
    return tblPr


def _grid_cols(tbl) -> list:
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        return []
    return [int(g.get(qn("w:w")) or 0) for g in grid.findall(qn("w:gridCol"))]


def _top_level_tables(doc):
    return [el for el in doc.element.body if el.tag == qn("w:tbl")]


def _norm_font_in_table(tbl, effective):
    from formatter import set_para_font
    west = effective["font"]["western"]
    east = effective["font"]["chinese"]
    size = effective["font"]["size_pt"]
    for p in tbl.iter(qn("w:p")):
        set_para_font(p, west, east, size)


def _header_cells(row):
    return row.findall(qn("w:tc"))


def _apply_cell_margins(tbl, dxa: int):
    tblPr = _tbl_pr(tbl)
    mar = ox.sub(tblPr, "w:tblCellMar")
    for tag in ("w:top", "w:left", "w:bottom", "w:right"):
        e = ox.get_child(mar, tag)
        if e is None:
            e = ox.make_elem(tag)
            mar.append(e)
        e.set(qn("w:w"), str(dxa))
        e.set(qn("w:type"), "dxa")


def format_tables(doc, effective: dict, options: dict, analysis: dict) -> dict:
    tpl = effective["tables"]
    borders_mode = options.get("table_borders", "auto")
    if borders_mode == "auto":
        borders_mode = tpl.get("borders", "grid")
    repeat_header = options.get("table_repeat_header", True) and tpl.get("repeat_header", True)

    usable = usable_width_emu(doc, effective) // _EMU_PER_TWIP  # dxa
    detail = {"tables": 0, "width_adjusted": 0, "warnings": [],
              "skipped_merged_width": []}

    for tbl in _top_level_tables(doc):
        detail["tables"] += 1
        tblPr = _tbl_pr(tbl)

        # ---- 宽度适配（可用宽度 = 节页宽 - 边距）----
        cols = _grid_cols(tbl)
        total = sum(cols)
        merged = ox.has_merged_cells(tbl)
        if cols and total > usable > 0:
            scale = usable / total
            new_cols = [max(int(c * scale), 200) for c in cols]
            grid = tbl.find(qn("w:tblGrid"))
            for g, w in zip(grid.findall(qn("w:gridCol")), new_cols):
                g.set(qn("w:w"), str(w))
            ox.set_tbl_width(tblPr, sum(new_cols))
            ox.set_tbl_layout_fixed(tblPr)
            if not merged:
                for tr in tbl.findall(qn("w:tr")):
                    for tc, w in zip(tr.findall(qn("w:tc")), new_cols):
                        ox.set_cell_width(tc, w)
            else:
                detail["skipped_merged_width"].append(detail["tables"] - 1)
                detail["warnings"].append(
                    f"Table {detail['tables'] - 1} 含合并单元格，仅做整表宽度适配")
            detail["width_adjusted"] += 1

        # ---- 边框 ----
        ox.set_tbl_borders(tblPr, borders_mode)

        # ---- 单元格边距（V1.5 intent.table_density 旋钮）----
        cell_margin = (effective.get("tables") or {}).get("cell_margin_dxa")
        if cell_margin:
            _apply_cell_margins(tbl, int(cell_margin))

        # ---- 行为控制：cantSplit + 表头重复 ----
        rows = tbl.findall(qn("w:tr"))
        for tr in rows:
            ox.set_row_cant_split(tr)
        if rows and repeat_header:
            ox.set_row_header_repeat(rows[0])

        # ---- 表头行：加粗 / 居中 / 底纹 / 三线表栏目线 ----
        if rows:
            header_row = rows[0]
            header_merged = bool(header_row.findall(".//" + qn("w:gridSpan")) or
                                 header_row.findall(".//" + qn("w:vMerge")))
            for tc in _header_cells(header_row):
                tcPr = ox.get_child(tc, "w:tcPr")
                if tcPr is None:
                    tcPr = ox.insert_ordered(tc, ox.make_elem("w:tcPr"))
                ox.set_cell_valign_center(tc)
                if tpl.get("header_shade"):
                    shd = ox.sub(tcPr, "w:shd", {"w:val": "clear",
                                                 "w:color": "auto",
                                                 "w:fill": tpl["header_shade"]})
                    shd.set(qn("w:val"), "clear")
                if borders_mode == "threeline" and not header_merged:
                    ox.set_cell_bottom_border(tc, 6)  # 0.75pt 栏目线
                for p in tc.iter(qn("w:p")):
                    if tpl.get("header_align") == "center":
                        pPr = p.find(qn("w:pPr"))
                        if pPr is None:
                            pPr = ox.insert_ordered(p, ox.make_elem("w:pPr"))
                        ox.set_alignment(pPr, "center")
                    if tpl.get("header_bold"):
                        for r in p.findall(qn("w:r")):
                            rPr = ox.get_child(r, "w:rPr")
                            if rPr is None:
                                rPr = ox.insert_ordered(r, ox.make_elem("w:rPr"))
                            ox.sub(rPr, "w:b")

        # ---- 单元格垂直居中（非表头行）----
        for tr in rows[1:]:
            for tc in tr.findall(qn("w:tc")):
                ox.set_cell_valign_center(tc)

        # ---- 字体统一（含嵌套表）----
        _norm_font_in_table(tbl, effective)

    return detail
