# -*- coding: utf-8 -*-
"""渲染级验证（DESIGN_V2.md Phase 7 / §18 增强）：用真实 Word 打开输出文档。

在**临时副本**上执行（交付文件保持与护栏验证时字节一致）：
- 更新全部域（目录实体化为真实条目、页码计算真实值）
- 统计真实页数
- 验证 TOC 域更新后有条目
- 可选导出 PDF（渲染所见即所得）
COM 不可用时返回 {"available": False}——是增强检查，不是门槛。
"""
import os
import shutil
import tempfile

from constants import temp_dir_root

_wdStatisticPages = 2
_wdExportFormatPDF = 17


def _launch_word():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        app.AutomationSecurity = 3  # 禁宏
    except Exception:
        pass
    return app


def _update_all_fields(doc):
    """主文档 + 全部故事区（页眉/页脚/脚注）的域更新。"""
    try:
        doc.Fields.Update()
    except Exception:
        pass
    try:
        story = doc.StoryRanges(1)  # wdMainTextStory
        while story is not None:
            try:
                story.Fields.Update()
            except Exception:
                pass
            story = story.NextStoryRange()
    except Exception:
        pass


def render_check(docx_path: str, export_pdf: str = None,
                 margins_cm: dict = None) -> dict:
    """返回渲染验证结果。任何 COM 层故障 -> {"available": False, "reason": ...}。"""
    docx_path = os.path.abspath(docx_path)
    if not os.path.isfile(docx_path):
        return {"available": False, "reason": f"file not found: {docx_path}"}
    try:
        app = _launch_word()
    except Exception as e:
        return {"available": False, "reason": f"Word COM unavailable: {e}"}

    workdir = tempfile.mkdtemp(prefix="render_", dir=temp_dir_root())
    work_copy = os.path.join(workdir, os.path.basename(docx_path))
    shutil.copy2(docx_path, work_copy)

    doc = None
    try:
        doc = app.Documents.Open(work_copy, AddToRecentFiles=False, Visible=False)
        doc.Repaginate()
        result = {
            "available": True,
            "word_version": str(app.Version),
            "page_count": doc.ComputeStatistics(_wdStatisticPages),
        }

        # 域更新（目录实体化、页码取真实值）
        _update_all_fields(doc)
        doc.Repaginate()
        result["page_count_after_field_update"] = doc.ComputeStatistics(_wdStatisticPages)

        # TOC 验证（§16 的渲染级确认：域存在且更新后有真实条目）
        toc_info = {"tables_of_contents": int(doc.TablesOfContents.Count)}
        if doc.TablesOfContents.Count > 0:
            toc = doc.TablesOfContents(1)
            text = toc.Range.Text or ""
            toc_info["entries"] = int(toc.Range.Paragraphs.Count)
            toc_info["preview"] = text.replace("\r", " ")[:80]
            toc_info["has_real_entries"] = bool(text.strip()) and toc_info["entries"] >= 1
        result["toc"] = toc_info

        # 页码域验证：页脚故事区存在 PAGE 域
        footer_has_page = False
        try:
            for i in range(1, doc.Sections.Count + 1):
                footer = doc.Sections(i).Footers(1)  # wdPrimaryFooterStory
                if footer.Exists:
                    for fld in footer.Range.Fields:
                        if "PAGE" in (fld.Code.Text or "").upper():
                            footer_has_page = True
                            break
                if footer_has_page:
                    break
        except Exception:
            pass
        result["page_number_rendered"] = footer_has_page

        # 可选 PDF 导出（从更新过域的副本导出，目录带真实页码）
        if export_pdf:
            export_pdf = os.path.abspath(export_pdf)
            os.makedirs(os.path.dirname(export_pdf) or ".", exist_ok=True)
            doc.ExportAsFixedFormat(export_pdf, _wdExportFormatPDF)
            result["pdf_exported"] = export_pdf
            # 视觉指标（§14）：客观指标代码算，PNG 交 Agent 目检
            try:
                import visual_check
                png_dir = os.path.join(os.path.dirname(export_pdf), "render_pages")
                result["visual"] = visual_check.check_pdf(
                    export_pdf, margins_cm=margins_cm, png_dir=png_dir)
            except Exception as e:
                result["visual"] = {"available": False, "reason": str(e)}

        return result
    except Exception as e:
        return {"available": False, "reason": f"render check failed: {e}"}
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        # 规范卸载：先释放 COM 引用再收 STA，避免 Word 退出后的 RPC 代理噪音
        doc = None
        app = None
        import gc
        gc.collect()
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass
        shutil.rmtree(workdir, ignore_errors=True)
