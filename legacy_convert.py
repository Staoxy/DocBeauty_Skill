# -*- coding: utf-8 -*-
"""旧格式转换（DESIGN_V2.md Phase 7 / §3.1）：.doc/.wps -> .docx。

仅在 Word/WPS COM 可用时工作（Windows + 已安装 Office）。
不可用时抛 LegacyConvertError，调用方回退为 UNSUPPORTED_FILE + 可操作提示。
"""
import os


class LegacyConvertError(Exception):
    def __init__(self, reason: str, hint: str = ""):
        super().__init__(reason)
        self.reason = reason
        self.hint = hint


def _launch_word():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0  # wdAlertsNone
    try:
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable：不执行宏
    except Exception:
        pass
    return app


def com_available() -> bool:
    try:
        app = _launch_word()
    except Exception:
        return False
    app.Quit()
    return True


def convert_to_docx(src_path: str, dst_path: str) -> dict:
    """用 Word COM 把 .doc/.wps 另存为 .docx（只读打开，不动原文件）。"""
    src_path = os.path.abspath(src_path)
    dst_path = os.path.abspath(dst_path)
    if not os.path.isfile(src_path):
        raise LegacyConvertError(f"source not found: {src_path}")
    try:
        app = _launch_word()
    except Exception as e:
        raise LegacyConvertError(
            f"Word COM unavailable: {e}",
            hint="请安装 Microsoft Word / WPS 后重试，或先在 Word 中将文档另存为 .docx")

    doc = None
    try:
        doc = app.Documents.Open(
            src_path, ReadOnly=True, AddToRecentFiles=False, Visible=False)
        # wdFormatDocumentDefault = 16（.docx）
        doc.SaveAs2(dst_path, FileFormat=16)
        info = {"from": os.path.basename(src_path), "to": os.path.basename(dst_path),
                "pages": doc.ComputeStatistics(2)}
        return info
    except Exception as e:
        raise LegacyConvertError(f"conversion failed: {e}",
                                 hint=f"文件可能已损坏或受保护：{src_path}")
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
