"""docx → 每頁 PNG 的渲染管線（reviewer agent 使用）。

技術選型：
- docx → PDF：Word COM `ExportAsFixedFormat`。透過 `DispatchEx` 開啟一個
  獨立 Word 實例，避免干擾使用者目前手動使用的 Word。
- PDF → PNG：PyMuPDF（>=1.24.3 起官方建議 `import pymupdf`；fitz 為
  legacy alias，仍可用於相容）。
- 不採用 docx2pdf（2021 年後未更新）。

需 Windows + Word + pymupdf；缺其一會以清楚的錯誤訊息回報。
"""

import os
import shutil
import tempfile

# Word.WdSaveFormat.wdFormatPDF
WORD_FORMAT_PDF = 17


def docx_to_images(
    docx_path: str,
    dpi: int = 150,
    output_dir: str = None,
    max_pages: int = 0,
) -> list:
    """渲染 docx 為一張一張 PNG。

    回傳 list of (page_number_1based, png_path)。
    output_dir 不指定則建立暫存目錄；呼叫者負責清理（cleanup_render_dir）。
    max_pages > 0 時只渲染前 N 頁。
    """
    if not docx_path or not os.path.isfile(docx_path):
        raise FileNotFoundError(docx_path)

    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix="docx_review_")
    os.makedirs(output_dir, exist_ok=True)

    pdf_path = os.path.join(output_dir, "_pages.pdf")

    err = _word_export_to_pdf(docx_path, pdf_path)
    if err:
        raise RuntimeError(f"Word COM 轉 PDF 失敗：{err}")

    return _pdf_to_pngs(pdf_path, output_dir, dpi=dpi, max_pages=max_pages)


def _word_export_to_pdf(docx_path: str, pdf_path: str):
    """回傳錯誤訊息字串；None 表示成功。"""
    try:
        import win32com.client
    except ImportError:
        return "win32com 未安裝（僅支援 Windows）"

    word = None
    doc = None
    try:
        # DispatchEx 一律建立新 Word 實例；不要動到使用者已開的 Word
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(
            os.path.abspath(docx_path),
            ReadOnly=True,
            ConfirmConversions=False,
            AddToRecentFiles=False,
        )
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(pdf_path),
            ExportFormat=WORD_FORMAT_PDF,
        )
        if not os.path.isfile(pdf_path):
            return "PDF 未產出"
        return None
    except Exception as e:
        return str(e)
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass


def _pdf_to_pngs(pdf_path: str, output_dir: str, dpi: int = 150, max_pages: int = 0):
    try:
        import pymupdf  # noqa: F401
        mod = pymupdf
    except ImportError:
        try:
            import fitz as mod  # legacy alias，相容舊環境
        except ImportError:
            raise RuntimeError("pymupdf 未安裝；pip install pymupdf")

    pages = []
    doc = mod.open(pdf_path)
    try:
        n = doc.page_count
        if max_pages and max_pages > 0:
            n = min(n, max_pages)
        for i in range(n):
            page = doc.load_page(i)
            pix = page.get_pixmap(dpi=int(dpi))
            png_path = os.path.join(output_dir, f"page_{i + 1:03d}.png")
            pix.save(png_path)
            pages.append((i + 1, png_path))
    finally:
        doc.close()
    return pages


def cleanup_render_dir(directory: str):
    """移除由 docx_to_images 建立的暫存目錄（含 PDF + PNG）。"""
    if directory and os.path.isdir(directory):
        try:
            shutil.rmtree(directory, ignore_errors=True)
        except Exception:
            pass
