"""PPTX → 每張投影片一張 PNG(用 PowerPoint COM 匯出)。供視覺配圖 / 預覽用。"""

import os
import tempfile


def pptx_to_images(pptx_path: str, width: int = 1280, height: int = 720, max_slides: int = 0):
    """回傳 [(slide_index_1based, png_path), ...]。需本機 PowerPoint。"""
    if not pptx_path or not os.path.isfile(pptx_path):
        raise FileNotFoundError(pptx_path)
    import win32com.client
    out_dir = tempfile.mkdtemp(prefix="pptxrender_")
    ppt = None
    pages = []
    try:
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        pr = ppt.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        n = pr.Slides.Count
        if max_slides and max_slides > 0:
            n = min(n, max_slides)
        for i in range(1, n + 1):
            p = os.path.join(out_dir, f"slide{i}.png")
            pr.Slides(i).Export(p, "PNG", width, height)
            if os.path.isfile(p):
                pages.append((i, p))
        pr.Close()
        return pages
    finally:
        try:
            if ppt is not None:
                ppt.Quit()
        except Exception:
            pass


def cleanup_render_dir(path):
    import shutil
    try:
        shutil.rmtree(path, ignore_errors=True)
    except Exception:
        pass
