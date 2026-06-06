"""原生 OS 檔案 / 資料夾選取對話框。

backend 在使用者本機執行,因此可以彈出真正的 Windows 檔案總管對話框,
比在瀏覽器裡自製檔案瀏覽器直覺得多。用獨立子行程跑 tkinter,避免與
FastAPI / asyncio 事件迴圈衝突。

重點:
- 對話框強制 topmost + lift + focus,避免跳在瀏覽器後面看不到。
- 子行程強制 UTF-8 輸出,正確帶回含中文的路徑。
"""

import os
import subprocess
import sys

_DIALOG_SCRIPT = r"""
import sys, tkinter as tk
from tkinter import filedialog
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

kind = sys.argv[1] if len(sys.argv) > 1 else "any"
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)
root.update()

opts = {"parent": root}
if kind == "directory":
    p = filedialog.askdirectory(title="選取資料夾", **opts)
elif kind == "word":
    p = filedialog.askopenfilename(title="選取 Word 範本", filetypes=[("Word", "*.docx")], **opts)
elif kind == "excel":
    p = filedialog.askopenfilename(title="選取 Excel", filetypes=[("Excel", "*.xlsx *.xls")], **opts)
elif kind == "pptx":
    p = filedialog.askopenfilename(title="選取 PPTX 範本", filetypes=[("PowerPoint", "*.pptx")], **opts)
elif kind == "save_excel":
    p = filedialog.asksaveasfilename(title="目標 Excel", defaultextension=".xlsx",
                                     filetypes=[("Excel", "*.xlsx")], **opts)
elif kind == "image":
    p = filedialog.askopenfilename(title="選取圖片",
        filetypes=[("Image", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All", "*.*")], **opts)
else:
    p = filedialog.askopenfilename(title="選取檔案", **opts)

sys.stdout.write(p or "")
root.destroy()
"""


def pick(kind: str = "any", timeout: int = 300) -> str:
    """彈出對話框,回傳選取路徑(取消則回傳空字串)。"""
    try:
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        proc = subprocess.run(
            [sys.executable, "-c", _DIALOG_SCRIPT, kind],
            capture_output=True, text=True, encoding="utf-8", timeout=timeout, env=env,
        )
        return (proc.stdout or "").strip()
    except Exception:
        return ""
