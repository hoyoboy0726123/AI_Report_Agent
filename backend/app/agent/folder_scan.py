"""任意資料夾掃描（agent 取得圖片清單等用途）。

設計原則：純讀取、不寫；錯誤皆回 dict 不拋例外。
"""

import os


_KIND_EXTENSIONS = {
    "image": (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"),
    "word": (".docx",),
    "excel": (".xlsx", ".xls"),
    "pdf": (".pdf",),
    "any": None,
}


def list_folder_files(folder_path: str, kind: str = "image", max_files: int = 0) -> dict:
    """列出資料夾中符合 kind 的檔案。

    回傳 {folder, files: [{name, path, size}], count}。
    kind 可為 image / word / excel / pdf / any。
    max_files > 0 時截斷至前 N 個。
    """
    if not folder_path:
        return {"error": "未提供資料夾路徑"}
    if not os.path.isdir(folder_path):
        return {"error": f"資料夾不存在或不是目錄: {folder_path}"}

    if kind not in _KIND_EXTENSIONS:
        return {
            "error": f"未知 kind: {kind}（可用：{', '.join(_KIND_EXTENSIONS.keys())}）"
        }
    exts = _KIND_EXTENSIONS[kind]

    files = []
    try:
        names = sorted(os.listdir(folder_path))
    except Exception as e:
        return {"error": f"讀取資料夾失敗: {e}"}

    for name in names:
        full = os.path.join(folder_path, name)
        if not os.path.isfile(full):
            continue
        if exts is not None and not name.lower().endswith(exts):
            continue
        try:
            size = os.path.getsize(full)
        except OSError:
            size = 0
        files.append({"name": name, "path": full, "size": size})

    total = len(files)
    if max_files and max_files > 0 and total > max_files:
        files = files[:max_files]

    return {
        "folder": folder_path,
        "kind": kind,
        "files": files,
        "count": total,
    }
