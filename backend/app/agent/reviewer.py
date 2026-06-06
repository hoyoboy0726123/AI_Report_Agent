"""VLM Reviewer：把 docx 渲染成圖頁，送給 vision-capable 模型依 rubric 評分。

回傳結構化結果（dict，不拋例外）：
    {passed: bool, score: int, issues: list[str], suggestions: list[str]}
或 error path：
    {error: str, raw?: str}
"""

import json
import os
import shutil

from app.agent.docx_render import cleanup_render_dir, docx_to_images


REVIEW_SYSTEM_PROMPT = """你是文件品質審查員。會收到一份 Word 報告轉成的 PNG 頁面，
以及產生這份報告所用的數據資料。

依使用者提供的 rubric 評分。

回覆必須是 **單一 JSON 物件**，不要任何其他文字、說明或 markdown 標記。格式：
{
  "passed": true | false,
  "score": 1-10 整數,
  "issues": ["問題描述 1", "問題描述 2"],
  "suggestions": ["建議 1", "建議 2"]
}

passed 為 false 的條件：rubric 中任一條未通過、有未替換的 {{...}} 變數、版面破版、
資料明顯有誤等。
"""


def review_report(
    llm,
    docx_path: str,
    row_context: dict,
    rubric: str,
    model: str,
    max_pages: int = 4,
    cleanup: bool = True,
) -> dict:
    """執行單份報告的審查；不拋例外。"""
    if llm is None or not llm.is_available():
        return {"error": "VLM client 不可用"}
    if not docx_path or not os.path.isfile(docx_path):
        return {"error": f"檔案不存在: {docx_path}"}
    if not model:
        return {"error": "未指定 reviewer 模型"}

    try:
        rendered = docx_to_images(docx_path, dpi=120, max_pages=max_pages)
    except Exception as e:
        return {"error": f"渲染失敗: {e}"}

    if not rendered:
        return {"error": "無頁面可審查"}

    output_dir = os.path.dirname(rendered[0][1])
    image_paths = [path for _, path in rendered]
    result = score_image_paths(llm, image_paths, row_context, rubric, model)
    if cleanup:
        cleanup_render_dir(output_dir)
    return result


def score_image_paths(llm, image_paths, row_context, rubric, model) -> dict:
    """給一組已渲染的頁面圖,用 VLM 依 rubric 評分。docx/pptx 共用核心。"""
    if llm is None or not llm.is_available():
        return {"error": "VLM client 不可用"}
    if not image_paths:
        return {"error": "無頁面可審查"}
    if not model:
        return {"error": "未指定 reviewer 模型"}
    user_text = _build_review_prompt(row_context or {}, rubric or "", len(image_paths))
    try:
        text = llm.vision_complete(system=REVIEW_SYSTEM_PROMPT, user_text=user_text,
                                   images=image_paths, model=model)
    except Exception as e:
        return {"error": f"VLM 呼叫失敗: {e}"}
    parsed = _extract_json(text)
    if parsed is None:
        return {"error": "VLM 回覆非有效 JSON", "raw": (text or "")[:500]}
    return {
        "passed": bool(parsed.get("passed")),
        "score": _safe_int(parsed.get("score", 0)),
        "issues": [str(x) for x in (parsed.get("issues") or [])],
        "suggestions": [str(x) for x in (parsed.get("suggestions") or [])],
    }


def review_pptx(llm, pptx_path: str, row_context: dict, rubric: str, model: str,
                max_slides: int = 6, cleanup: bool = True) -> dict:
    """審查 .pptx:渲染投影片 → VLM 依 rubric 評分。"""
    from app.pptx_render import pptx_to_images, cleanup_render_dir as _cl
    if not pptx_path or not os.path.isfile(pptx_path):
        return {"error": f"檔案不存在: {pptx_path}"}
    try:
        rendered = pptx_to_images(pptx_path, max_slides=max_slides)
    except Exception as e:
        return {"error": f"渲染失敗(需本機 PowerPoint):{e}"}
    if not rendered:
        return {"error": "無投影片可審查"}
    out_dir = os.path.dirname(rendered[0][1])
    result = score_image_paths(llm, [p for _, p in rendered], row_context, rubric, model)
    if cleanup:
        _cl(out_dir)
    return result


def _build_review_prompt(row_context, rubric, page_count):
    return f"""=== Rubric ===
{rubric or "（使用者未提供 rubric；請以一般辦公文件品質基本原則評分）"}

=== 該份報告對應的數據 ===
{json.dumps(row_context, ensure_ascii=False, indent=2)}

=== 待審頁面 ===
共 {page_count} 頁；已附在此訊息後面。請逐一比對 rubric 後評分。
"""


def _extract_json(text: str):
    """從 VLM 回覆抽出 JSON 物件。"""
    if not text:
        return None
    text = text.strip()
    # 去掉 ```json ... ``` 之類 fence
    if text.startswith("```"):
        lines = text.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines).strip()

    try:
        return json.loads(text)
    except (json.JSONDecodeError, TypeError):
        pass

    # 退而求其次：找最外層 { ... }
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        try:
            return json.loads(text[start : end + 1])
        except (json.JSONDecodeError, TypeError):
            return None
    return None


def _safe_int(v, default=0):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def move_to_failed_reports(docx_path: str, target_dir: str) -> str:
    """複製失敗報告到 target_dir；同檔名時加序號。回傳目標路徑。"""
    if not target_dir:
        target_dir = "Failed_Reports"
    os.makedirs(target_dir, exist_ok=True)

    base = os.path.basename(docx_path)
    target = os.path.join(target_dir, base)

    counter = 1
    while os.path.exists(target):
        name, ext = os.path.splitext(base)
        target = os.path.join(target_dir, f"{name}_{counter}{ext}")
        counter += 1

    shutil.copy2(docx_path, target)
    return target
