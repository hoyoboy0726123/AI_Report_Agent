"""自動範本對應建議：用 LLM 比對 Word 內容 + Excel 欄位，
給出 renames（改名）+ inserts（插入）建議。
"""

import json

from app.agent.llm.base import Message
from app.agent.reviewer import _extract_json


SUGGEST_SYSTEM = """你是 Word 範本對應助理。會收到：
1. Word 範本中的段落文字（list of strings）
2. Word 範本目前的 Jinja 變數（list of strings；可能為空）
3. Excel 欄位名稱（list of strings）

任務：
A. **renames**：找出 Word 範本既有變數中，與 Excel 欄位名稱不一致者，建議改名。
   範例：範本 {{ 客戶 }}、Excel 欄位「客戶名稱」 → rename 客戶 → 客戶名稱。
B. **inserts**：找出範本中還沒標註但**明顯應該標註**的位置（段落中含 「：」、「: 」、
   「____」、「______」 等空白佔位符後面的位置），建議插入對應的 Excel 欄位。
   範例：段落「聯絡電話：____」 + Excel 欄位「電話」 → insert anchor=「聯絡電話：」、
   var=「電話」、position=after。

回覆必須是 **單一 JSON 物件**，不要任何其他文字、說明或 markdown 標記：
{
  "renames": [
    {"from": "客戶", "to": "客戶名稱", "reason": "Excel 用「客戶名稱」"}
  ],
  "inserts": [
    {"anchor": "聯絡電話：", "var": "電話", "position": "after", "reason": "範本有空白等待填入"}
  ]
}

position 可為 "after" / "before" / "replace"；多數情況用 "after"。
信心不足或語意不明時，**寧可不建議也不要亂猜**（renames 與 inserts 可為空陣列）。
"""


def suggest_mappings(llm, word_paragraphs, template_vars, excel_columns, model) -> dict:
    """呼叫 LLM 取得對應建議。回傳 {renames, inserts} 或 {error}。"""
    if llm is None or not llm.is_available():
        return {"error": "LLM client 不可用"}
    if not model:
        return {"error": "未指定 planner 模型"}

    paragraphs = list(word_paragraphs or [])
    if len(paragraphs) > 60:
        paragraphs = paragraphs[:60] + [f"...（其餘 {len(word_paragraphs) - 60} 段省略）"]

    user_text = f"""=== Word 範本段落（共 {len(word_paragraphs or [])} 段）===
{json.dumps(paragraphs, ensure_ascii=False, indent=2)}

=== Word 範本既有變數 ===
{json.dumps(sorted(template_vars or []), ensure_ascii=False)}

=== Excel 欄位 ===
{json.dumps(sorted(excel_columns or []), ensure_ascii=False)}

請依規則回 JSON。
"""

    try:
        resp = llm.chat(
            [
                Message(role="system", text=SUGGEST_SYSTEM),
                Message(role="user", text=user_text),
            ],
            model=model,
        )
    except Exception as e:
        return {"error": f"LLM 呼叫失敗: {e}"}

    text = (resp.text or "").strip()
    parsed = _extract_json(text)
    if parsed is None:
        return {"error": "LLM 回覆非有效 JSON", "raw": text[:500]}

    return {
        "renames": _normalize_renames(parsed.get("renames")),
        "inserts": _normalize_inserts(parsed.get("inserts")),
    }


def _normalize_renames(items):
    out = []
    for x in items or []:
        if not isinstance(x, dict):
            continue
        f = str(x.get("from") or "").strip()
        t = str(x.get("to") or "").strip()
        if not f or not t or f == t:
            continue
        out.append(
            {
                "from": f,
                "to": t,
                "reason": str(x.get("reason") or ""),
            }
        )
    return out


def _normalize_inserts(items):
    out = []
    for x in items or []:
        if not isinstance(x, dict):
            continue
        a = str(x.get("anchor") or "").strip()
        v = str(x.get("var") or "").strip()
        if not a or not v:
            continue
        pos = str(x.get("position") or "after").strip()
        if pos not in ("after", "before", "replace"):
            pos = "after"
        out.append(
            {
                "anchor": a,
                "var": v,
                "position": pos,
                "reason": str(x.get("reason") or ""),
            }
        )
    return out


# ---------- 圖片檔名 → Word 位置（P9） ----------

SUGGEST_IMAGE_SYSTEM = """你是 Word 範本圖片配對助理。會收到：
1. Word 範本中的段落文字（list of strings）
2. 圖片檔名清單（只有檔名與副檔名，沒有圖片內容）

任務：根據**檔名語意**，找出該圖片應該插入的 Word 位置。
- 圖片會插在 anchor 段落下面一行（不是同段落內）。
- anchor 必須是 Word 範本段落清單中的某段（完整文字或具識別性的子字串）。

例：
- 圖片「客戶簽名_流程示意.png」 + 段落「圖 1：流程示意」 → anchor=「圖 1：流程示意」、image=「客戶簽名_流程示意.png」
- 圖片「組織架構.jpg」 + 段落「組織架構圖：」 → anchor=「組織架構圖：」、image=「組織架構.jpg」

回 **單一 JSON 物件**，不要任何其他文字、說明或 markdown 標記：
{
  "placements": [
    {"image": "filename.png", "anchor": "段落文字", "reason": "簡短理由"}
  ]
}

規則：
- 信心不足或圖片找不到對應段落時，**跳過該圖片**（不要硬塞）。
- 同一段落最多對應一張圖片。
- 同一張圖片最多放一個位置。
"""


def suggest_image_placements(llm, word_paragraphs, image_filenames, model) -> dict:
    """LLM 配對圖片檔名與 Word 段落。回傳 {placements} 或 {error}。"""
    if llm is None or not llm.is_available():
        return {"error": "LLM client 不可用"}
    if not model:
        return {"error": "未指定 planner 模型"}

    paragraphs = list(word_paragraphs or [])
    if len(paragraphs) > 80:
        paragraphs = paragraphs[:80] + [
            f"...（其餘 {len(word_paragraphs) - 80} 段省略）"
        ]
    images = list(image_filenames or [])
    if not images:
        return {"placements": []}

    user_text = f"""=== Word 範本段落（共 {len(word_paragraphs or [])} 段）===
{json.dumps(paragraphs, ensure_ascii=False, indent=2)}

=== 圖片檔名 ===
{json.dumps(images, ensure_ascii=False, indent=2)}

請依規則回 JSON。
"""

    try:
        resp = llm.chat(
            [
                Message(role="system", text=SUGGEST_IMAGE_SYSTEM),
                Message(role="user", text=user_text),
            ],
            model=model,
        )
    except Exception as e:
        return {"error": f"LLM 呼叫失敗: {e}"}

    text = (resp.text or "").strip()
    parsed = _extract_json(text)
    if parsed is None:
        return {"error": "LLM 回覆非有效 JSON", "raw": text[:500]}

    return {"placements": _normalize_placements(parsed.get("placements"))}


def _normalize_placements(items):
    out = []
    seen_anchors = set()
    seen_images = set()
    for x in items or []:
        if not isinstance(x, dict):
            continue
        img = str(x.get("image") or "").strip()
        anc = str(x.get("anchor") or "").strip()
        if not img or not anc:
            continue
        if img in seen_images or anc in seen_anchors:
            continue
        seen_images.add(img)
        seen_anchors.add(anc)
        out.append(
            {
                "image": img,
                "anchor": anc,
                "reason": str(x.get("reason") or ""),
            }
        )
    return out
