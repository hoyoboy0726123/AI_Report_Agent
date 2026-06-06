"""視覺驅動的「無標註」文字填寫:Word 沒有任何 {{標籤}},靠版面 + 標籤語意,
把 Excel 每個欄位的值填到正確位置。

與 visual_placer(配圖)對稱,差別在落點型態:
  - after_label:值接在標籤文字後面(如「客戶名稱:」→ 後面補值)
  - next_cell  :值填進標籤右邊的表格儲存格(如表格 |撰寫人|(空)| → 填右格)

做法:把報告每頁 PNG + 段落/儲存格文字 + 整包資料,一次送視覺模型,
回傳每個欄位的 {anchor, placement};再用 pywin32 逐一寫入。
"""

import json
import os
import time

from app.agent.docx_render import cleanup_render_dir, docx_to_images
from app.agent.visual_placer import _anchor_variants, _extract_json, _get_limiter


FILL_SYSTEM_PROMPT = """你是文件填寫助理。會收到一份「沒有任何 {{標籤}}/佔位符」的 Word 報告(每頁 PNG + 文字),
以及一批要填入的「欄位:值」資料。請依版面與標籤語意,判斷每個欄位的值該填在文件的哪個位置。

對每個欄位回傳:
- anchor:文件中真實存在、用來定位的「標籤文字」(通常是該欄位的名稱或提示,如「客戶名稱:」「撰寫人」)。必須是頁面上連續且夠獨特的一小段文字。
- placement:
    "after_label" = 值要接在 anchor 文字後面(同一行,如「客戶名稱:____」)
    "next_cell"   = anchor 在表格內,值要填到它右邊那一格(如表格 |撰寫人|(空格)|)

只回傳單一 JSON:
{"fills": [{"field": "欄位名", "anchor": "定位用標籤文字", "placement": "after_label" 或 "next_cell"}]}
找不到合適位置的欄位就略過(不要硬填)。
"""


def auto_fill_text_visual(vlm, model, word_path, data, rpm=0, retries=2):
    """data = {欄位: 值}。回傳結果摘要 + timings。"""
    t0 = time.time()
    timings = {}
    if vlm is None or not vlm.is_available():
        return {"error": "視覺模型不可用(檢查 AI 引擎 reviewer 模型 / API key)"}
    if not model:
        return {"error": "未指定視覺(reviewer)模型"}
    if not word_path or not os.path.isfile(word_path):
        return {"error": f"Word 不存在: {word_path}"}
    data = {str(k): ("" if v is None else str(v)) for k, v in (data or {}).items()}
    data = {k: v for k, v in data.items() if v.strip()}
    if not data:
        return {"error": "沒有可填的資料"}

    t = time.time()
    try:
        rendered = docx_to_images(word_path, dpi=110)
    except Exception as e:
        return {"error": f"渲染頁面失敗(需本機 Word):{e}"}
    if not rendered:
        return {"error": "未渲染出任何頁面"}
    render_dir = os.path.dirname(rendered[0][1])
    page_pngs = [p for _, p in rendered]
    timings["render_pages"] = round(time.time() - t, 1)

    # 視覺判定(一次呼叫處理所有欄位)
    t = time.time()
    limiter = _get_limiter(model, override_rpm=rpm)
    fills = _locate_fills(vlm, model, page_pngs, data, limiter, retries)
    timings["vision_locate"] = round(time.time() - t, 1)
    cleanup_render_dir(render_dir)

    if isinstance(fills, dict) and fills.get("error"):
        return {"error": fills["error"], "timings": timings}
    if not fills:
        return {"error": "視覺模型未能判定任何欄位落點", "timings": timings}

    # 套上實際值
    for f in fills:
        f["value"] = data.get(f.get("field", ""), "")

    t = time.time()
    ins = _fill_via_com(word_path, [f for f in fills if f.get("value")])
    timings["fill"] = round(time.time() - t, 1)
    timings["total"] = round(time.time() - t0, 1)

    return {
        "fills": fills,
        "filled": ins.get("filled", []),
        "fill_failed": ins.get("failed", []),
        "method": ins.get("method", ""),
        "filled_count": len(ins.get("filled", [])),
        "total_fields": len(data),
        "timings": timings,
    }


def _locate_fills(vlm, model, page_pngs, data, limiter, retries):
    user_text = (
        "以下是要填入報告的資料(欄位:值):\n"
        + "\n".join(f"- {k}:{v}" for k, v in data.items())
        + f"\n\n報告共 {len(page_pngs)} 頁(附後)。請判斷每個欄位的值該填在哪裡,回傳 JSON。"
    )
    last_err = ""
    for attempt in range(retries + 1):
        try:
            if limiter is not None:
                limiter.acquire()
            text = vlm.vision_complete(system=FILL_SYSTEM_PROMPT, user_text=user_text,
                                       images=page_pngs, model=model)
            parsed = _extract_json(text)
            if parsed is not None:
                fills = parsed.get("fills") if isinstance(parsed, dict) else parsed
                if isinstance(fills, list):
                    return [{"field": str(f.get("field", "")).strip(),
                             "anchor": str(f.get("anchor", "")).strip(),
                             "placement": f.get("placement", "after_label")}
                            for f in fills if isinstance(f, dict) and f.get("anchor")]
            last_err = "回覆非預期 JSON"
        except Exception as e:
            last_err = str(e)[:120]
        if attempt < retries:
            time.sleep(1.5 * (attempt + 1))
    return {"error": f"視覺判定失敗(重試 {retries} 次):{last_err}"}


def _fill_via_com(word_path, fills):
    """pywin32:Find 定位 anchor,依 placement 寫入值。"""
    try:
        import win32com.client
    except Exception as e:
        return {"error": f"pywin32 不可用: {e}"}

    WD_WITHIN_TABLE = 12
    word = None
    filled, failed = [], []
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(os.path.abspath(word_path))
        for f in fills:
            anchor, value, placement = f["anchor"], f["value"], f.get("placement", "after_label")
            try:
                rng = None
                for variant in _anchor_variants(anchor):
                    cand = doc.Content
                    fd = cand.Find
                    fd.ClearFormatting()
                    if fd.Execute(variant):
                        rng = cand
                        break
                if rng is None:
                    failed.append({"field": f["field"], "reason": f"找不到錨點: {anchor[:20]}"})
                    continue

                in_table = False
                try:
                    in_table = bool(rng.Information(WD_WITHIN_TABLE))
                except Exception:
                    in_table = False

                if placement == "next_cell" and in_table:
                    cell = rng.Cells(1)
                    table = rng.Tables(1)
                    r, c = cell.RowIndex, cell.ColumnIndex
                    try:
                        nxt = table.Cell(r, c + 1)
                    except Exception:
                        nxt = None
                    if nxt is not None:
                        # 清掉原本(可能是空白/底線)再寫值
                        nxt.Range.Text = value
                        filled.append({"field": f["field"], "anchor": anchor[:20], "where": f"表格({r},{c+1})"})
                    else:
                        rng.Collapse(0)
                        rng.InsertAfter(" " + value)
                        filled.append({"field": f["field"], "anchor": anchor[:20], "where": "標籤後(無右格)"})
                else:
                    rng.Collapse(0)  # wdCollapseEnd
                    rng.InsertAfter(value)
                    filled.append({"field": f["field"], "anchor": anchor[:20], "where": "標籤後"})
            except Exception as e:
                failed.append({"field": f["field"], "reason": str(e)})
        doc.Save()
        doc.Close(False)
        return {"filled": filled, "failed": failed, "method": "pywin32(Word COM)"}
    except Exception as e:
        return {"error": f"Word COM 失敗: {e}"}
    finally:
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
