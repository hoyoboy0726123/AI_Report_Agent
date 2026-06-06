"""視覺驅動的自動配圖:在 Word「無標註」下,靠多模態視覺把資料夾圖片貼到正確位置。

針對多頁(20~30 頁)+ 表格優化:
  1. 用 Word COM 取得「每頁文字」,以圖檔名關鍵字先定位候選頁碼(快、省 token)。
  2. 只把候選頁的 PNG + 圖片本身送視覺模型確認 anchor(找不到候選才退回全部頁)。
  3. pywin32 一次貼上;若 anchor 落在表格儲存格內,圖片寬度自適應該儲存格(鎖比例)。
     非表格則用指定 width_mm。無本機 Word 時退回 python-docx。

回傳 dict,含各階段耗時(timings),不拋例外。
"""

import json
import os
import re
import threading
import time
from collections import deque

from app.agent.docx_render import cleanup_render_dir, docx_to_images
from app.image_mapper import list_images, _norm


# ============================================================
# 主動配速器(滑動 60 秒視窗)
# ============================================================
# Gemini/Gemma 的 RPM 是看「過去 60 秒內的請求數」。
# 我們在視窗內累積到 (上限-1) 時就主動停下,等最舊那筆滿 60 秒、騰出空位再續,
# 全程不碰到上限 → 不會 429。對使用者只是「慢一點但一定跑完」,不中斷。

def _model_rpm(model: str) -> int:
    """各模型在免費層的概略 RPM(查 ai.google.dev / AI Studio 後設保守值)。"""
    m = (model or "").lower()
    if "flash-lite" in m or m.startswith("gemma"):
        return 30
    if "pro" in m:
        return 5
    if "flash" in m:
        return 15
    return 10


class _RateLimiter:
    """滑動視窗配速:每 60 秒最多 max_per_min 次;達標則阻塞到空出名額。"""

    def __init__(self, max_per_min: int):
        self.max = max(1, int(max_per_min))
        self.calls = deque()
        self.lock = threading.Lock()

    def acquire(self):
        """取得一個名額;必要時 sleep 等視窗清出空位。回傳實際等待秒數。"""
        waited = 0.0
        while True:
            with self.lock:
                now = time.time()
                while self.calls and now - self.calls[0] >= 60.0:
                    self.calls.popleft()
                if len(self.calls) < self.max:
                    self.calls.append(now)
                    return waited
                # 視窗已滿 → 算出最舊那筆還要多久滿 60 秒
                sleep_for = 60.0 - (now - self.calls[0]) + 0.1
            sleep_for = max(0.2, min(sleep_for, 60.0))
            time.sleep(sleep_for)
            waited += sleep_for


_LIMITERS = {}
_LIMITERS_LOCK = threading.Lock()


def _get_limiter(model: str, override_rpm: int = 0) -> "_RateLimiter":
    """取得(或建立)該模型的程序級配速器。override_rpm>0 時優先用。

    實際上限 = (RPM - 1),即『撞到上限前一格就冷卻』。
    """
    rpm = int(override_rpm) if override_rpm and override_rpm > 0 else _model_rpm(model)
    effective = max(1, rpm - 1)
    key = (model or "", effective)
    with _LIMITERS_LOCK:
        lim = _LIMITERS.get(key)
        if lim is None:
            lim = _RateLimiter(effective)
            _LIMITERS[key] = lim
        return lim


LOCATE_SYSTEM_PROMPT = """你是會看圖的文件排版助理。使用者要把一張圖片放進一份 Word 報告,
報告裡「沒有」任何 {{標籤}} 或佔位符,你必須靠視覺判斷正確位置。

你會收到:要放置的「目標圖片」(第一張),接著是幾頁候選頁面的 PNG,以及這些頁面的段落文字。
請判斷這張圖片最該緊接在哪一段「頁面上真實存在」的文字後面。

重要規則:
- 若頁面有表格,且某個儲存格明顯是為這張圖預留的(例如含「照片」「此格放」「圖」等字樣),
  優先選「該儲存格內的文字」當 anchor —— 圖片要放進那個儲存格。
- anchor 必須是「單一連續、真實存在」的一小段文字(8~30 字),不可把不同段落或不同儲存格的文字拼在一起。
- 不要包含換行;盡量挑最獨特、最不會重複的片段。

只回傳單一 JSON 物件:
{"page": 頁碼整數, "anchor": "插入點緊鄰的真實連續文字", "reason": "一句話理由", "confidence": 0~1}
找不到合適位置則 anchor 回空字串。
"""


# ---------- 每頁文字(Word COM) ----------

def _page_texts_via_com(word_path):
    """用 Word COM 取得 {頁碼: 該頁文字}(含表格內文字)。失敗回 {}。"""
    try:
        import win32com.client
    except Exception:
        return {}
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(os.path.abspath(word_path), ReadOnly=True)
        pages = {}
        WD_ACTIVE_END_PAGE = 3  # wdActiveEndPageNumber
        for para in doc.Paragraphs:
            try:
                txt = para.Range.Text
                if not txt or not txt.strip():
                    continue
                pg = int(para.Range.Information(WD_ACTIVE_END_PAGE))
                pages.setdefault(pg, []).append(txt.strip())
            except Exception:
                continue
        doc.Close(False)
        return {pg: " ".join(parts) for pg, parts in pages.items()}
    except Exception:
        return {}
    finally:
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass


def _candidate_pages(stem, page_texts, top=2):
    """以檔名 stem 對每頁文字評分,回傳最相關的前 top 個頁碼。"""
    s = _norm(stem)
    if not s or not page_texts:
        return []
    scores = {}
    grams = {s[i:i + 2] for i in range(len(s) - 1)} if len(s) >= 2 else {s}
    for pg, txt in page_texts.items():
        nt = _norm(txt)
        if not nt:
            continue
        if s in nt:
            scores[pg] = 100 + len(s)
            continue
        hit = sum(1 for g in grams if g and g in nt)
        if hit:
            scores[pg] = hit
    return sorted(scores, key=lambda p: -scores[p])[:top]


def _locate_one(vlm, model, page_pngs, page_texts_subset, image_path, image_name,
                retries=2, limiter=None):
    """送候選頁 + 圖片給視覺模型,回傳 {page, anchor, reason, confidence} 或 {error}。

    呼叫前先過 limiter(主動配速,不會撞 RPM);萬一仍遇暫時性 429 再退避重試保險。
    """
    para_block = "\n".join(f"[第{pg}頁] {txt[:600]}" for pg, txt in page_texts_subset)
    user_text = (
        f"目標圖片檔名:{image_name}\n"
        f"以下提供 {len(page_pngs)} 個候選頁面(圖片在前,頁面在後)及其文字。\n\n"
        f"=== 候選頁文字 ===\n{para_block}\n\n"
        "請判斷這張圖片該緊接哪段文字之後,回傳 JSON。"
    )
    images = [image_path] + [p for _, p in page_pngs]
    last_err = ""
    for attempt in range(retries + 1):
        try:
            if limiter is not None:
                limiter.acquire()  # 主動配速:必要時在此等待視窗清出名額
            text = vlm.vision_complete(system=LOCATE_SYSTEM_PROMPT, user_text=user_text,
                                       images=images, model=model)
            parsed = _extract_json(text)
            if parsed is not None:
                return {
                    "page": _safe_int(parsed.get("page", 0)),
                    "anchor": str(parsed.get("anchor", "")).strip(),
                    "reason": str(parsed.get("reason", "")),
                    "confidence": parsed.get("confidence", 0),
                }
            last_err = "回覆非 JSON"
        except Exception as e:
            last_err = str(e)[:120]
        if attempt < retries:
            time.sleep(1.5 * (attempt + 1))  # 退避:1.5s, 3s
    return {"error": f"視覺定位失敗(重試 {retries} 次):{last_err}"}


def auto_place_images_visual(vlm, model, word_path, image_folder,
                             paragraphs, width_mm=80, use_com=True, rpm=0):
    """主流程:關鍵字定頁 → 視覺確認 → 批次貼上(表格自適應)。回傳結果摘要 + timings。

    rpm:該模型每分鐘上限(0=依模型自動推定)。內部以「上限-1」滑動視窗主動配速,
    全程不碰 RPM、不會 429;名額用完時自動等待視窗清出再續(慢但穩、不中斷)。
    """
    t0 = time.time()
    timings = {}
    if vlm is None or not vlm.is_available():
        return {"error": "視覺模型不可用(檢查 AI 引擎 reviewer 模型 / API key)"}
    if not model:
        return {"error": "未指定視覺(reviewer)模型"}
    if not word_path or not os.path.isfile(word_path):
        return {"error": f"Word 不存在: {word_path}"}
    images = list_images(image_folder)
    if not images:
        return {"error": f"資料夾無圖片: {image_folder}"}

    # 1) 渲染全部頁(一次)
    t = time.time()
    try:
        rendered = docx_to_images(word_path, dpi=110)
    except Exception as e:
        return {"error": f"渲染頁面失敗(需本機 Word):{e}"}
    if not rendered:
        return {"error": "未渲染出任何頁面"}
    render_dir = os.path.dirname(rendered[0][1])
    png_by_page = {pg: path for pg, path in rendered}
    total_pages = len(rendered)
    timings["render_pages"] = round(time.time() - t, 1)

    # 2) 每頁文字(用於關鍵字定頁)
    t = time.time()
    page_texts = _page_texts_via_com(word_path)
    timings["page_texts"] = round(time.time() - t, 1)

    # 3) 逐張:關鍵字定頁 → 只送候選頁給視覺(主動配速,不撞 RPM)
    t = time.time()
    limiter = _get_limiter(model, override_rpm=rpm)

    def _work(im):
        cand = _candidate_pages(_stem(im["name"]), page_texts, top=2) if page_texts else []
        if cand:
            page_pngs = [(pg, png_by_page[pg]) for pg in cand if pg in png_by_page]
            subset = [(pg, page_texts.get(pg, "")) for pg in cand]
        else:
            use = sorted(png_by_page)[:10]
            page_pngs = [(pg, png_by_page[pg]) for pg in use]
            subset = [(pg, page_texts.get(pg, "")) for pg in use]
        loc = _locate_one(vlm, model, page_pngs, subset, im["path"], im["name"], limiter=limiter)
        loc.update({"image": im["name"], "image_path": im["path"], "candidate_pages": cand})
        return loc

    # 並發控制:太高會撞 API 速率上限(RPM)導致暫時性失敗;3 是速度與穩定的平衡點
    from concurrent.futures import ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=min(3, len(images))) as ex:
        placements = list(ex.map(_work, images))
    timings["vision_locate"] = round(time.time() - t, 1)

    cleanup_render_dir(render_dir)

    valid = [p for p in placements if p.get("anchor") and not p.get("error")]
    if not valid:
        return {"error": "視覺模型未能為任何圖片定位出有效錨點",
                "placements": placements, "timings": timings, "total_pages": total_pages}

    # 4) 批次插入(表格自適應)
    t = time.time()
    if use_com:
        ins = _insert_via_com(word_path, valid, width_mm)
        if ins.get("error"):
            ins = _insert_via_docx(word_path, valid, width_mm)
            ins["fallback"] = "python-docx"
    else:
        ins = _insert_via_docx(word_path, valid, width_mm)
    timings["insert"] = round(time.time() - t, 1)
    timings["total"] = round(time.time() - t0, 1)

    return {
        "placements": placements,
        "inserted": ins.get("inserted", []),
        "insert_failed": ins.get("failed", []),
        "method": ins.get("method", ""),
        "fallback": ins.get("fallback"),
        "placed_count": len(ins.get("inserted", [])),
        "total_images": len(images),
        "total_pages": total_pages,
        "timings": timings,
    }


def _insert_via_com(word_path, placements, width_mm):
    """pywin32:Find 定位 anchor → 插入圖片;在表格內則自適應儲存格寬度。"""
    try:
        import win32com.client
    except Exception as e:
        return {"error": f"pywin32 不可用: {e}"}

    MSO_TRUE = -1
    PT_PER_MM = 2.834645669
    WD_WITHIN_TABLE = 12  # wdWithInTable
    word = None
    inserted, failed = [], []
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(os.path.abspath(word_path))
        for p in placements:
            anchor = p["anchor"]
            img = p["image_path"]
            try:
                rng = None
                for variant in _anchor_variants(anchor):
                    cand_rng = doc.Content
                    f = cand_rng.Find
                    f.ClearFormatting()
                    if f.Execute(variant):
                        rng = cand_rng
                        break
                if rng is None:
                    failed.append({"image": p["image"], "reason": f"找不到錨點: {anchor[:24]}"})
                    continue

                in_table = False
                cell_w = None
                try:
                    in_table = bool(rng.Information(WD_WITHIN_TABLE))
                    if in_table:
                        cell = rng.Cells(1)
                        cell_w = float(cell.Width)  # points
                except Exception:
                    in_table = False

                # 在錨點後插入新段落,圖片放新段落(表格內則同格新行)
                rng.Collapse(0)  # wdCollapseEnd
                rng.InsertParagraphAfter()
                rng.Collapse(0)
                shape = doc.InlineShapes.AddPicture(
                    FileName=os.path.abspath(img), LinkToFile=False,
                    SaveWithDocument=True, Range=rng)
                shape.LockAspectRatio = MSO_TRUE
                if in_table and cell_w and cell_w > 10:
                    shape.Width = max(20.0, cell_w - 10)   # 自適應儲存格(留邊)
                    fit = f"cell({round(cell_w)}pt)"
                else:
                    shape.Width = width_mm * PT_PER_MM
                    fit = f"{width_mm}mm"
                inserted.append({"image": p["image"], "anchor": anchor[:24],
                                 "page": p.get("page"), "in_table": in_table, "fit": fit})
            except Exception as e:
                failed.append({"image": p["image"], "reason": str(e)})
        doc.Save()
        doc.Close(False)
        return {"inserted": inserted, "failed": failed, "method": "pywin32(Word COM)"}
    except Exception as e:
        return {"error": f"Word COM 失敗: {e}"}
    finally:
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass


def _insert_via_docx(word_path, placements, width_mm):
    """無 Word 時退路:python-docx 在 anchor 段落後插圖(不支援表格自適應)。"""
    from app.agent.template_edit import insert_image_at_anchor
    inserted, failed = [], []
    for p in placements:
        r = insert_image_at_anchor(word_path, p["anchor"], p["image_path"], width_mm=width_mm)
        if r.get("error"):
            failed.append({"image": p["image"], "reason": r["error"]})
        else:
            inserted.append({"image": p["image"], "anchor": p["anchor"][:24], "page": p.get("page")})
    return {"inserted": inserted, "failed": failed, "method": "python-docx"}


def _anchor_variants(anchor):
    """從 anchor 衍生多個候選搜尋字串,提高 Find 命中率。

    處理:模型把『標籤格 + 圖片格』或多段文字拼成一個 anchor 的情況 ——
    拆成各片段,並去除全形/半形括號,長片段優先。
    """
    seen, out = set(), []

    def add(s):
        s = (s or "").strip()
        if len(s) >= 2 and s not in seen:  # 中文標籤常 2~3 字(撰寫人/備註),門檻放寬
            seen.add(s)
            out.append(s)

    add(anchor)
    # 去括號類符號
    stripped = re.sub(r"[\[\]（）()【】「」［］<>]", " ", anchor)
    add(stripped)
    # 依空白 / 標點切片段,長的先試
    parts = re.split(r"[\s，,。:：;；、\[\]（）()【】「」［］]+", stripped)
    for seg in sorted(parts, key=len, reverse=True):
        add(seg)
    # 最後退路:前 12 字
    if len(anchor) > 12:
        add(anchor[:12])
    return out


def _stem(name):
    return os.path.splitext(os.path.basename(name))[0]


def _extract_json(text):
    if not text:
        return None
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```\w*\n?", "", text)
        text = re.sub(r"\n?```\s*$", "", text)
    try:
        return json.loads(text)
    except Exception:
        s, e = text.find("{"), text.rfind("}")
        if s >= 0 and e > s:
            try:
                return json.loads(text[s:e + 1])
            except Exception:
                return None
    return None


def _safe_int(v, default=0):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default
