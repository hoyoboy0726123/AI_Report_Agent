"""HeadlessContext:無 Tk 版的 AppContext。

原本 app/agent/context.py 把所有狀態綁在 customtkinter 變數上(self._app.xxx.get())。
這裡改成持有一個純 dict(settings),所有 agent 工具與 REST 端點共用同一份狀態,
方法簽名與回傳結構完全對齊原 AppContext,讓 app/agent/tools.py 不必修改即可運作。

互動型工具(ask_user / request_file)在無介面環境改由注入的 interaction callback 處理;
未提供時回傳 cancelled,讓 agent 流程優雅停止而非崩潰。
"""

import os
import threading

from app.settings import DEFAULTS, save_settings


class HeadlessContext:
    def __init__(self, settings: dict = None, budget=None, interaction=None):
        # settings:可變 dict,直接當作真實狀態來源
        self.settings = dict(DEFAULTS)
        if settings:
            self.settings.update(settings)
        self.budget = budget
        # interaction(kind, payload) -> dict;kind ∈ {"ask_user","request_file"}
        self._interaction = interaction
        self._lock = threading.Lock()
        self.is_generating = False
        self.cancel_event = None

    # ---------- 持久化 ----------

    def _persist(self):
        try:
            save_settings(self.settings)
        except Exception:
            pass

    def _set(self, key, value):
        with self._lock:
            self.settings[key] = value
        self._persist()

    def _g(self, key, default=""):
        return self.settings.get(key, default)

    def _int(self, key, default):
        try:
            return max(1, int(float(self.settings.get(key, default))))
        except (TypeError, ValueError):
            return default

    # ---------- 讀取 ----------

    def get_settings(self) -> dict:
        return {
            "word_path": self._g("word_path"),
            "excel_path": self._g("excel_path"),
            "sheet_name": self._g("sheet_name"),
            "header_row": self._int("header_row", 1),
            "output_dir": self._g("output_dir"),
            "filename_template": self._g("filename_template"),
            "image_width_mm": self._int("image_width_mm", 80),
            "grid_columns": self._int("grid_columns", 2),
        }

    # ---------- 寫入(Word 報告流程) ----------

    def set_word_path(self, path: str) -> dict:
        if not path:
            return {"error": "未提供路徑"}
        if not os.path.isfile(path):
            return {"error": f"檔案不存在: {path}"}
        if not path.lower().endswith(".docx"):
            return {"error": "Word 範本必須是 .docx"}
        self._set("word_path", path)
        return {"ok": True, "value": path}

    def set_excel_path(self, path: str) -> dict:
        if not path:
            return {"error": "未提供路徑"}
        if not os.path.isfile(path):
            return {"error": f"檔案不存在: {path}"}
        if not path.lower().endswith((".xlsx", ".xls")):
            return {"error": "Excel 必須是 .xlsx 或 .xls"}
        self._set("excel_path", path)
        return {"ok": True, "value": path}

    def set_sheet_name(self, name: str) -> dict:
        self._set("sheet_name", name or "")
        return {"ok": True, "value": name or ""}

    def set_header_row(self, row) -> dict:
        try:
            r = max(1, int(row))
        except (TypeError, ValueError):
            return {"error": "row 必須是正整數"}
        self._set("header_row", r)
        return {"ok": True, "value": r}

    def set_output_dir(self, path: str) -> dict:
        if not path:
            return {"error": "未提供路徑"}
        self._set("output_dir", path)
        return {"ok": True, "value": path}

    def set_filename_template(self, template: str) -> dict:
        if not template:
            return {"error": "模板不可為空"}
        self._set("filename_template", template)
        return {"ok": True, "value": template}

    def set_image_width_mm(self, value) -> dict:
        try:
            v = max(1, int(float(value)))
        except (TypeError, ValueError):
            return {"error": "value 必須是正整數（mm）"}
        self._set("image_width_mm", v)
        return {"ok": True, "value": v}

    # ---------- 產生器 ----------

    def _build_generator(self):
        from app.generator import ReportGenerator
        return ReportGenerator(
            word_path=self._g("word_path"),
            excel_path=self._g("excel_path"),
            output_dir=self._g("output_dir") or "Generated_Reports",
            sheet_name=self._g("sheet_name") or None,
            header_row=self._int("header_row", 1),
            filename_template=self._g("filename_template"),
            image_width_mm=self._int("image_width_mm", 80),
        )

    def validate_template(self) -> dict:
        if not self._g("word_path") or not self._g("excel_path"):
            return {"error": "請先設定 Word 與 Excel 路徑"}
        try:
            missing, extra = self._build_generator().validate()
        except Exception as e:
            return {"error": str(e)}
        return {
            "missing_in_excel": sorted(missing),
            "extra_in_excel": sorted(extra),
            "passed": not missing,
        }

    def generate_reports(self, progress_callback=None) -> dict:
        """產出 Word 報告;啟用審查時逐份交給 reviewer 模型評分。

        progress_callback(produced, total, saved_path, row_dict):選用,供 REST 串流進度。
        """
        if self.is_generating:
            return {"error": "已有生成任務在執行中"}
        if not self._g("word_path") or not self._g("excel_path"):
            return {"error": "請先設定 Word 與 Excel 路徑"}

        cancel_event = threading.Event()
        self.cancel_event = cancel_event
        self.is_generating = True
        try:
            if bool(self._g("enable_review", True)):
                return self._generate_with_review(cancel_event, progress_callback)
            return self._generate_simple(cancel_event, progress_callback)
        except Exception as e:
            return {"error": str(e)}
        finally:
            self.is_generating = False

    def _generate_simple(self, cancel_event, progress_callback) -> dict:
        generator = self._build_generator()
        produced = 0
        total = 0
        for prod, tot, saved_path, row_dict in generator.generate_iter(cancel_event=cancel_event):
            produced, total = prod, tot
            if progress_callback:
                progress_callback(prod, tot, saved_path, row_dict)
        return {
            "produced": produced,
            "total": total,
            "output_dir": self._g("output_dir"),
            "cancelled": cancel_event.is_set(),
        }

    def _generate_with_review(self, cancel_event, progress_callback) -> dict:
        import random

        from app.agent.reviewer import move_to_failed_reports, review_report
        from app.config import FAILED_REPORTS_DIR

        vlm, model, err = self._build_reviewer_client()
        if err:
            # reviewer 不可用時退回單純產出,並附註
            res = self._generate_simple(cancel_event, progress_callback)
            res["review_skipped"] = err
            return res

        rubric = self._g("review_rubric", "")
        sampling = self._int("review_sampling_percent", 100)
        output_dir = self._g("output_dir") or "."
        if os.path.isabs(output_dir):
            failed_dir = os.path.join(os.path.dirname(output_dir), FAILED_REPORTS_DIR)
        else:
            failed_dir = FAILED_REPORTS_DIR

        generator = self._build_generator()
        produced = total = reviewed = 0
        failed = []
        budget = self.budget
        budget_exhausted = False

        for prod, tot, saved_path, row_dict in generator.generate_iter(cancel_event=cancel_event):
            produced, total = prod, tot
            if progress_callback:
                progress_callback(prod, tot, saved_path, row_dict)

            if random.random() * 100 > sampling:
                continue
            if budget is not None and not budget.can_use_reviewer():
                budget_exhausted = True
                continue

            result = review_report(vlm, saved_path, row_dict, rubric, model, max_pages=4)
            if budget is not None:
                budget.use_reviewer()
            reviewed += 1
            if "error" in result:
                continue
            if not result.get("passed"):
                try:
                    target = move_to_failed_reports(saved_path, failed_dir)
                except Exception:
                    target = saved_path
                failed.append({
                    "index": prod,
                    "path": target,
                    "score": result.get("score", 0),
                    "issues": result.get("issues", [])[:5],
                })

        result_dict = {
            "produced": produced,
            "total": total,
            "reviewed": reviewed,
            "failed_count": len(failed),
            "failed": failed[:10],
            "output_dir": self._g("output_dir"),
            "failed_dir": failed_dir,
            "cancelled": cancel_event.is_set(),
        }
        if budget_exhausted:
            result_dict["review_budget_exhausted"] = True
        return result_dict

    def review_single_docx(self, docx_path: str, row_context_json: str = "") -> dict:
        """審查單份報告(VLM 依 rubric 評分)。docx → 渲染頁面;pptx → 渲染投影片。"""
        if not docx_path:
            return {"error": "未提供檔案路徑"}
        if not os.path.isfile(docx_path):
            return {"error": f"檔案不存在: {docx_path}"}
        low = docx_path.lower()
        if not (low.endswith(".docx") or low.endswith(".pptx")):
            return {"error": "必須是 .docx 或 .pptx 檔"}
        vlm, model, err = self._build_reviewer_client()
        if err:
            return {"error": err}
        budget = self.budget
        if budget is not None and not budget.can_use_reviewer():
            return {"error": f"已達 reviewer 預算上限 {budget.reviewer_limit}"}
        rubric = self._g("review_rubric", "")
        row_context = {}
        if row_context_json:
            try:
                import json as _json
                row_context = _json.loads(row_context_json)
                if not isinstance(row_context, dict):
                    row_context = {"data": row_context}
            except Exception:
                row_context = {"raw": row_context_json}
        if low.endswith(".pptx"):
            from app.agent.reviewer import review_pptx
            result = review_pptx(vlm, docx_path, row_context, rubric, model, max_slides=6)
        else:
            from app.agent.reviewer import review_report
            result = review_report(vlm, docx_path, row_context, rubric, model, max_pages=4)
        if budget is not None and "error" not in result:
            budget.use_reviewer()
        return result

    def open_output_folder(self) -> dict:
        path = self._g("output_dir")
        if not path:
            return {"error": "尚未設定輸出資料夾"}
        os.makedirs(path, exist_ok=True)
        try:
            os.startfile(os.path.abspath(path))  # Windows
            ok = True
        except Exception:
            ok = False
        return {"ok": ok, "path": path}

    # ---------- LLM client 建構 ----------

    def _build_reviewer_client(self):
        provider = self._g("llm_provider", "Gemini")
        if provider == "Gemini":
            from app.agent.llm import GeminiClient
            client = GeminiClient()
            model = self._g("gemini_reviewer_model")
        else:
            from app.agent.llm import OllamaClient
            client = OllamaClient(endpoint=self._g("ollama_endpoint"),
                                  num_ctx=self._g("ollama_num_ctx"))
            model = self._g("ollama_reviewer_model")
        if not client.is_available():
            return client, model, f"{provider} reviewer 不可用（檢查 API key / endpoint）"
        if not model:
            return client, model, f"未選 {provider} reviewer 模型"
        return client, model, None

    def _build_planner_client(self):
        provider = self._g("llm_provider", "Gemini")
        if provider == "Gemini":
            from app.agent.llm import GeminiClient
            client = GeminiClient()
            model = self._g("gemini_planner_model")
        else:
            from app.agent.llm import OllamaClient
            client = OllamaClient(endpoint=self._g("ollama_endpoint"),
                                  num_ctx=self._g("ollama_num_ctx"))
            model = self._g("ollama_planner_model")
        if not client.is_available():
            return client, model, f"{provider} planner 不可用（檢查 API key / endpoint）"
        if not model:
            return client, model, f"未選 {provider} planner 模型"
        return client, model, None

    # ---------- 範本對應 ----------

    def read_docx_text(self, word_path: str = "", max_paragraphs: int = 0) -> dict:
        from app.agent.template_edit import read_docx_text as _read
        return _read(word_path or self._g("word_path"), max_paragraphs=max_paragraphs)

    def rename_template_variable(self, old: str, new: str, word_path: str = "") -> dict:
        from app.agent.template_edit import rename_template_variable as _rename
        return _rename(word_path or self._g("word_path"), old, new)

    def insert_template_variable(self, anchor: str, variable: str, position: str = "after", word_path: str = "") -> dict:
        from app.agent.template_edit import insert_template_variable as _insert
        return _insert(word_path or self._g("word_path"), anchor, variable, position)

    def suggest_mappings(self, word_path: str = "", excel_path: str = "") -> dict:
        from app.agent.mapping_suggester import suggest_mappings as _suggest
        from app.agent.template_edit import read_docx_text as _read_docx

        wp = word_path or self._g("word_path")
        ep = excel_path or self._g("excel_path")
        if not wp:
            return {"error": "未提供 Word 路徑"}
        if not ep:
            return {"error": "未提供 Excel 路徑"}
        rd = _read_docx(wp, max_paragraphs=80)
        if "error" in rd:
            return {"error": f"讀範本失敗: {rd['error']}"}
        paragraphs = rd.get("paragraphs", [])
        try:
            from docxtpl import DocxTemplate
            tpl_vars = list(DocxTemplate(wp).get_undeclared_template_variables())
        except Exception as e:
            return {"error": f"讀範本變數失敗: {e}"}
        try:
            import pandas as pd
            df = pd.read_excel(ep, sheet_name=self._g("sheet_name") or 0,
                               header=max(0, self._int("header_row", 1) - 1), nrows=0)
            excel_columns = [str(c) for c in df.columns]
        except Exception as e:
            return {"error": f"讀 Excel 欄位失敗: {e}"}
        llm, model, err = self._build_planner_client()
        if err:
            return {"error": err}
        budget = self.budget
        if budget is not None and not budget.can_use_planner():
            return {"error": f"已達 planner 預算上限 {budget.planner_limit}"}
        result = _suggest(llm, paragraphs, tpl_vars, excel_columns, model)
        if "error" not in result and budget is not None:
            budget.use_planner()
        return result

    # ---------- 圖片資料夾 ----------

    def list_folder_files(self, folder_path: str, kind: str = "image", max_files: int = 0) -> dict:
        from app.agent.folder_scan import list_folder_files as _list
        return _list(folder_path, kind=kind, max_files=max_files)

    def insert_image_at_anchor(self, anchor: str, image_path: str, width_mm: int = 0, word_path: str = "") -> dict:
        from app.agent.template_edit import insert_image_at_anchor as _insert
        w = width_mm if width_mm else self._int("image_width_mm", 80)
        return _insert(word_path or self._g("word_path"), anchor, image_path, width_mm=w)

    def suggest_image_placements(self, image_folder: str, word_path: str = "") -> dict:
        import os as _os
        from app.agent.folder_scan import list_folder_files as _list_files
        from app.agent.mapping_suggester import suggest_image_placements as _suggest_imgs
        from app.agent.template_edit import read_docx_text as _read_docx

        wp = word_path or self._g("word_path")
        if not wp:
            return {"error": "未提供 Word 路徑"}
        if not image_folder or not _os.path.isdir(image_folder):
            return {"error": f"資料夾不存在: {image_folder}"}
        rd = _read_docx(wp, max_paragraphs=80)
        if "error" in rd:
            return {"error": f"讀範本失敗: {rd['error']}"}
        paragraphs = rd.get("paragraphs", [])
        listing = _list_files(image_folder, kind="image")
        if "error" in listing:
            return {"error": f"列圖片失敗: {listing['error']}"}
        files = listing.get("files", [])
        if not files:
            return {"error": f"資料夾中沒有圖片: {image_folder}"}
        image_names = [f["name"] for f in files]
        llm, model, err = self._build_planner_client()
        if err:
            return {"error": err}
        budget = self.budget
        if budget is not None and not budget.can_use_planner():
            return {"error": f"已達 planner 預算上限 {budget.planner_limit}"}
        result = _suggest_imgs(llm, paragraphs, image_names, model)
        if "error" not in result and budget is not None:
            budget.use_planner()
        if "placements" in result:
            name_to_path = {f["name"]: f["path"] for f in files}
            for p in result["placements"]:
                p["image_path"] = name_to_path.get(p["image"], "")
        return result

    # ---------- 範本共用靜態圖 / 每列圖片(agent 用) ----------

    def _read_template_vars_any(self, path: str) -> list:
        """依副檔名讀範本 {{標籤}}(docx/pptx/xlsx)。"""
        low = (path or "").lower()
        try:
            if low.endswith(".pptx"):
                from app.pptx_template import PptxReportGenerator
                return sorted(PptxReportGenerator(template_path=path, excel_path="").template_variables())
            if low.endswith((".xlsx", ".xls")):
                from app.excel_template import ExcelReportGenerator
                return sorted(ExcelReportGenerator(template_path=path, excel_path="").template_variables())
            from docxtpl import DocxTemplate
            return sorted(DocxTemplate(path).get_undeclared_template_variables())
        except Exception:
            return []

    def apply_folder_images_static(self, image_folder: str, template_path: str = "",
                                   use_ai: bool = False) -> dict:
        """把資料夾照片依檔名對到範本的 {{圖片欄}},就地換成圖片(所有報告共用)。docx/pptx/xlsx。"""
        import os as _os
        from app.image_mapper import (list_images, deterministic_match, ai_match,
                                      apply_static_to_template)
        tp = template_path or self._g("word_path")
        if not tp or not _os.path.isfile(tp):
            return {"error": f"範本不存在: {tp}"}
        if not image_folder or not _os.path.isdir(image_folder):
            return {"error": f"資料夾不存在: {image_folder}"}
        variables = self._read_template_vars_any(tp)
        if not variables:
            return {"error": "範本沒有 {{標籤}}(無法靜態換圖)"}
        imgs = list_images(image_folder)
        if not imgs:
            return {"error": f"資料夾無圖片: {image_folder}"}
        names = [i["name"] for i in imgs]
        name2path = {i["name"]: i["path"] for i in imgs}
        if use_ai:
            llm, model, err = self._build_planner_client()
            match = deterministic_match(names, variables) if err else ai_match(llm, model, names, variables)
            if isinstance(match, dict) and "error" in match:
                match = deterministic_match(names, variables)
        else:
            match = deterministic_match(names, variables)
        field_to_path = {field: name2path[img] for img, field in match.items() if field}
        if not field_to_path:
            return {"error": "沒有圖片對應到任何範本標籤(檢查檔名 vs 標籤)",
                    "variables": variables, "images": names}
        return apply_static_to_template(tp, field_to_path, width_mm=self._int("image_width_mm", 80))

    def fill_per_row_images(self, image_folder: str, key_column: str, image_column: str) -> dict:
        """每列不同的圖:資料夾照片依檔名對到來源 Excel 的 key 欄值,寫進 image_column(路徑),
        並把資料來源切到含路徑的副本,供後續批次產出逐列嵌圖。"""
        import os as _os
        from app.image_mapper import fill_excel_image_column
        ep = self._g("excel_path")
        if not ep:
            return {"error": "未設定來源 Excel(請先 set_excel_path)"}
        if not image_folder or not _os.path.isdir(image_folder):
            return {"error": f"資料夾不存在: {image_folder}"}
        res = fill_excel_image_column(ep, self._g("sheet_name"), self._int("header_row", 1),
                                      key_column, image_column, image_folder)
        if res.get("output_path"):
            self._set("excel_path", res["output_path"])
        return res

    # ---------- 視覺驅動自動配圖(無標註) ----------

    def auto_place_images_visual(self, image_folder: str, word_path: str = "",
                                 width_mm: int = 0, use_com: bool = True) -> dict:
        """在 Word「無標註」下,靠視覺模型把資料夾圖片貼到正確位置(全自動)。

        消耗 reviewer 預算(每張圖 1 次視覺呼叫)。
        """
        import os as _os
        from app.agent.template_edit import read_docx_text
        from app.agent.visual_placer import auto_place_images_visual as _run

        wp = word_path or self._g("word_path")
        if not wp:
            return {"error": "未提供報告路徑"}
        if not image_folder or not _os.path.isdir(image_folder):
            return {"error": f"資料夾不存在: {image_folder}"}

        vlm, model, err = self._build_reviewer_client()  # 視覺能力在 reviewer 模型
        if err:
            return {"error": err}

        # PPTX 走 PPT 視覺配圖流程
        if wp.lower().endswith(".pptx"):
            from app.pptx_template import auto_place_images_pptx
            llm, lmodel, _e = self._build_planner_client()
            if _e:
                llm, lmodel = None, ""
            base, ext = _os.path.splitext(wp)
            res = auto_place_images_pptx(vlm, model, wp, image_folder, base + "_filled" + ext,
                                         llm=llm, llm_model=lmodel, mode="auto")
            budget = self.budget
            if budget is not None and "error" not in res:
                budget.use_reviewer()
            return res

        rd = read_docx_text(wp, max_paragraphs=0)
        if "error" in rd:
            return {"error": f"讀範本失敗: {rd['error']}"}
        paragraphs = rd.get("paragraphs", [])

        w = width_mm if width_mm else self._int("image_width_mm", 80)
        budget = self.budget
        if budget is not None and not budget.can_use_reviewer():
            return {"error": f"已達 reviewer 預算上限 {budget.reviewer_limit}"}

        rpm = 0
        try:
            rpm = int(self.settings.get("vision_rpm", 0) or 0)
        except (TypeError, ValueError):
            rpm = 0
        result = _run(vlm, model, wp, image_folder, paragraphs, width_mm=w,
                      use_com=use_com, rpm=rpm)
        # 以實際定位的圖片數計預算
        if budget is not None and "error" not in result:
            for _ in range(max(1, len(result.get("placements", [])))):
                if budget.can_use_reviewer():
                    budget.use_reviewer()
        return result

    # ---------- 視覺驅動無標註文字填寫 ----------

    def auto_fill_text_visual(self, word_path: str = "", row_index: int = 0) -> dict:
        """在 Word「無標註」下,靠視覺把 Excel 某一列資料填到正確位置(全自動)。

        row_index:用來源 Excel 的第幾列(0 起)。消耗 reviewer 預算 1 次。
        """
        from app.agent.visual_filler import auto_fill_text_visual as _run

        wp = word_path or self._g("word_path")
        if not wp:
            return {"error": "未提供 Word 路徑"}
        if not self._g("excel_path"):
            return {"error": "未設定來源 Excel"}
        # 讀指定列 → dict
        try:
            import pandas as pd
            df = pd.read_excel(self._g("excel_path"), sheet_name=self._g("sheet_name") or 0,
                               header=max(0, self._int("header_row", 1) - 1))
            if df.empty:
                return {"error": "來源 Excel 沒有資料列"}
            ridx = max(0, min(int(row_index), len(df) - 1))
            row = df.iloc[ridx]
            data = {str(k): row[k] for k in df.columns}
        except Exception as e:
            return {"error": f"讀 Excel 失敗: {e}"}

        # PPTX / XLSX 走文字配對(planner 文字模型即可,不必視覺)
        import os as _os
        low = wp.lower()
        if low.endswith(".pptx") or low.endswith(".xlsx") or low.endswith(".xls"):
            llm, lmodel, _e = self._build_planner_client()
            if _e:
                return {"error": _e}
            base, ext = _os.path.splitext(wp)
            outp = base + "_filled" + ext
            if low.endswith(".pptx"):
                from app.pptx_template import auto_fill_text_pptx
                res = auto_fill_text_pptx(llm, lmodel, wp, data, outp)
            else:
                from app.excel_template import auto_fill_text_xlsx
                res = auto_fill_text_xlsx(llm, lmodel, wp, data, outp)
            budget = self.budget
            if budget is not None and "error" not in res:
                budget.use_planner()
            return res

        vlm, model, err = self._build_reviewer_client()
        if err:
            return {"error": err}
        budget = self.budget
        if budget is not None and not budget.can_use_reviewer():
            return {"error": f"已達 reviewer 預算上限 {budget.reviewer_limit}"}
        rpm = 0
        try:
            rpm = int(self.settings.get("vision_rpm", 0) or 0)
        except (TypeError, ValueError):
            rpm = 0
        result = _run(vlm, model, wp, data, rpm=rpm)
        if budget is not None and "error" not in result:
            budget.use_reviewer()
        return result

    # ---------- 結構化報告填圖(媒體位元組置換) ----------

    def fill_report_from_folders(self, template_path: str, photo_root: str,
                                 output_path: str = "", dry_run: bool = False,
                                 match_mode: str = "auto",
                                 max_dim: int = 1500, quality: int = 85) -> dict:
        """把 photo_root 下各子資料夾的照片,依範本結構(就近標籤)填進範本對應槽位。

        - 子資料夾(排序後)依序對應到範本的各「照片表」(第 1 表=第 1 個子資料夾…)。
        - 用 docx media 位元組置換:document.xml 完全不動,圖片留在原位、不錯位。
        - 安全閘:資料夾數 ≠ 照片表數、或任一樣本對不齊 → 中止回報,不寫半成品。
        - dry_run=True:只回傳推導對應供人確認(驗證關卡),不寫檔。
        """
        import os as _os
        from app.docx_report_filler import (parse_template, parse_template_pptx,
                                            parse_filenames, derive_mapping, place_photos)
        if not template_path or not _os.path.isfile(template_path):
            return {"error": f"範本不存在: {template_path}"}
        if not photo_root or not _os.path.isdir(photo_root):
            return {"error": f"照片根目錄不存在: {photo_root}"}

        # 依副檔名選結構解析器(docx 表格 / pptx 投影片)
        ext = _os.path.splitext(template_path)[1].lower()
        if ext == ".pptx":
            pt = parse_template_pptx(template_path)
        else:
            pt = parse_template(template_path)
        if "error" in pt:
            return {"error": f"解析範本失敗: {pt['error']}"}
        tables = pt.get("tables", [])
        if not tables:
            return {"error": "範本中找不到任何含圖片的表格槽位"}

        # 通用配對:planner(文字)做「標籤↔檔名」語意配對;reviewer(VLM)做「照片內容↔範本示範圖」視覺配對。
        # match_mode: auto(text→VLM 補強) / text / vlm。取不到模型時退回啟發式(備援)。
        llm, model, _e1 = self._build_planner_client()
        if _e1:
            llm, model = None, ""
        vlm, vmodel, _e2 = self._build_reviewer_client()
        if _e2:
            vlm, vmodel = None, ""

        subdirs = sorted(d for d in _os.listdir(photo_root)
                         if _os.path.isdir(_os.path.join(photo_root, d)))
        if not subdirs:
            return {"error": "照片根目錄下沒有子資料夾"}
        # 安全閘:數量對不上
        if len(subdirs) != len(tables):
            return {"error": f"安全閘中止:子資料夾數({len(subdirs)})≠ 範本照片表數({len(tables)});"
                             f"請確認對應關係。子資料夾={subdirs}"}

        from app.docx_report_filler import derive_mapping_semantic_batched, derive_mapping_vlm
        vstrat = self._g("vlm_match_strategy", "batched") or "batched"

        media_to_photo = {}
        samples = []
        review_issues = []

        # 先讀各樣本檔案
        sample_files = [parse_filenames(_os.path.join(photo_root, sub))["files"] for sub in subdirs]

        # 合併呼叫:text / auto 用「一次配完所有樣本」(5→1,少曝險)
        batched = {}
        if match_mode in ("text", "auto") and llm and model:
            b = derive_mapping_semantic_batched(
                [{"key": subdirs[i], "slots": tables[i]["slots"], "files": sample_files[i]}
                 for i in range(len(subdirs))], llm, model)
            if "error" not in b:
                batched = b

        for idx, (tbl, sub) in enumerate(zip(tables, subdirs)):
            files = sample_files[idx]
            if match_mode == "vlm":
                m = derive_mapping_vlm(tbl["slots"], files, vlm, vmodel, template_path, strategy=vstrat) \
                    if (vlm and vmodel) else derive_mapping(tbl["slots"], files)
            else:
                m = batched.get(sub)
                if m is None:  # 合併失敗 → 逐樣本退回
                    m = derive_mapping(tbl["slots"], files, llm=llm, model=model)
                # auto:對不齊的槽位/檔案,用 VLM 看圖補強
                if match_mode == "auto" and m.get("need_review") and vlm and vmodel:
                    un_rids = {u["rid"] for u in m.get("unmatched_slots", [])}
                    un_names = set(m.get("unmatched_files", []))
                    un_slots = [s for s in tbl["slots"] if s["rid"] in un_rids]
                    un_files = [f for f in files if f["name"] in un_names]
                    if un_slots and un_files:
                        v = derive_mapping_vlm(un_slots, un_files, vlm, vmodel, template_path, strategy=vstrat)
                        if "error" not in v and v.get("pairs"):
                            m["pairs"].extend(v["pairs"])
                            done = {p["rid"] for p in m["pairs"]}
                            usedf = {p["file"] for p in m["pairs"]}
                            m["unmatched_slots"] = [u for u in m["unmatched_slots"] if u["rid"] not in done]
                            m["unmatched_files"] = [n for n in m["unmatched_files"] if n not in usedf]
                            m["need_review"] = bool(m["unmatched_slots"] or m["unmatched_files"])
                            m["method"] = (m.get("method", "") + "+vlm")
            samples.append({
                "sample": idx + 1, "folder": sub,
                "table_index": tbl["table_index"], "matched": len(m["pairs"]),
                "slots": len(tbl["slots"]), "method": m.get("method", "heuristic"),
                "unmatched_slots": m["unmatched_slots"],
                "unmatched_files": m["unmatched_files"],
                "pairs": [{"caption": p["caption"], "file": p["file"]} for p in m["pairs"]],
            })
            if m["need_review"]:
                review_issues.append({"folder": sub,
                                      "unmatched_slots": [s["caption"] for s in m["unmatched_slots"]],
                                      "unmatched_files": m["unmatched_files"]})
            for p in m["pairs"]:
                if p["media"]:
                    media_to_photo[p["media"]] = p["file_path"]

        summary = {
            "template": template_path,
            "photo_tables": len(tables),
            "sample_folders": subdirs,
            "total_slots": pt["total_slots"],
            "total_matched": len(media_to_photo),
            "samples": samples,
            "need_review": bool(review_issues),
            "review_issues": review_issues,
        }

        # 驗證關卡:dry_run 只回對應
        if dry_run:
            summary["dry_run"] = True
            return summary
        # 安全閘:有對不齊就不寫
        if review_issues:
            summary["error"] = "安全閘中止:有樣本對應不齊(need_review),未寫檔。請先用 dry_run 檢視。"
            return summary

        if not output_path:
            stem = _os.path.splitext(_os.path.basename(template_path))[0]
            output_path = _os.path.join(_os.path.dirname(template_path), f"{stem}_filled_agent.docx")
        res = place_photos(template_path, media_to_photo, output_path,
                           max_dim=max_dim, quality=quality)
        summary.update(res)
        return summary

    # ---------- 渲染 ----------

    def render_docx_pages(self, docx_path: str, dpi: int = 150, max_pages: int = 0) -> dict:
        from app.agent.docx_render import docx_to_images
        if not docx_path:
            return {"error": "未提供 docx 路徑"}
        if not os.path.isfile(docx_path):
            return {"error": f"檔案不存在: {docx_path}"}
        if not docx_path.lower().endswith(".docx"):
            return {"error": "必須是 .docx 檔"}
        try:
            d = max(72, int(dpi or 150))
        except (TypeError, ValueError):
            d = 150
        try:
            mp = max(0, int(max_pages or 0))
        except (TypeError, ValueError):
            mp = 0
        try:
            pages = docx_to_images(docx_path, dpi=d, max_pages=mp)
        except FileNotFoundError as e:
            return {"error": f"檔案不存在: {e}"}
        except Exception as e:
            return {"error": str(e)}
        if not pages:
            return {"error": "未渲染出任何頁面"}
        return {
            "pages": [{"page": p, "path": path} for p, path in pages],
            "output_dir": os.path.dirname(pages[0][1]),
            "page_count": len(pages),
        }

    def render_pptx_pages(self, pptx_path: str, max_slides: int = 0) -> dict:
        """把 .pptx 每張投影片渲染成 PNG(需本機 PowerPoint)。"""
        from app.pptx_render import pptx_to_images
        if not pptx_path or not os.path.isfile(pptx_path):
            return {"error": f"檔案不存在: {pptx_path}"}
        if not pptx_path.lower().endswith(".pptx"):
            return {"error": "必須是 .pptx 檔"}
        try:
            mp = max(0, int(max_slides or 0))
        except (TypeError, ValueError):
            mp = 0
        try:
            pages = pptx_to_images(pptx_path, max_slides=mp)
        except Exception as e:
            return {"error": f"渲染失敗(需本機 PowerPoint):{e}"}
        if not pages:
            return {"error": "未渲染出任何投影片"}
        return {
            "pages": [{"page": p, "path": path} for p, path in pages],
            "output_dir": os.path.dirname(pages[0][1]),
            "page_count": len(pages),
        }

    # ---------- 互動(human-in-the-loop) ----------

    def ask_user(self, question: str, choices=None) -> dict:
        if not question:
            return {"error": "question 不可為空"}
        if self._interaction:
            return self._interaction("ask_user", {"question": question, "choices": list(choices) if choices else None})
        return {"cancelled": True, "note": "無介面模式不支援互動詢問"}

    def request_file(self, prompt: str, kind: str = "any") -> dict:
        if self._interaction:
            return self._interaction("request_file", {"prompt": prompt, "kind": kind})
        return {"cancelled": True, "note": "無介面模式不支援檔案選取"}

    # ============================================================
    # Excel → Excel 搬移
    # ============================================================

    def set_transfer_target_path(self, path: str) -> dict:
        if not path:
            return {"error": "未提供路徑"}
        if not path.lower().endswith((".xlsx", ".xls")):
            return {"error": "目標必須是 .xlsx 或 .xls"}
        self._set("transfer_target_path", path)
        return {"ok": True, "value": path, "exists": os.path.isfile(path)}

    def set_transfer_target_sheet(self, name: str) -> dict:
        self._set("transfer_target_sheet", name or "")
        return {"ok": True, "value": name or ""}

    def set_transfer_mode(self, mode: str) -> dict:
        from app.excel_transfer import VALID_MODES
        if mode not in VALID_MODES:
            return {"error": f"mode 必須是 {VALID_MODES} 之一"}
        self._set("transfer_mode", mode)
        return {"ok": True, "value": mode}

    def set_transfer_column_map(self, mapping: dict) -> dict:
        if not isinstance(mapping, dict):
            return {"error": "mapping 必須是 {來源欄: 目標欄} 字典"}
        clean = {str(k): str(v) for k, v in mapping.items() if v}
        self._set("transfer_column_map", clean)
        return {"ok": True, "count": len(clean), "value": clean}

    def auto_match_transfer_columns(self) -> dict:
        from app.excel_transfer import ExcelTransfer
        src_cols = ExcelTransfer.list_columns(
            self._g("excel_path"), self._g("sheet_name"), self._int("header_row", 1))
        tgt_cols = ExcelTransfer.list_columns(
            self._g("transfer_target_path"), self._g("transfer_target_sheet"),
            self._int("transfer_target_header_row", 1))
        tgt_set = set(tgt_cols)
        matched = {c: c for c in src_cols if c in tgt_set}
        # 目標無資料時,同名自動對映到自己(會自動建立標題)
        if not tgt_cols:
            matched = {c: c for c in src_cols}
        self._set("transfer_column_map", matched)
        return {"ok": True, "count": len(matched), "value": matched}

    def transfer_excel_data(self) -> dict:
        from app.excel_transfer import ExcelTransfer
        if not self._g("excel_path"):
            return {"error": "未指定來源 Excel"}
        if not self._g("transfer_target_path"):
            return {"error": "未指定目標 Excel"}
        column_map = self._g("transfer_column_map") or {}
        if not column_map:
            return {"error": "欄位對應為空"}
        try:
            t = ExcelTransfer(
                source_path=self._g("excel_path"),
                target_path=self._g("transfer_target_path"),
                column_map=column_map,
                source_sheet=self._g("sheet_name"),
                source_header_row=self._int("header_row", 1),
                target_sheet=self._g("transfer_target_sheet"),
                target_header_row=self._int("transfer_target_header_row", 1),
                mode=self._g("transfer_mode", "append"),
            )
            return t.transfer()
        except Exception as e:
            return {"error": str(e)}

    # ============================================================
    # Excel 範本標籤(批次產出 Excel)
    # ============================================================

    def set_excel_template_path(self, path: str) -> dict:
        if not path:
            return {"error": "未提供路徑"}
        if not os.path.isfile(path):
            return {"error": f"檔案不存在: {path}"}
        if not path.lower().endswith((".xlsx", ".xls")):
            return {"error": "Excel 範本必須是 .xlsx 或 .xls"}
        self._set("excel_template_path", path)
        return {"ok": True, "value": path}

    def set_excel_template_sheet(self, name: str) -> dict:
        self._set("excel_template_sheet", name or "")
        return {"ok": True, "value": name or ""}

    def set_excel_filename_template(self, template: str) -> dict:
        if not template:
            return {"error": "模板不可為空"}
        self._set("excel_filename_template", template)
        return {"ok": True, "value": template}

    def _build_excel_report_generator(self):
        from app.excel_template import ExcelReportGenerator
        return ExcelReportGenerator(
            template_path=self._g("excel_template_path"),
            excel_path=self._g("excel_path"),
            output_dir=self._g("output_dir") or "Generated_Reports",
            sheet_name=self._g("excel_template_sheet"),
            header_row=self._int("header_row", 1),
            filename_template=self._g("excel_filename_template"),
            image_width_px=self._int("excel_image_width_px", 320),
        )

    def read_excel_template_variables(self) -> dict:
        if not self._g("excel_template_path"):
            return {"error": "未指定 Excel 範本"}
        try:
            vars_found = self._build_excel_report_generator().template_variables()
        except Exception as e:
            return {"error": str(e)}
        return {"variables": sorted(str(v) for v in vars_found)}

    def validate_excel_template(self) -> dict:
        if not self._g("excel_template_path") or not self._g("excel_path"):
            return {"error": "請先設定 Excel 範本與來源 Excel"}
        try:
            missing, extra = self._build_excel_report_generator().validate()
        except Exception as e:
            return {"error": str(e)}
        return {
            "missing_in_excel": sorted(missing),
            "extra_in_excel": sorted(extra),
            "passed": not missing,
        }

    # ============================================================
    # PPTX 範本標籤(批次產出 PPTX)
    # ============================================================

    def set_pptx_template_path(self, path: str) -> dict:
        if not path:
            return {"error": "未提供路徑"}
        if not os.path.isfile(path):
            return {"error": f"檔案不存在: {path}"}
        if not path.lower().endswith(".pptx"):
            return {"error": "PPTX 範本必須是 .pptx"}
        self._set("pptx_template_path", path)
        return {"ok": True, "value": path}

    def set_pptx_filename_template(self, template: str) -> dict:
        if not template:
            return {"error": "模板不可為空"}
        self._set("pptx_filename_template", template)
        return {"ok": True, "value": template}

    def _build_pptx_report_generator(self):
        from app.pptx_template import PptxReportGenerator
        return PptxReportGenerator(
            template_path=self._g("pptx_template_path"),
            excel_path=self._g("excel_path"),
            output_dir=self._g("output_dir") or "Generated_Reports",
            sheet_name=self._g("sheet_name"),
            header_row=self._int("header_row", 1),
            filename_template=self._g("pptx_filename_template"),
        )

    def read_pptx_template_variables(self) -> dict:
        if not self._g("pptx_template_path"):
            return {"error": "未指定 PPTX 範本(請先 set_pptx_template_path)"}
        try:
            vars_found = self._build_pptx_report_generator().template_variables()
        except Exception as e:
            return {"error": str(e)}
        return {"variables": sorted(str(v) for v in vars_found)}

    def validate_pptx_template(self) -> dict:
        if not self._g("pptx_template_path") or not self._g("excel_path"):
            return {"error": "請先設定 PPTX 範本與來源 Excel"}
        try:
            missing, extra = self._build_pptx_report_generator().validate()
        except Exception as e:
            return {"error": str(e)}
        return {
            "missing_in_excel": sorted(missing),
            "extra_in_excel": sorted(extra),
            "passed": not missing,
        }

    def generate_pptx_reports(self, progress_callback=None) -> dict:
        if not self._g("pptx_template_path"):
            return {"error": "請先設定 PPTX 範本"}
        if not self._g("excel_path"):
            return {"error": "請先設定來源 Excel"}
        if not self._g("output_dir"):
            return {"error": "請先設定輸出資料夾"}
        cancel_event = threading.Event()
        try:
            gen = self._build_pptx_report_generator()
            produced = total = 0
            for prod, tot, saved_path, row_dict in gen.generate_iter(cancel_event=cancel_event):
                produced, total = prod, tot
                if progress_callback:
                    progress_callback(prod, tot, saved_path, row_dict)
        except Exception as e:
            return {"error": str(e)}
        return {
            "produced": produced,
            "total": total,
            "output_dir": self._g("output_dir"),
            "cancelled": cancel_event.is_set(),
        }

    def generate_excel_reports(self, progress_callback=None) -> dict:
        if not self._g("excel_template_path"):
            return {"error": "請先設定 Excel 範本"}
        if not self._g("excel_path"):
            return {"error": "請先設定來源 Excel"}
        if not self._g("output_dir"):
            return {"error": "請先設定輸出資料夾"}
        cancel_event = threading.Event()
        try:
            gen = self._build_excel_report_generator()
            produced = total = 0
            for prod, tot, saved_path, row_dict in gen.generate_iter(cancel_event=cancel_event):
                produced, total = prod, tot
                if progress_callback:
                    progress_callback(prod, tot, saved_path, row_dict)
        except Exception as e:
            return {"error": str(e)}
        return {
            "produced": produced,
            "total": total,
            "output_dir": self._g("output_dir"),
            "cancelled": cancel_event.is_set(),
        }
