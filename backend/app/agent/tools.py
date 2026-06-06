"""3 個 read-only 工具與預設 registry 工廠。

Tool 為純函式 + JSON schema，不直接依賴任何 LLM provider；
由 LLMClient.chat 將 schema 轉為各 provider 的原生格式，
由 AgentOrchestrator 在 LLM 要求呼叫時執行。
"""

import os

from app.agent.registry import Tool, ToolRegistry


def _list_excel_sheets(path: str) -> dict:
    if not path:
        return {"error": "未提供檔案路徑。"}
    if not os.path.isfile(path):
        return {"error": f"檔案不存在: {path}"}
    try:
        import pandas as pd
        sheets = pd.ExcelFile(path).sheet_names
    except Exception as e:
        return {"error": str(e)}
    return {"sheets": list(sheets)}


def _read_excel_columns(path: str, sheet: str = "", header_row: int = 1) -> dict:
    if not path:
        return {"error": "未提供檔案路徑。"}
    if not os.path.isfile(path):
        return {"error": f"檔案不存在: {path}"}
    try:
        header_idx = max(0, int(header_row) - 1)
    except (TypeError, ValueError):
        header_idx = 0
    try:
        import pandas as pd
        df = pd.read_excel(
            path,
            sheet_name=sheet if sheet else 0,
            header=header_idx,
            nrows=0,
        )
    except Exception as e:
        return {"error": str(e)}
    return {
        "sheet": sheet or "(first)",
        "columns": [str(c) for c in df.columns],
    }


def _read_template_variables(word_path: str) -> dict:
    if not word_path:
        return {"error": "未提供 Word 檔案路徑。"}
    if not os.path.isfile(word_path):
        return {"error": f"檔案不存在: {word_path}"}
    try:
        from docxtpl import DocxTemplate
        doc = DocxTemplate(word_path)
        variables = doc.get_undeclared_template_variables()
    except Exception as e:
        return {"error": str(e)}
    return {"variables": sorted(str(v) for v in variables)}


def _read_only_tools() -> list:
    return [
        Tool(
            name="list_excel_sheets",
            description="列出 Excel 檔案中的所有工作表名稱。",
            parameters={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Excel 檔案的完整路徑（.xlsx / .xls）",
                    },
                },
                "required": ["path"],
            },
            func=_list_excel_sheets,
        ),
        Tool(
            name="read_excel_columns",
            description="讀取 Excel 指定工作表的欄位名稱（標題列）。若未提供 sheet 則使用第一個。",
            parameters={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Excel 檔案完整路徑"},
                    "sheet": {"type": "string", "description": "工作表名稱；不指定則讀第一個"},
                    "header_row": {"type": "integer", "description": "標題列編號（1-based，預設 1）"},
                },
                "required": ["path"],
            },
            func=_read_excel_columns,
        ),
        Tool(
            name="read_template_variables",
            description="讀取 Word 範本中所有未替換的 Jinja 範本變數（{{ ... }}）。",
            parameters={
                "type": "object",
                "properties": {
                    "word_path": {
                        "type": "string",
                        "description": "Word 範本檔（.docx）的完整路徑",
                    },
                },
                "required": ["word_path"],
            },
            func=_read_template_variables,
        ),
    ]


def make_writable_tools(ctx) -> list:
    """以 AppContext 包裝的「狀態變更」與「執行」類工具。"""
    return [
        Tool(
            name="get_current_settings",
            description="查看 UI 當前所有設定值（路徑 / 工作表 / 標題列 / 輸出 / 檔名規則 / 圖片寬度等）。",
            parameters={"type": "object", "properties": {}},
            func=ctx.get_settings,
        ),
        Tool(
            name="set_word_path",
            description="設定 Word 範本檔案路徑（絕對路徑、.docx）。",
            parameters={
                "type": "object",
                "properties": {"path": {"type": "string", "description": "Word 範本路徑"}},
                "required": ["path"],
            },
            func=ctx.set_word_path,
        ),
        Tool(
            name="set_excel_path",
            description="設定 Excel 數據檔案路徑；自動刷新工作表清單。",
            parameters={
                "type": "object",
                "properties": {"path": {"type": "string", "description": "Excel 路徑（.xlsx / .xls）"}},
                "required": ["path"],
            },
            func=ctx.set_excel_path,
        ),
        Tool(
            name="set_sheet_name",
            description="設定要使用的 Excel 工作表名稱（空字串代表使用第一個）。",
            parameters={
                "type": "object",
                "properties": {"name": {"type": "string", "description": "工作表名稱"}},
                "required": ["name"],
            },
            func=ctx.set_sheet_name,
        ),
        Tool(
            name="set_header_row",
            description="設定 Excel 標題列編號（1-based）。",
            parameters={
                "type": "object",
                "properties": {"row": {"type": "integer", "description": "標題列編號，預設 1"}},
                "required": ["row"],
            },
            func=ctx.set_header_row,
        ),
        Tool(
            name="set_output_dir",
            description="設定報告輸出資料夾。",
            parameters={
                "type": "object",
                "properties": {"path": {"type": "string", "description": "輸出資料夾路徑"}},
                "required": ["path"],
            },
            func=ctx.set_output_dir,
        ),
        Tool(
            name="set_filename_template",
            description="設定輸出檔名規則。可用 {欄位名} 與 {index} 佔位符（如 {客戶}_{日期}.docx）。",
            parameters={
                "type": "object",
                "properties": {"template": {"type": "string", "description": "檔名模板"}},
                "required": ["template"],
            },
            func=ctx.set_filename_template,
        ),
        Tool(
            name="set_image_width_mm",
            description="設定圖片插入時的預設寬度（毫米）。",
            parameters={
                "type": "object",
                "properties": {"value": {"type": "integer", "description": "寬度，單位 mm"}},
                "required": ["value"],
            },
            func=ctx.set_image_width_mm,
        ),
        Tool(
            name="validate_template",
            description="檢查 Word 範本變數與 Excel 欄位是否一致。回傳 missing_in_excel / extra_in_excel / passed。",
            parameters={"type": "object", "properties": {}},
            func=ctx.validate_template,
        ),
        Tool(
            name="generate_reports",
            description=(
                "依當前設定批次產出所有報告。建議先執行 validate_template；missing 欄位非空時應先告知使用者。"
                "若 AI 引擎頁籤的「啟用審查」開啟，會在每份產出後自動由 reviewer 模型審查；"
                "失敗者複製到 Failed_Reports/ 等下一批處理。"
                "回傳 produced / total / output_dir；啟用審查時另含 reviewed / failed_count / failed[] / failed_dir。"
            ),
            parameters={"type": "object", "properties": {}},
            func=ctx.generate_reports,
        ),
        Tool(
            name="open_output_folder",
            description="於檔案總管開啟當前輸出資料夾。",
            parameters={"type": "object", "properties": {}},
            func=ctx.open_output_folder,
        ),
        Tool(
            name="review_single_docx",
            description=(
                "用 VLM reviewer 審查單一報告(.docx 或 .pptx),依 rubric 回傳 passed / score / issues / suggestions。"
                "需先在「AI 引擎」選好 reviewer 模型;呼叫後會自動渲染成 PNG(docx 每頁 / pptx 每張投影片)送 VLM。"
                "可選 row_context_json(產生報告所用的數據 JSON)讓 reviewer 有上下文比對。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "docx_path": {
                        "type": "string",
                        "description": "Word 檔（.docx）的完整路徑",
                    },
                    "row_context_json": {
                        "type": "string",
                        "description": "可選：產生此份報告所用的資料（JSON 字串）",
                    },
                },
                "required": ["docx_path"],
            },
            func=ctx.review_single_docx,
        ),
        Tool(
            name="read_docx_text",
            description=(
                "讀取 Word 範本的所有段落文字（用於對應前先看內容）。"
                "max_paragraphs > 0 時截斷至前 N 段。"
                "未指定 word_path 則用當前設定。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "word_path": {
                        "type": "string",
                        "description": "Word 路徑；空則用當前設定",
                    },
                    "max_paragraphs": {
                        "type": "integer",
                        "description": "段落數上限；0 = 全部",
                    },
                },
                "required": [],
            },
            func=ctx.read_docx_text,
        ),
        Tool(
            name="rename_template_variable",
            description=(
                "把 Word 範本中的 {{ old }} 全部改成 {{ new }}（容忍空白變化），存檔。"
                "用於把既有變數名對齊 Excel 欄位。回傳 {changed: N}；找不到時 changed=0。"
                "未指定 word_path 則用當前設定。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "old": {"type": "string", "description": "舊變數名（不含 {{ }}）"},
                    "new": {"type": "string", "description": "新變數名（不含 {{ }}）"},
                    "word_path": {"type": "string"},
                },
                "required": ["old", "new"],
            },
            func=ctx.rename_template_variable,
        ),
        Tool(
            name="insert_template_variable",
            description=(
                "在範本中找到 anchor 文字並插入 {{ variable }}（用於從零標註空白範本）。"
                "範例：anchor=「客戶姓名：」、variable=「客戶名稱」、position=after → "
                "「客戶姓名：{{ 客戶名稱 }}」。只插入第一個出現的 anchor。"
                "未指定 word_path 則用當前設定。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "anchor": {
                        "type": "string",
                        "description": "範本中要對齊的文字（例：「客戶姓名：」）",
                    },
                    "variable": {
                        "type": "string",
                        "description": "要插入的變數名（不含 {{ }}）",
                    },
                    "position": {
                        "type": "string",
                        "enum": ["after", "before", "replace"],
                        "description": "插入位置；預設 after",
                    },
                    "word_path": {"type": "string"},
                },
                "required": ["anchor", "variable"],
            },
            func=ctx.insert_template_variable,
        ),
        Tool(
            name="suggest_mappings",
            description=(
                "讓 planner LLM 一次性比對 Word 範本內容與 Excel 欄位，回傳建議的 "
                "renames（變數改名）與 inserts（在某段文字附近插入變數）清單；不會自動套用。"
                "得到建議後請用 ask_user 確認，再呼叫 rename_template_variable / "
                "insert_template_variable 套用。消耗 1 次 planner 預算。"
                "word_path / excel_path 留空時使用當前設定。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "word_path": {"type": "string"},
                    "excel_path": {"type": "string"},
                },
                "required": [],
            },
            func=ctx.suggest_mappings,
        ),
        Tool(
            name="list_folder_files",
            description=(
                "列出任意資料夾中的檔案（用於圖片 / Word / Excel 等批次素材）。"
                "kind 可為 image / word / excel / pdf / any。"
                "回傳 {folder, files: [{name, path, size}], count}。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "folder_path": {"type": "string", "description": "資料夾完整路徑"},
                    "kind": {
                        "type": "string",
                        "enum": ["image", "word", "excel", "pdf", "any"],
                        "description": "檔案類型篩選；預設 image",
                    },
                    "max_files": {
                        "type": "integer",
                        "description": "回傳上限（0 = 全部）",
                    },
                },
                "required": ["folder_path"],
            },
            func=ctx.list_folder_files,
        ),
        Tool(
            name="insert_image_at_anchor",
            description=(
                "在 Word 範本中找到 anchor 文字（出現在哪一段），於該段下面新增一行並插入圖片。"
                "用途：把圖片資料夾的圖貼到 Word 對應位置（如「圖 1：流程圖」下面）。"
                "width_mm 留 0 時用 UI 的圖片寬度設定。未指定 word_path 用當前設定。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "anchor": {
                        "type": "string",
                        "description": "範本中要對齊的文字（圖片插在該段落下方）",
                    },
                    "image_path": {
                        "type": "string",
                        "description": "圖片完整路徑",
                    },
                    "width_mm": {
                        "type": "integer",
                        "description": "圖片寬度（mm）；0 = 用 UI 設定",
                    },
                    "word_path": {"type": "string"},
                },
                "required": ["anchor", "image_path"],
            },
            func=ctx.insert_image_at_anchor,
        ),
        Tool(
            name="suggest_image_placements",
            description=(
                "讓 planner LLM 看 Word 段落 + 圖片檔名，給出建議的「哪張圖放哪段下面」配對；"
                "不會自動套用。回傳 {placements:[{image, image_path, anchor, reason}]}。"
                "得到建議後請用 ask_user 確認，再呼叫 insert_image_at_anchor 套用。"
                "消耗 1 次 planner 預算。word_path 留空用當前設定。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "image_folder": {
                        "type": "string",
                        "description": "存放圖片的資料夾路徑",
                    },
                    "word_path": {"type": "string"},
                },
                "required": ["image_folder"],
            },
            func=ctx.suggest_image_placements,
        ),
        Tool(
            name="render_docx_pages",
            description=(
                "將 docx 檔渲染成每頁一張 PNG，供視覺檢查（reviewer 用）。"
                "需 Windows + Word + pymupdf。回傳 pages: [{page, path}]、output_dir、page_count。"
                "max_pages > 0 時只渲染前 N 頁。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "docx_path": {
                        "type": "string",
                        "description": "Word 檔（.docx）的完整路徑",
                    },
                    "dpi": {
                        "type": "integer",
                        "description": "渲染解析度 DPI；預設 150，不可低於 72",
                    },
                    "max_pages": {
                        "type": "integer",
                        "description": "只渲染前 N 頁；0 = 全部",
                    },
                },
                "required": ["docx_path"],
            },
            func=ctx.render_docx_pages,
        ),
        Tool(
            name="render_pptx_pages",
            description=("將 .pptx 每張投影片渲染成 PNG(供視覺檢查 / 預覽)。需 Windows + PowerPoint。"
                         "回傳 pages:[{page,path}]、page_count。max_slides>0 只渲染前 N 張。"),
            parameters={
                "type": "object",
                "properties": {
                    "pptx_path": {"type": "string", "description": "PPTX 檔完整路徑"},
                    "max_slides": {"type": "integer", "description": "只渲染前 N 張;0=全部"},
                },
                "required": ["pptx_path"],
            },
            func=ctx.render_pptx_pages,
        ),
        Tool(
            name="auto_place_images_visual",
            description=(
                "【視覺全自動配圖】Word 或 PPT 報告『沒有任何 {{標籤}}』時,靠多模態視覺"
                "把資料夾裡的圖片貼到正確位置,全程無需人工。"
                "word_path 為 .docx → 渲染每頁、判斷貼第幾頁哪段、pywin32 貼上;"
                "為 .pptx → 依各投影片文字情境把照片配到對的投影片(空圖片版面配置區優先)。"
                "需先在 AI 引擎選好『reviewer(視覺)模型』。回傳每張圖的配對與插入結果。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "image_folder": {"type": "string", "description": "含圖片的資料夾路徑"},
                    "word_path": {"type": "string", "description": "Word 報告路徑;省略則用目前設定的 word_path"},
                    "width_mm": {"type": "integer", "description": "插入圖片寬度(mm);0=用預設"},
                },
                "required": ["image_folder"],
            },
            func=ctx.auto_place_images_visual,
        ),
        Tool(
            name="fill_report_from_folders",
            description=(
                "【結構化報告填圖】針對『制式範本』(已內含照片、位置固定、有就近標籤的表格,如測試報告)。"
                "把照片根目錄下各子資料夾(排序後 01,02,...)依序對應到範本的各『照片表』,"
                "再依每格『就近標籤』(如 Top side / Bottom corner / Before test)與照片檔名語意配對,"
                "用『置換 docx 內部 media 位元組』方式填入 —— document.xml 完全不動,圖片留原位不錯位、保留版面。"
                "比 auto_place_images_visual 更適合『規律重複、同標籤多次出現』的範本(不需逐圖視覺呼叫)。"
                "建議先用 dry_run=true 取得推導對應給使用者確認(驗證關卡),再正式寫檔。"
                "安全閘:子資料夾數 ≠ 照片表數、或樣本對應不齊,會中止不寫半成品。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "template_path": {"type": "string", "description": "範本 docx(已含照片、結構固定)"},
                    "photo_root": {"type": "string", "description": "照片根目錄;其下每個子資料夾=一個樣本"},
                    "output_path": {"type": "string", "description": "輸出 docx 路徑;省略則範本同目錄 *_filled_agent.docx"},
                    "dry_run": {"type": "boolean", "description": "true=只回推導對應供確認、不寫檔(預設 false)"},
                    "match_mode": {"type": "string", "enum": ["auto", "text", "vlm"],
                                   "description": "配對方式:auto=檔名語意配對+VLM看圖補強(預設,最萬用);text=只用檔名語意;vlm=只用VLM看圖內容(檔名無意義時用)"},
                },
                "required": ["template_path", "photo_root"],
            },
            func=lambda template_path, photo_root, output_path="", dry_run=False, match_mode="auto":
                ctx.fill_report_from_folders(template_path, photo_root,
                                             output_path=output_path, dry_run=dry_run, match_mode=match_mode),
        ),
        Tool(
            name="set_pptx_template_path",
            description="設定 PPTX 範本(含 {{tag}})的完整路徑(.pptx)。",
            parameters={"type": "object", "properties": {
                "path": {"type": "string", "description": "PPTX 範本絕對路徑"}}, "required": ["path"]},
            func=ctx.set_pptx_template_path,
        ),
        Tool(
            name="set_pptx_filename_template",
            description="設定 PPTX 批次產出的檔名規則(可用 {index} 與任意欄位名,如 {客戶}.pptx)。",
            parameters={"type": "object", "properties": {
                "template": {"type": "string", "description": "檔名規則"}}, "required": ["template"]},
            func=ctx.set_pptx_filename_template,
        ),
        Tool(
            name="read_pptx_template_variables",
            description="讀 PPTX 範本內所有 {{tag}}(掃所有投影片的文字框與表格)。",
            parameters={"type": "object", "properties": {}, "required": []},
            func=ctx.read_pptx_template_variables,
        ),
        Tool(
            name="validate_pptx_template",
            description="驗證 PPTX 範本 {{tag}} 與來源 Excel 欄位是否對齊;回 missing / extra / passed。",
            parameters={"type": "object", "properties": {}, "required": []},
            func=ctx.validate_pptx_template,
        ),
        Tool(
            name="generate_pptx_reports",
            description=("依 PPTX 範本(含 {{tag}})+ 來源 Excel,逐列產出新 .pptx(每列一份,保留版面與格式;"
                         "單一標籤值為圖片路徑時就地嵌圖)。需先 set_pptx_template_path / set_excel_path / set_output_dir。"),
            parameters={"type": "object", "properties": {}, "required": []},
            func=ctx.generate_pptx_reports,
        ),
        Tool(
            name="apply_folder_images_static",
            description=(
                "【範本共用靜態圖】把一個資料夾的圖片,依檔名對到範本的 {{圖片欄}} 標籤就地換成圖片"
                "(所有報告共用同一組,如 logo / 固定圖表)。支援 .docx / .pptx / .xlsx 範本。"
                "use_ai=true 用 LLM 依語意配檔名↔標籤,否則用檔名比對。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "image_folder": {"type": "string", "description": "含圖片的資料夾"},
                    "template_path": {"type": "string", "description": "範本路徑(docx/pptx/xlsx);省略用目前 word_path"},
                    "use_ai": {"type": "boolean", "description": "true=LLM 語意配對檔名↔標籤(預設 false=檔名比對)"},
                },
                "required": ["image_folder"],
            },
            func=ctx.apply_folder_images_static,
        ),
        Tool(
            name="fill_per_row_images",
            description=(
                "【每列不同的圖】資料夾照片依檔名對到來源 Excel 某 key 欄位的值,把圖片路徑寫進指定圖片欄,"
                "並把資料來源切到含路徑的副本;之後用 generate_reports/excel/pptx 批次產出時會逐列嵌入各自的圖。"
                "需先 set_excel_path。例:每個客戶一張店面照。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "image_folder": {"type": "string", "description": "含照片的資料夾"},
                    "key_column": {"type": "string", "description": "檔名要對到的 Excel 欄(如 客戶名稱)"},
                    "image_column": {"type": "string", "description": "要寫入圖片路徑的欄/範本標籤名(如 店面照)"},
                },
                "required": ["image_folder", "key_column", "image_column"],
            },
            func=ctx.fill_per_row_images,
        ),
        Tool(
            name="auto_fill_text_visual",
            description=(
                "【無標註填字】把來源 Excel 一列資料的各欄位值,填到『沒有 {{標籤}}』的報告中語意對應的位置。"
                "word_path 為 .docx → 視覺判斷填在標籤後/表格右格(pywin32);"
                "為 .pptx → 接在對應投影片的標籤文字後;為 .xlsx 表單 → 填到標籤格的右/下/同格。"
                "需先設定 excel_path(資料來源)。常與 auto_place_images_visual 一起用:先填字再配圖。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "word_path": {"type": "string", "description": "Word 報告路徑;省略則用目前 word_path"},
                    "row_index": {"type": "integer", "description": "用來源 Excel 第幾列(0 起);預設 0"},
                },
                "required": [],
            },
            func=ctx.auto_fill_text_visual,
        ),
        Tool(
            name="ask_user",
            description=(
                "向使用者顯示對話框詢問補充資訊（短問句）。"
                "缺少必要資料、需要使用者決定 yes/no 或從幾個選項挑一個時呼叫。"
                "使用者回覆字串放在 answer 欄位；按取消或關閉視窗回 cancelled=true。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "question": {
                        "type": "string",
                        "description": "要問使用者的問題（簡短一句）",
                    },
                    "choices": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "可選：候選答案清單。提供時介面顯示為單選 radio；不提供則為自由輸入。",
                    },
                },
                "required": ["question"],
            },
            func=ctx.ask_user,
        ),
        Tool(
            name="request_file",
            description=(
                "開啟檔案 / 資料夾選取對話框讓使用者選一個路徑。"
                "缺少 word_path / excel_path / output_dir 這類路徑型資料時優先用此工具。"
                "回傳 path；使用者取消回 cancelled=true。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "prompt": {
                        "type": "string",
                        "description": "對話框標題或說明（簡短）",
                    },
                    "kind": {
                        "type": "string",
                        "enum": ["word", "excel", "image", "directory", "any"],
                        "description": "要選的類型；word=*.docx, excel=*.xlsx/*.xls, image=圖片, directory=資料夾, any=任意檔",
                    },
                },
                "required": ["prompt"],
            },
            func=ctx.request_file,
        ),
        # ---------- Excel → Excel 整欄搬移 ----------
        Tool(
            name="set_transfer_target_path",
            description=(
                "設定整欄搬移模式的目標 Excel 檔路徑。檔案不存在也可以(會於執行時建立),"
                "存在則自動讀工作表清單。"
            ),
            parameters={
                "type": "object",
                "properties": {"path": {"type": "string", "description": "目標 Excel 路徑"}},
                "required": ["path"],
            },
            func=ctx.set_transfer_target_path,
        ),
        Tool(
            name="set_transfer_target_sheet",
            description="設定整欄搬移目標 Excel 的工作表名稱(空字串=第一個 sheet)。",
            parameters={
                "type": "object",
                "properties": {"name": {"type": "string"}},
                "required": ["name"],
            },
            func=ctx.set_transfer_target_sheet,
        ),
        Tool(
            name="set_transfer_mode",
            description=(
                "設定整欄搬移寫入模式:"
                "append=寫到目標既有資料下方;"
                "overwrite=清掉標題以下舊資料再寫;"
                "fresh=新建目標檔/sheet。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "mode": {
                        "type": "string",
                        "enum": ["append", "overwrite", "fresh"],
                    }
                },
                "required": ["mode"],
            },
            func=ctx.set_transfer_mode,
        ),
        Tool(
            name="set_transfer_column_map",
            description=(
                "設定整欄搬移的欄位對應(來源欄位 → 目標欄位)。"
                "範例:{\"客戶名稱\":\"Customer\",\"業務員\":\"Sales\"}。"
                "目標欄位不存在於目標 Excel 時會自動補在標題列右側。"
            ),
            parameters={
                "type": "object",
                "properties": {
                    "mapping": {
                        "type": "object",
                        "description": "{source_col: target_col} 字典",
                    }
                },
                "required": ["mapping"],
            },
            func=ctx.set_transfer_column_map,
        ),
        Tool(
            name="auto_match_transfer_columns",
            description=(
                "依「同名」自動建立整欄搬移的對應(來源 == 目標)。"
                "目標檔已存在時:只配同名;不存在時:用全部來源欄。回傳 mapping。"
            ),
            parameters={"type": "object", "properties": {}},
            func=ctx.auto_match_transfer_columns,
        ),
        Tool(
            name="transfer_excel_data",
            description=(
                "執行 Excel → Excel 整欄搬移(需先 set_excel_path / set_transfer_target_path / "
                "set_transfer_column_map / set_transfer_mode)。回傳 rows_written / target_path / sheet / mode。"
            ),
            parameters={"type": "object", "properties": {}},
            func=ctx.transfer_excel_data,
        ),
        # ---------- Excel 範本標籤(批次產出 Excel 報告) ----------
        Tool(
            name="set_excel_template_path",
            description="設定 Excel 範本(含 {{tag}})的完整路徑。",
            parameters={
                "type": "object",
                "properties": {"path": {"type": "string"}},
                "required": ["path"],
            },
            func=ctx.set_excel_template_path,
        ),
        Tool(
            name="set_excel_template_sheet",
            description="指定 Excel 範本要用的工作表(空 = 全部 sheet 都掃)。",
            parameters={
                "type": "object",
                "properties": {"name": {"type": "string"}},
                "required": ["name"],
            },
            func=ctx.set_excel_template_sheet,
        ),
        Tool(
            name="set_excel_filename_template",
            description=(
                "Excel 範本批次產出的檔名規則。可用 {欄位名} 與 {index} 佔位符,"
                "預設 \"報告_{index}.xlsx\"。會自動補 .xlsx 副檔名。"
            ),
            parameters={
                "type": "object",
                "properties": {"template": {"type": "string"}},
                "required": ["template"],
            },
            func=ctx.set_excel_filename_template,
        ),
        Tool(
            name="read_excel_template_variables",
            description="讀 Excel 範本內所有 {{tag}}(掃所有 worksheet 所有 cell)。",
            parameters={"type": "object", "properties": {}},
            func=ctx.read_excel_template_variables,
        ),
        Tool(
            name="validate_excel_template",
            description=(
                "驗證 Excel 範本 {{tag}} 與來源 Excel 欄位的對齊。"
                "回傳 missing_in_excel / extra_in_excel / passed。"
            ),
            parameters={"type": "object", "properties": {}},
            func=ctx.validate_excel_template,
        ),
        Tool(
            name="generate_excel_reports",
            description=(
                "依 Excel 範本(含 {{tag}}) + 來源 Excel,逐列產出新 Excel(保留範本樣式 / 公式 / 合併儲存格)。"
                "建議先 validate_excel_template 確認對齊。回傳 produced / total / output_dir。"
            ),
            parameters={"type": "object", "properties": {}},
            func=ctx.generate_excel_reports,
        ),
    ]


def build_default_registry(context=None) -> ToolRegistry:
    """建立預設工具集。

    無 context（無法操作 UI 狀態）時，僅註冊 read-only 工具；
    傳入 AppContext 時，加入設定 / 驗證 / 執行類工具。
    """
    reg = ToolRegistry()
    for tool in _read_only_tools():
        reg.register(tool)
    if context is not None:
        for tool in make_writable_tools(context):
            reg.register(tool)
    return reg
