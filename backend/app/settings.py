"""使用者設定持久化（路徑、檔名規則、工作表、AI 引擎等）。

注意：API key 不寫進 settings.json，改由 .env 透過 python-dotenv 載入到環境變數。
"""

import json
from pathlib import Path

from app.config import DEFAULT_REVIEW_RUBRIC

SETTINGS_PATH = Path.home() / ".auto_report" / "settings.json"

DEFAULTS = {
    # 既有
    "word_path": "",
    "excel_path": "",
    "output_dir": "Generated_Reports",
    "filename_template": "報告_{index}.docx",
    "sheet_name": "",
    "header_row": 1,
    "image_width_mm": 80,
    "grid_columns": 2,
    # AI 引擎
    "llm_provider": "Gemini",
    "gemini_planner_model": "",
    "gemini_reviewer_model": "",
    "ollama_endpoint": "http://localhost:11434",
    "ollama_planner_model": "",
    "ollama_reviewer_model": "",
    # Ollama context 視窗;太小會塞不下工具 schema 導致 agent 靜默失效
    "ollama_num_ctx": 32768,
    "enable_review": True,
    "review_sampling_percent": 100,
    "max_review_retries": 3,
    "review_rubric": DEFAULT_REVIEW_RUBRIC,
    "max_planner_calls": 50,
    "max_reviewer_calls": 100,
    "appearance_mode": "System",  # System / Dark / Light
    # Excel → Excel 搬移
    "transfer_ui_mode": "column_map",  # column_map / template_tag
    "transfer_target_path": "",
    "transfer_target_sheet": "",
    "transfer_target_header_row": 1,
    "transfer_mode": "append",  # append / overwrite / fresh
    "transfer_column_map": {},  # {source_col: target_col}
    # Excel 範本(用 {{tag}} 批次產出 Excel 報告)
    "excel_template_path": "",
    "excel_template_sheet": "",
    "excel_filename_template": "報告_{index}.xlsx",
    "excel_image_width_px": 320,
    # Hotkey 模式:Ctrl+Shift+M 寫到 Word 還是 Excel 範本
    "hotkey_target": "word",  # word / excel_template
    # 視覺自動配圖:每分鐘呼叫上限(0=依模型自動推定);內部以「上限-1」滑動視窗主動配速
    "vision_rpm": 0,
    # VLM 看圖配對策略:batched=整批送(快,需多圖能力強的模型)/ describe=單張送(慢,弱模型較準)
    "vlm_match_strategy": "batched",
    # PPTX 範本(含 {{tag}})批次產出
    "pptx_template_path": "",
    "pptx_filename_template": "報告_{index}.pptx",
}


def load_settings():
    if not SETTINGS_PATH.exists():
        return DEFAULTS.copy()
    try:
        data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return DEFAULTS.copy()
    merged = DEFAULTS.copy()
    merged.update(data)
    return merged


def save_settings(settings):
    SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    SETTINGS_PATH.write_text(
        json.dumps(settings, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
