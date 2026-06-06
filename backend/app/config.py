"""應用程式常數與主題設定。集中管理 UI 樣式、預設值與快捷鍵。"""

APPEARANCE_MODE = "System"
COLOR_THEME = "blue"

WINDOW_TITLE = "AI 辦公自動化 - 視覺化映射工具 v1.1"
WINDOW_HEADER = "Office 視覺化對應與自動生成系統"
WINDOW_SIZE = "880x920"

HOTKEY = "ctrl+shift+m"

DEFAULT_OUTPUT_DIR = "Generated_Reports"
DEFAULT_FILENAME_TEMPLATE = "報告_{index}.docx"
DEFAULT_HEADER_ROW = 1
DEFAULT_IMAGE_WIDTH_MM = 80

IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".bmp")

FAILED_REPORTS_DIR = "Failed_Reports"

DEFAULT_REVIEW_RUBRIC = """1. 文件中無未替換的範本佔位符（如 {{...}}、{%...%}）。
2. 欄位填入的值符合該欄位語意（例如「日期」欄為合法日期格式）。
3. 圖片若有指定，已正確嵌入於對應位置且未變形。
4. 版面無明顯破版（文字疊圖、表格錯位等）。
5. 整體風格與範本一致。"""

COLOR_GREEN = "#2ecc71"
COLOR_GREEN_HOVER = "#27ae60"
COLOR_RED = "#e74c3c"
COLOR_BLUE = "#3498db"
COLOR_BLUE_HOVER = "#2980b9"
