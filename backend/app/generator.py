"""依 Excel 數據與 Word 範本批次產出報告，支援驗證、取消與圖片欄位。"""

import os

import pandas as pd
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage

from app.config import (
    DEFAULT_FILENAME_TEMPLATE,
    DEFAULT_HEADER_ROW,
    DEFAULT_IMAGE_WIDTH_MM,
    DEFAULT_OUTPUT_DIR,
    IMAGE_EXTENSIONS,
)
from app.filename import render_filename


class ReportGenerator:
    """讀取 Excel 表格，逐列以 docxtpl 渲染 Word 範本並儲存。

    額外支援：
    - 指定 sheet 與標題列
    - 自訂檔名規則
    - 產出前驗證範本變數 vs Excel 欄位
    - cancel_event 中途取消
    - Excel 中的圖片路徑欄位自動轉為 InlineImage
    """

    def __init__(
        self,
        word_path,
        excel_path,
        output_dir=DEFAULT_OUTPUT_DIR,
        sheet_name=None,
        header_row=DEFAULT_HEADER_ROW,
        filename_template=DEFAULT_FILENAME_TEMPLATE,
        image_width_mm=DEFAULT_IMAGE_WIDTH_MM,
    ):
        self.word_path = word_path
        self.excel_path = excel_path
        self.output_dir = output_dir
        self.sheet_name = sheet_name if sheet_name else 0
        self.header_row = max(1, int(header_row))
        self.filename_template = filename_template or DEFAULT_FILENAME_TEMPLATE
        self.image_width_mm = image_width_mm

    def list_sheets(self):
        return pd.ExcelFile(self.excel_path).sheet_names

    def template_variables(self):
        doc = DocxTemplate(self.word_path)
        return doc.get_undeclared_template_variables()

    def _read_dataframe(self):
        return pd.read_excel(
            self.excel_path,
            sheet_name=self.sheet_name,
            header=self.header_row - 1,
        )

    def validate(self):
        """檢查範本變數與 Excel 欄位一致性。回傳 (missing, extra)。"""
        template_vars = self.template_variables()
        df = self._read_dataframe()
        excel_cols = {str(c) for c in df.columns}
        missing = template_vars - excel_cols
        extra = excel_cols - template_vars
        return missing, extra

    def _build_context(self, doc, row_data):
        """將圖片路徑欄位自動包裝為 InlineImage，其他欄位原樣傳入。"""
        context = {}
        for key, value in row_data.items():
            if (
                isinstance(value, str)
                and value.lower().endswith(IMAGE_EXTENSIONS)
                and os.path.isfile(value)
            ):
                context[key] = InlineImage(doc, value, width=Mm(self.image_width_mm))
            else:
                context[key] = value
        return context

    def generate_iter(self, cancel_event=None):
        """逐列產出報告，yield (produced, total, saved_path, row_dict)。

        提供給 reviewer 流程（P6）：每存一份就交給 caller 立刻處理。
        """
        df = self._read_dataframe()
        os.makedirs(self.output_dir, exist_ok=True)

        total = len(df)
        produced = 0
        for index, row in df.iterrows():
            if cancel_event is not None and cancel_event.is_set():
                return

            doc = DocxTemplate(self.word_path)
            row_dict = row.to_dict()
            doc.render(self._build_context(doc, row_dict))
            filename = render_filename(self.filename_template, row_dict, index + 1)
            saved_path = os.path.join(self.output_dir, filename)
            doc.save(saved_path)

            produced += 1
            yield produced, total, saved_path, row_dict

    def generate(self, progress_callback=None, cancel_event=None):
        """產出所有報告。

        progress_callback(current, total): 回報進度（UI 用）。
        cancel_event: threading.Event；set 後在下一輪迴圈跳出。
        回傳 (produced, total)。
        """
        produced = 0
        total = 0
        for prod, tot, _saved_path, _row in self.generate_iter(cancel_event=cancel_event):
            produced = prod
            total = tot
            if progress_callback:
                progress_callback(produced, total)
        return produced, total
