"""Word + Excel COM 連線、標籤映射、圖片插入與還原。"""

import os

import win32com.client

# Word/Office 常數：用於 LockAspectRatio
MSO_TRUE = -1
PT_PER_MM = 2.834645669


class OfficeMapper:
    """連動執行中的 Word 與 Excel；提供標籤、圖片插入與範圍刪除。"""

    @staticmethod
    def _word():
        return win32com.client.GetActiveObject("Word.Application")

    @staticmethod
    def _excel():
        return win32com.client.GetActiveObject("Excel.Application")

    @staticmethod
    def insert_tag_at_cursor(header_row=1):
        """讀取 Excel 當前選取欄的標題，於 Word 游標插入 {{ tag }}。

        回傳 (success, message, range)。range 為 (start, end)，可用於後續復原；
        失敗時 range 為 None，message 為錯誤訊息。
        """
        try:
            word = OfficeMapper._word()
            excel = OfficeMapper._excel()
        except Exception:
            return False, "連線失敗: 確保 Word 與 Excel 已開啟。", None

        try:
            col_index = excel.Selection.Column
            tag = excel.ActiveSheet.Cells(header_row, col_index).Value
            if not tag:
                return False, "找不到欄位標題（標題列為空）。", None

            text = f"{{{{ {tag} }}}}"
            start = word.Selection.Start
            word.Selection.TypeText(text)
            end = word.Selection.Start
            return True, str(tag), (start, end)
        except Exception as e:
            return False, f"映射失敗: {e}", None

    @staticmethod
    def insert_image_at_cursor(image_path, width_mm=80):
        """從檔案總管選到的圖片，於 Word 游標位置插入。

        回傳 (success, message, range)；range 用於復原。
        """
        if not os.path.isfile(image_path):
            return False, f"找不到圖片檔: {image_path}", None
        try:
            word = OfficeMapper._word()
        except Exception:
            return False, "連線失敗: 確保 Word 已開啟。", None

        try:
            shape = word.ActiveDocument.InlineShapes.AddPicture(
                FileName=image_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Range=word.Selection.Range,
            )
            if width_mm:
                shape.LockAspectRatio = MSO_TRUE
                shape.Width = width_mm * PT_PER_MM

            shape_range = shape.Range
            return True, os.path.basename(image_path), (shape_range.Start, shape_range.End)
        except Exception as e:
            return False, f"圖片插入失敗: {e}", None

    @staticmethod
    def insert_image_grid_at_cursor(image_paths, columns=2, width_mm=80):
        """將多張圖片以 N x M 表格網格形式插入 Word 游標位置。

        回傳 (success, message, range)；range 為整個表格的 (start, end)，可整體復原。
        """
        valid = [p for p in image_paths if p and os.path.isfile(p)]
        if not valid:
            return False, "未選取有效圖片。", None

        try:
            word = OfficeMapper._word()
        except Exception:
            return False, "連線失敗: 確保 Word 已開啟。", None

        try:
            cols = max(1, int(columns))
            rows = (len(valid) + cols - 1) // cols

            table = word.ActiveDocument.Tables.Add(
                word.Selection.Range, rows, cols
            )
            for i, img_path in enumerate(valid):
                cell = table.Cell(i // cols + 1, i % cols + 1)
                shape = word.ActiveDocument.InlineShapes.AddPicture(
                    FileName=img_path,
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Range=cell.Range,
                )
                if width_mm:
                    shape.LockAspectRatio = MSO_TRUE
                    shape.Width = width_mm * PT_PER_MM

            tbl_range = table.Range
            return (
                True,
                f"已插入 {len(valid)} 張圖片 ({cols}×{rows} 網格)",
                (tbl_range.Start, tbl_range.End),
            )
        except Exception as e:
            return False, f"網格插入失敗: {e}", None

    @staticmethod
    def remove_range(start, end):
        """刪除 Word 中 [start, end) 範圍的內容（用於復原映射或圖片）。"""
        try:
            word = OfficeMapper._word()
            word.ActiveDocument.Range(start, end).Delete()
            return True, None
        except Exception as e:
            return False, str(e)

    # ---- Excel 範本標籤模式(讀來源欄位 / 寫到 Excel 範本 cell) ----

    @staticmethod
    def read_excel_active_column_header(header_row=1):
        """讀 Excel 當前 Selection 所在欄的標題列值;回傳 (success, tag_or_msg)。"""
        try:
            excel = OfficeMapper._excel()
        except Exception:
            return False, "連線失敗: 確保 Excel 已開啟。"
        try:
            col_index = excel.Selection.Column
            tag = excel.ActiveSheet.Cells(header_row, col_index).Value
            if not tag:
                return False, "找不到欄位標題(標題列為空)。"
            return True, str(tag)
        except Exception as e:
            return False, f"讀取失敗: {e}"

    @staticmethod
    def insert_tag_at_excel_active_cell(tag):
        """把 {{ tag }} 寫到 Excel 當前 ActiveCell;cell 已有內容則前面追加。

        透過 Excel COM 寫入,避免與已開啟檔案產生鎖檔衝突。
        回傳 (success, message, ref)。ref = (workbook_fullname, sheet_name, cell_address)
        供未來復原使用。
        """
        if not tag:
            return False, "tag 不可為空。", None
        try:
            excel = OfficeMapper._excel()
        except Exception:
            return False, "連線失敗: 確保 Excel 已開啟。", None
        try:
            cell = excel.ActiveCell
            placeholder = f"{{{{ {tag} }}}}"
            existing = cell.Value
            if existing is None or str(existing).strip() == "":
                cell.Value = placeholder
            else:
                cell.Value = f"{existing}{placeholder}"
            ws_name = cell.Worksheet.Name
            wb_full = cell.Worksheet.Parent.FullName
            addr = cell.Address.replace("$", "")
            return True, f"{ws_name}!{addr} ← {tag}", (wb_full, ws_name, addr)
        except Exception as e:
            return False, f"寫入失敗: {e}", None
