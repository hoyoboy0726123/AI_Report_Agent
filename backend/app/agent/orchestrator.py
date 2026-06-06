"""Agent planner loop。

P2 範圍：read-only 工具 + tool calling 對話流。
P3+ 會擴充寫入類工具、ask_user、retry escalation 等。
"""

import json
import threading
from typing import Iterator, Optional

from app.agent.llm.base import LLMClient, Message
from app.agent.registry import ToolRegistry


SYSTEM_BASE = """你是辦公自動化助理，協助使用者操作 Excel/Word 與批次產報告。

可用工具分為七類（請按情境選用，不要憑空作答）：
- 查詢：get_current_settings、list_excel_sheets、read_excel_columns、
        read_template_variables、read_docx_text
- 設定：set_word_path、set_excel_path、set_sheet_name、set_header_row、
        set_output_dir、set_filename_template、set_image_width_mm
- 範本對應：suggest_mappings（一鍵建議 renames + inserts）、
            rename_template_variable（改名）、insert_template_variable（插入）
- 圖片資料夾：list_folder_files（列任意資料夾檔案）、suggest_image_placements
            （依檔名語意配對 Word 位置）、insert_image_at_anchor（把圖貼到指定段落下面）
- 驗證：validate_template
- 執行：generate_reports（啟用審查時自動由 reviewer 模型評每份；失敗的搬到
        Failed_Reports/）、open_output_folder
- 審查：review_single_docx（單獨審查任一 docx）、render_docx_pages（轉每頁 PNG）
- 互動：ask_user、request_file

行為準則：
- 開始任務前，先呼叫 get_current_settings 看當前 UI 狀態，已有資料就不要再問。
- 缺少 word_path / excel_path / output_dir 這類路徑型資料時，優先呼叫
  request_file 讓使用者直接選；不要叫使用者把路徑打字回覆。
- 需要使用者決定（例如缺欄位仍要繼續嗎？挑哪個工作表？檔名規則？）時，
  用 ask_user 問；若有有限選項，務必提供 choices 讓使用者單選。
- 任一互動工具回傳 cancelled=true 時，視為使用者放棄，立刻停止流程並回報。
- 接到「全部產出」「跑完報告」這類執行指令時，順序：
  1) get_current_settings 確認路徑齊全；缺者用 request_file 補。
  2) validate_template 檢查欄位對齊。
  3) 若 missing_in_excel 非空，用 ask_user 問是否仍要繼續（提供 choices=
     ["是，仍然產出", "否，先處理 Excel"]）；不要直接 generate。
  4) 通過後再呼叫 generate_reports；產出後可主動建議 open_output_folder。
- 當 generate_reports 回傳 failed_count > 0 時，告知使用者哪些報告被搬到
  Failed_Reports/ 與主要 issues；不要自動重產，讓使用者下批人工處理。
- 接到「幫我把標籤對好」「自動對應」「自動標註範本」這類指令時，順序：
  1) 確認 word_path / excel_path / sheet_name 已設定（缺者用 request_file 補）。
  2) 呼叫 suggest_mappings 取得 {renames, inserts} 建議；可能為空陣列。
  3) 用 ask_user 一筆一筆（或合併）確認要套用哪些建議；提供 choices。
  4) 通過確認的 renames 用 rename_template_variable 套用；inserts 用
     insert_template_variable 套用。
  5) 套用完呼叫 validate_template 確認對齊；若仍有 missing 欄位告知使用者。
- 任何寫入類工具（rename / insert / set_*）回傳 error 或 changed=0 時，先告知
  使用者，不要在錯誤上重試。
- 接到「把圖片放到範本對應位置」「資料夾裡的圖貼到 Word」這類指令時，順序：
  1) 確認 word_path 已設定；缺者用 request_file 補。
  2) 用 request_file(kind="directory") 讓使用者選圖片資料夾。
  3) 呼叫 suggest_image_placements 取得 placements 建議；可能為空陣列。
  4) 用 ask_user 確認要套用哪些 placement（提供 choices）。
  5) 通過確認的逐一呼叫 insert_image_at_anchor 套用；失敗逐筆告知。
- 工具回傳含 "error" 欄位即代表失敗，先告知使用者問題並停止；不要重試同一錯誤。
- 完成且無待辦時，最後一句以「DONE」結尾。
"""


class AgentOrchestrator:
    """單回合 = 一次 user 訊息進，跑完所有 tool call 直到 LLM 給最終文字。"""

    def __init__(
        self,
        llm: LLMClient,
        registry: ToolRegistry,
        model: str,
        context: Optional[dict] = None,
        max_iters: int = 8,
        budget=None,
    ):
        self.llm = llm
        self.registry = registry
        self.model = model
        self.context = context or {}
        self.max_iters = max_iters
        self.budget = budget
        self._cancel = threading.Event()
        self.messages = [Message(role="system", text=self._build_system_prompt())]

    # ---- public ----

    def reset(self):
        self._cancel.clear()
        self.messages = [Message(role="system", text=self._build_system_prompt())]

    def cancel(self):
        self._cancel.set()

    def add_user_message(self, text: str):
        self._cancel.clear()
        self.messages.append(Message(role="user", text=text))

    def step(self) -> Iterator[Message]:
        """處理已加入的 user 訊息：呼叫 LLM、執行工具、再呼叫 LLM…直到拿到最終文字。"""
        for _ in range(self.max_iters):
            if self._cancel.is_set():
                yield Message(role="assistant", text="[已中止]")
                return

            if self.budget is not None and not self.budget.can_use_planner():
                msg = Message(
                    role="assistant",
                    text=f"[已達 planner 預算上限 {self.budget.planner_limit}，請至「AI 引擎」頁籤重置計數或提高上限]",
                )
                self.messages.append(msg)
                yield msg
                return

            try:
                resp = self.llm.chat(
                    self.messages,
                    model=self.model,
                    tools=self.registry.schemas(),
                )
                if self.budget is not None:
                    self.budget.use_planner()
            except Exception as e:
                err = Message(role="assistant", text=f"[LLM 錯誤] {e}")
                self.messages.append(err)
                yield err
                return

            self.messages.append(resp)
            yield resp

            if not resp.tool_calls:
                return  # 最終文字回覆

            for tc in resp.tool_calls:
                if self._cancel.is_set():
                    yield Message(role="assistant", text="[已中止]")
                    return
                result = self.registry.run(tc.name, tc.arguments)
                tool_msg = Message(
                    role="tool",
                    text=json.dumps(result, ensure_ascii=False, default=str),
                    tool_name=tc.name,
                )
                self.messages.append(tool_msg)
                yield tool_msg

        yield Message(role="assistant", text=f"[已達迴圈上限 {self.max_iters}]")

    # ---- internal ----

    def _build_system_prompt(self):
        prompt = SYSTEM_BASE
        if self.context:
            prompt += "\n當前 UI 已選設定（你可直接使用，不需再問使用者）："
            for key, value in self.context.items():
                if value not in (None, "", 0):
                    prompt += f"\n- {key}: {value}"
        return prompt
