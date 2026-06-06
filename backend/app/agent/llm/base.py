"""LLM provider 抽象介面與訊息資料結構。"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field


@dataclass
class ToolCall:
    name: str
    arguments: dict
    id: str = ""  # OpenAI 風格 provider 用


@dataclass
class Message:
    """跨 provider 的中性訊息結構。"""

    role: str  # "system" | "user" | "assistant" | "tool"
    text: str = ""
    tool_calls: list = field(default_factory=list)  # list[ToolCall]
    tool_name: str = ""  # role="tool" 時，對應呼叫的工具名


class LLMClient(ABC):
    @abstractmethod
    def is_available(self) -> bool:
        """是否可用（依賴已安裝、有 API key 或 endpoint 可達）。"""

    @abstractmethod
    def list_models(self) -> list:
        """列出支援文字生成的模型名稱。"""

    @abstractmethod
    def list_vision_models(self) -> list:
        """列出支援多模態（vision）的模型名稱（reviewer 使用）。"""

    def chat(self, messages, model=None, tools=None) -> Message:
        """送出對話、可選工具，回傳 assistant Message（可能含 tool_calls）。

        - messages: list[Message]，含 system / user / assistant / tool 各角色。
        - tools: list[dict]，每個工具的 {name, description, parameters} JSON schema。
        - model: 模型名稱字串。
        """
        raise NotImplementedError

    def vision_complete(self, system: str, user_text: str, images: list, model: str) -> str:
        """單回合多模態文字補全（reviewer 用）。

        - system: system instruction（可空字串）
        - user_text: 使用者要傳給模型的文字
        - images: list；每個元素為圖片檔案路徑（str）或 raw bytes
        - model: 模型名稱

        回傳：模型輸出的文字。
        """
        raise NotImplementedError
