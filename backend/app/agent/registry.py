"""Tool / ToolRegistry — 跨 provider 的工具註冊框架（無業務邏輯依賴）。"""

from dataclasses import dataclass
from typing import Any, Callable


@dataclass
class Tool:
    name: str
    description: str
    parameters: dict  # JSON schema (lowercase types)
    func: Callable

    def run(self, **kwargs) -> Any:
        return self.func(**kwargs)

    def schema(self) -> dict:
        return {
            "name": self.name,
            "description": self.description,
            "parameters": self.parameters,
        }


class ToolRegistry:
    def __init__(self):
        self._tools = {}

    def register(self, tool: Tool):
        self._tools[tool.name] = tool

    def get(self, name):
        return self._tools.get(name)

    def all(self):
        return list(self._tools.values())

    def schemas(self):
        return [t.schema() for t in self._tools.values()]

    def run(self, name, args):
        tool = self.get(name)
        if not tool:
            return {"error": f"未知工具: {name}"}
        try:
            return tool.run(**(args or {}))
        except TypeError as e:
            return {"error": f"參數錯誤: {e}"}
        except Exception as e:
            return {"error": str(e)}
