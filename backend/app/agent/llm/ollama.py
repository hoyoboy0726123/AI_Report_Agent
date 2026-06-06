"""Ollama 本地引擎 client。

使用 stdlib urllib 直打 REST API，避免額外依賴。
P3 若需要更穩的 tool calling 再考慮加入官方 ollama Python lib。
"""

import base64
import json
import os
import urllib.request

from .base import LLMClient, Message, ToolCall


class OllamaClient(LLMClient):
    VISION_KEYWORDS = (
        "llava",
        "bakllava",
        "moondream",
        "minicpm-v",
        "qwen2-vl",
        "qwen2.5-vl",
        "llama3.2-vision",
        "llama4-vision",
        "vision",
    )

    def __init__(self, endpoint=None, timeout=5, num_ctx=None):
        self.endpoint = (
            endpoint
            or os.environ.get("OLLAMA_ENDPOINT")
            or "http://localhost:11434"
        ).rstrip("/")
        self.timeout = timeout
        # Ollama 預設 context 只有 2048/4096,塞不下 48 個工具 schema 會被靜默截斷
        # 導致模型回空、完全不呼叫工具。固定拉高,且每次請求都要帶(否則會被打回預設)。
        try:
            self.num_ctx = int(num_ctx) if num_ctx else int(
                os.environ.get("OLLAMA_NUM_CTX", "32768"))
        except (TypeError, ValueError):
            self.num_ctx = 32768

    def _get(self, path):
        with urllib.request.urlopen(f"{self.endpoint}{path}", timeout=self.timeout) as r:
            return json.loads(r.read().decode("utf-8"))

    def _post_json(self, path, payload, timeout=None):
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            f"{self.endpoint}{path}",
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=timeout or 120) as r:
            return json.loads(r.read().decode("utf-8"))

    def is_available(self):
        try:
            self._get("/api/tags")
            return True
        except Exception:
            return False

    def list_models(self):
        try:
            data = self._get("/api/tags")
        except Exception:
            return []
        return sorted(
            m.get("name", "")
            for m in data.get("models", [])
            if m.get("name")
        )

    def list_vision_models(self):
        return [
            m
            for m in self.list_models()
            if any(k in m.lower() for k in self.VISION_KEYWORDS)
        ]

    # ---- chat ----

    def chat(self, messages, model=None, tools=None):
        if not model:
            raise RuntimeError("未指定 Ollama 模型。")

        payload = {
            "model": model,
            "stream": False,
            "messages": [self._msg_to_ollama(m) for m in messages],
            "options": {"num_ctx": self.num_ctx},
        }
        if tools:
            payload["tools"] = [
                {
                    "type": "function",
                    "function": {
                        "name": t["name"],
                        "description": t.get("description", ""),
                        "parameters": t.get("parameters", {"type": "object"}),
                    },
                }
                for t in tools
            ]

        try:
            body = self._post_json("/api/chat", payload, timeout=180)
        except Exception as e:
            raise RuntimeError(f"Ollama 連線失敗: {e}")

        msg = body.get("message", {}) or {}
        out = Message(role="assistant", text=msg.get("content", "") or "")
        for tc in msg.get("tool_calls", []) or []:
            fn = tc.get("function", {}) or {}
            args = fn.get("arguments", {})
            if isinstance(args, str):
                # 部分模型回傳 JSON 字串而非 dict
                try:
                    args = json.loads(args)
                except (json.JSONDecodeError, TypeError):
                    args = {"_raw": args}
            out.tool_calls.append(
                ToolCall(name=fn.get("name", ""), arguments=args or {})
            )
        return out

    # ---- vision (P6 reviewer) ----

    def vision_complete(self, system, user_text, images, model):
        if not model:
            raise RuntimeError("未指定 Ollama 模型。")

        encoded = []
        for img in images or []:
            if isinstance(img, (bytes, bytearray)):
                data = bytes(img)
            else:
                with open(img, "rb") as f:
                    data = f.read()
            encoded.append(base64.b64encode(data).decode("ascii"))

        messages = []
        if system:
            messages.append({"role": "system", "content": system})
        messages.append(
            {
                "role": "user",
                "content": user_text or "",
                "images": encoded,
            }
        )

        payload = {"model": model, "stream": False, "messages": messages,
                   "options": {"num_ctx": self.num_ctx}}

        try:
            body = self._post_json("/api/chat", payload, timeout=300)
        except Exception as e:
            raise RuntimeError(f"Ollama 連線失敗: {e}")

        msg = body.get("message", {}) or {}
        return msg.get("content", "") or ""

    @staticmethod
    def _msg_to_ollama(msg):
        if msg.role == "user":
            return {"role": "user", "content": msg.text or ""}
        if msg.role == "system":
            return {"role": "system", "content": msg.text or ""}
        if msg.role == "assistant":
            d = {"role": "assistant", "content": msg.text or ""}
            if msg.tool_calls:
                d["tool_calls"] = [
                    {"function": {"name": tc.name, "arguments": tc.arguments or {}}}
                    for tc in msg.tool_calls
                ]
            return d
        if msg.role == "tool":
            d = {"role": "tool", "content": msg.text or ""}
            if msg.tool_name:
                d["name"] = msg.tool_name
            return d
        return {"role": "user", "content": msg.text or ""}
