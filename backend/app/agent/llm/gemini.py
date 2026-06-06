"""Gemini 引擎 client（使用 Google GenAI 統一 SDK）。

注意：舊版 `google-generativeai` 已於 2025 年棄用；本檔使用新版
`google-genai`（匯入 `from google import genai`）。
"""

import json
import os
import time

try:
    from google import genai
    from google.genai import types as genai_types

    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False
    genai = None
    genai_types = None

from .base import LLMClient, Message, ToolCall


# ── 集中式呼叫冷卻(參考架構文件 2.11)──────────────────────────
# 所有模型呼叫(文字 + VLM)都經 _retry_call,冷卻集中在這裡,分兩層:
#   第1層 節流:每次呼叫前距上次至少 min_interval 秒(主動把 RPM 壓在限制下)。
#   第2層 退避:遇可重試錯誤(429/quota/rate_limit/5xx/timeout)sleep 後重試,倍增封頂。
import threading as _threading

_MIN_INTERVAL = float(os.environ.get("GEMINI_MIN_INTERVAL", "4.0"))
_last_call_t = [0.0]
_throttle_lock = _threading.Lock()

_RETRY_MAX_ATTEMPTS = 6
_RETRY_BACKOFF = (5.0, 10.0, 20.0, 40.0, 60.0)  # 封頂 60


def _throttle():
    with _throttle_lock:
        wait = _MIN_INTERVAL - (time.time() - _last_call_t[0])
        if wait > 0:
            time.sleep(wait)
        _last_call_t[0] = time.time()


def _is_retryable_error(exc):
    code = getattr(exc, "code", None) or getattr(exc, "status_code", None)
    if code in (429, 500, 502, 503, 504):
        return True
    msg = str(exc).lower()
    for kw in ("429", "quota", "rate", "resource_exhausted", "unavailable",
               "internal", "timeout", "deadline", "500", "503"):
        if kw in msg:
            return True
    return False


def _retry_call(fn):
    last_exc = None
    delay_idx = 0
    for attempt in range(_RETRY_MAX_ATTEMPTS):
        _throttle()  # 第1層:每次送出前主動節流
        try:
            return fn()
        except Exception as e:
            last_exc = e
            if attempt >= _RETRY_MAX_ATTEMPTS - 1 or not _is_retryable_error(e):
                raise
            time.sleep(_RETRY_BACKOFF[min(delay_idx, len(_RETRY_BACKOFF) - 1)])  # 第2層:退避
            delay_idx += 1
    if last_exc:
        raise last_exc


def _load_image(img):
    """接受路徑字串或 bytes，回 (bytes, mime_type)。"""
    if isinstance(img, (bytes, bytearray)):
        return bytes(img), "image/png"
    path = str(img).lower()
    with open(img, "rb") as f:
        data = f.read()
    if path.endswith(".jpg") or path.endswith(".jpeg"):
        return data, "image/jpeg"
    if path.endswith(".gif"):
        return data, "image/gif"
    if path.endswith(".webp"):
        return data, "image/webp"
    return data, "image/png"


class GeminiClient(LLMClient):
    def __init__(self, api_key=None):
        self.api_key = (
            api_key
            or os.environ.get("GEMINI_API_KEY")
            or os.environ.get("GOOGLE_API_KEY")
            or ""
        )
        self._client = None

    def _ensure_client(self):
        if not GEMINI_AVAILABLE or not self.api_key:
            return None
        if self._client is None:
            try:
                self._client = genai.Client(api_key=self.api_key)
            except Exception:
                self._client = None
        return self._client

    def is_available(self):
        return GEMINI_AVAILABLE and bool(self.api_key)

    def list_models(self):
        client = self._ensure_client()
        if client is None:
            return []
        try:
            names = []
            for m in client.models.list():
                actions = (
                    getattr(m, "supported_actions", None)
                    or getattr(m, "supported_generation_methods", None)
                    or []
                )
                if "generateContent" in actions:
                    raw = getattr(m, "name", "") or ""
                    names.append(raw.replace("models/", ""))
            return sorted(set(n for n in names if n))
        except Exception:
            return []

    def list_vision_models(self):
        # Gemini 1.5 / 2.x / 3.x 全系列為多模態
        return [m for m in self.list_models() if "gemini" in m.lower()]

    # ---- vision (P6 reviewer) ----

    def vision_complete(self, system, user_text, images, model):
        client = self._ensure_client()
        if client is None:
            raise RuntimeError("Gemini client 不可用：套件未安裝或 API key 未設定。")

        parts = [genai_types.Part(text=user_text or "")]
        for img in images or []:
            data, mime = _load_image(img)
            parts.append(genai_types.Part.from_bytes(data=data, mime_type=mime))

        contents = [genai_types.Content(role="user", parts=parts)]

        # Gemma 4 在 Gemini API 上要求 image Parts 在前 + 非空 system。對其他模型也安全。
        is_gemma = "gemma" in (model or "").lower()
        sys_text = system or ("你是圖像分析助理。" if is_gemma else None)
        config = genai_types.GenerateContentConfig(
            system_instruction=sys_text,
        )

        # 同上 — 圖片在前
        if parts and parts[0].text is not None:
            text_first = parts[0]
            image_parts = parts[1:]
            parts = image_parts + [text_first]
            contents = [genai_types.Content(role="user", parts=parts)]

        def _call():
            return client.models.generate_content(
                model=model or "gemini-2.5-flash",
                contents=contents,
                config=config,
            )
        response = _retry_call(_call)

        return getattr(response, "text", "") or ""

    # ---- chat ----

    def chat(self, messages, model=None, tools=None):
        client = self._ensure_client()
        if client is None:
            raise RuntimeError("Gemini client 不可用：套件未安裝或 API key 未設定。")

        contents, system_instruction = self._convert_messages(messages)

        gemini_tools = None
        if tools:
            decls = [
                genai_types.FunctionDeclaration(
                    name=t["name"],
                    description=t.get("description", ""),
                    parameters_json_schema=t.get("parameters", {"type": "object"}),
                )
                for t in tools
            ]
            gemini_tools = [genai_types.Tool(function_declarations=decls)]

        config = genai_types.GenerateContentConfig(
            tools=gemini_tools,
            system_instruction=system_instruction or None,
        )

        def _call():
            return client.models.generate_content(
                model=model or "gemini-2.5-flash",
                contents=contents,
                config=config,
            )
        response = _retry_call(_call)

        return self._parse_response(response)

    @staticmethod
    def _convert_messages(messages):
        """轉成 Gemini Content 列表 + system_instruction 字串。"""
        contents = []
        system_parts = []

        for msg in messages:
            if msg.role == "system":
                if msg.text:
                    system_parts.append(msg.text)
                continue

            if msg.role == "user":
                contents.append(
                    genai_types.Content(
                        role="user",
                        parts=[genai_types.Part(text=msg.text or "")],
                    )
                )
            elif msg.role == "assistant":
                parts = []
                if msg.text:
                    parts.append(genai_types.Part(text=msg.text))
                for tc in msg.tool_calls:
                    parts.append(
                        genai_types.Part(
                            function_call=genai_types.FunctionCall(
                                name=tc.name,
                                args=tc.arguments or {},
                            )
                        )
                    )
                if parts:
                    contents.append(genai_types.Content(role="model", parts=parts))
            elif msg.role == "tool":
                # 解析工具回傳結果為 dict（function_response 需要 dict）
                response_obj = GeminiClient._coerce_response(msg.text)
                contents.append(
                    genai_types.Content(
                        role="tool",
                        parts=[
                            genai_types.Part.from_function_response(
                                name=msg.tool_name,
                                response=response_obj,
                            )
                        ],
                    )
                )

        return contents, "\n\n".join(system_parts)

    @staticmethod
    def _coerce_response(text):
        if not text:
            return {"result": ""}
        try:
            value = json.loads(text)
            if isinstance(value, dict):
                return value
            return {"result": value}
        except (json.JSONDecodeError, TypeError):
            return {"result": text}

    @staticmethod
    def _parse_response(response):
        out = Message(role="assistant")

        # 便利存取：response.function_calls
        function_calls = getattr(response, "function_calls", None) or []
        for fc in function_calls:
            args = dict(fc.args) if getattr(fc, "args", None) else {}
            out.tool_calls.append(ToolCall(name=fc.name, arguments=args))

        # 文字部分
        text = getattr(response, "text", None)
        if text:
            out.text = text
            return out

        # 退而求其次：自行掃 candidates parts
        candidates = getattr(response, "candidates", None) or []
        if candidates:
            content = getattr(candidates[0], "content", None)
            parts = getattr(content, "parts", None) or []
            collected_text = []
            for part in parts:
                if getattr(part, "function_call", None) and not function_calls:
                    fc = part.function_call
                    args = dict(fc.args) if getattr(fc, "args", None) else {}
                    out.tool_calls.append(ToolCall(name=fc.name, arguments=args))
                t = getattr(part, "text", None)
                if t:
                    collected_text.append(t)
            if collected_text:
                out.text = "".join(collected_text)

        return out
