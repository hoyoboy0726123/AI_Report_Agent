"""LLM provider 工廠。"""

from .base import LLMClient
from .gemini import GeminiClient, GEMINI_AVAILABLE
from .ollama import OllamaClient

PROVIDERS = ("Gemini", "Ollama")


def get_client(provider, **kwargs):
    if provider == "Gemini":
        return GeminiClient(**kwargs)
    if provider == "Ollama":
        return OllamaClient(**kwargs)
    raise ValueError(f"Unknown provider: {provider}")


__all__ = [
    "LLMClient",
    "GeminiClient",
    "OllamaClient",
    "GEMINI_AVAILABLE",
    "PROVIDERS",
    "get_client",
]
