from __future__ import annotations

from typing import Any


class LLMClient:
    """Unified LLM client facade reserved for future narrative analysis.

    The current ADX report pipelines are deterministic SQL/Excel workflows, so
    these methods intentionally raise until a concrete provider is configured.
    """

    def chat(self, messages: list[dict[str, str]], **kwargs: Any) -> str:
        raise NotImplementedError("LLM provider is not configured.")

    def chat_json(self, messages: list[dict[str, str]], **kwargs: Any) -> dict[str, Any]:
        raise NotImplementedError("LLM provider is not configured.")

    def chat_with_tools(self, messages: list[dict[str, str]], tools: list[dict[str, Any]], **kwargs: Any) -> dict[str, Any]:
        raise NotImplementedError("LLM provider is not configured.")

    def chat_stream(self, messages: list[dict[str, str]], **kwargs: Any):
        raise NotImplementedError("LLM provider is not configured.")
