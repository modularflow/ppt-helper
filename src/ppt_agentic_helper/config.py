from __future__ import annotations

import os
from dataclasses import dataclass


@dataclass(frozen=True)
class Settings:
    openai_api_key: str
    openai_model: str
    openai_base_url: str
    openai_timeout_seconds: float = 60.0

    @classmethod
    def from_env(cls) -> "Settings":
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise RuntimeError("OPENAI_API_KEY is required")

        base_url = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/")
        model = os.getenv("OPENAI_CHAT_MODEL", "gpt-4o-mini")
        timeout = float(os.getenv("OPENAI_CHAT_TIMEOUT_SECONDS", "60"))

        return cls(
            openai_api_key=api_key,
            openai_model=model,
            openai_base_url=base_url,
            openai_timeout_seconds=timeout,
        )
