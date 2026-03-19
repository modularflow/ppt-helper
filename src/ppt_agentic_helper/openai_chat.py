from __future__ import annotations

import json
from typing import TypeVar

import requests
from pydantic import BaseModel, ValidationError

from .config import Settings

T = TypeVar("T", bound=BaseModel)


class OpenAIChatClient:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings

    def complete_json(self, *, system_prompt: str, user_prompt: str, response_model: type[T]) -> T:
        payload = {
            "model": self._settings.openai_model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": 0.2,
            "response_format": {"type": "json_object"},
        }
        response = requests.post(
            f"{self._settings.openai_base_url}/chat/completions",
            headers={
                "Authorization": f"Bearer {self._settings.openai_api_key}",
                "Content-Type": "application/json",
            },
            json=payload,
            timeout=self._settings.openai_timeout_seconds,
        )
        response.raise_for_status()

        data = response.json()
        content = data["choices"][0]["message"]["content"]
        raw_json = extract_json_object(content)

        try:
            return response_model.model_validate_json(raw_json)
        except ValidationError as exc:
            raise RuntimeError(
                f"OpenAI response did not match expected schema: {exc}\nRaw content: {content}"
            ) from exc


def extract_json_object(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        lines = [line for line in text.splitlines() if not line.strip().startswith("```")]
        text = "\n".join(lines).strip()

    try:
        json.loads(text)
        return text
    except json.JSONDecodeError:
        pass

    start = text.find("{")
    if start == -1:
        raise RuntimeError(f"Expected JSON object in model output, got: {text}")

    depth = 0
    in_string = False
    escape = False
    for index in range(start, len(text)):
        char = text[index]

        if escape:
            escape = False
            continue

        if char == "\\":
            escape = True
            continue

        if char == '"':
            in_string = not in_string
            continue

        if in_string:
            continue

        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[start : index + 1]

    raise RuntimeError(f"Could not extract JSON object from model output: {text}")
