from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .models import GanttStyle


def load_style_file(path: str | None) -> GanttStyle | None:
    if not path:
        return None
    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    return GanttStyle.model_validate(raw)


def write_json_file(path: str, payload: Any) -> str:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return str(output_path.resolve())
