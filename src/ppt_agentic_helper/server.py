from __future__ import annotations

from functools import lru_cache

from fastmcp import FastMCP

from .models import GanttStyle
from .service import GanttAgentService

mcp = FastMCP("PowerPoint Gantt Helper")


@lru_cache(maxsize=1)
def get_service() -> GanttAgentService:
    return GanttAgentService()


@mcp.tool
def create_gantt_from_notes(
    notes: str,
    presentation_path: str,
    slide_title: str = "Project Gantt",
    style: dict | None = None,
    output_path: str | None = None,
) -> dict:
    """Create or replace a PowerPoint Gantt slide from free-form notes."""
    result = get_service().create_gantt_from_notes(
        notes=notes,
        presentation_path=presentation_path,
        slide_title=slide_title,
        style=GanttStyle.model_validate(style) if style else None,
        output_path=output_path,
    )
    return result.model_dump(mode="json")


@mcp.tool
def update_gantt_from_notes(
    notes: str,
    presentation_path: str,
    slide_title: str = "Project Gantt",
    style: dict | None = None,
    output_path: str | None = None,
) -> dict:
    """Update a previously generated Gantt slide using new project notes."""
    result = get_service().update_gantt_from_notes(
        notes=notes,
        presentation_path=presentation_path,
        slide_title=slide_title,
        style=GanttStyle.model_validate(style) if style else None,
        output_path=output_path,
    )
    return result.model_dump(mode="json")


@mcp.tool
def plan_notes(notes: str, title_hint: str | None = None) -> dict:
    """Convert notes into a structured Gantt plan without writing a PowerPoint file."""
    plan = get_service().plan_notes(notes=notes, title_hint=title_hint)
    return plan.model_dump(mode="json")


@mcp.tool
def inspect_gantt_slide(presentation_path: str, slide_title: str = "Project Gantt") -> dict:
    """Read the stored plan and style metadata from a previously generated Gantt slide."""
    managed_slide = get_service().inspect_slide(
        presentation_path=presentation_path,
        slide_title=slide_title,
    )
    return managed_slide.model_dump(mode="json")


@mcp.tool
def bootstrap_gantt_style(
    presentation_path: str,
    slide_title: str | None = None,
    slide_index: int | None = None,
    theme_name: str = "bootstrapped",
) -> dict:
    """Infer a reusable Gantt style from an existing PowerPoint slide."""
    style = get_service().bootstrap_style(
        presentation_path=presentation_path,
        slide_title=slide_title,
        slide_index=slide_index,
        theme_name=theme_name,
    )
    return style.model_dump(mode="json")


@mcp.resource("help://setup")
def setup_instructions() -> str:
    return """
Set these environment variables before running the server:
- OPENAI_API_KEY
- OPENAI_CHAT_MODEL (optional, default: gpt-4o-mini)
- OPENAI_BASE_URL (optional, default: https://api.openai.com/v1)
- OPENAI_CHAT_TIMEOUT_SECONDS (optional, default: 60)

The server uses the legacy /chat/completions endpoint for compatibility with older OpenAI-compatible stacks.
"""


def run() -> None:
    mcp.run()


if __name__ == "__main__":
    run()
