from __future__ import annotations

import json

from .models import GanttPlan, GanttStyle, ManagedGanttSlide
from .openai_chat import OpenAIChatClient

SYSTEM_PROMPT = """You convert messy project notes into a clean Gantt plan.

Return JSON only.
Use this schema:
{
  "title": "string",
  "summary": "string or null",
  "tasks": [
    {
      "name": "string",
      "start_date": "YYYY-MM-DD",
      "end_date": "YYYY-MM-DD",
      "owner": "string or null",
      "status": "not_started|in_progress|blocked|complete",
      "notes": "string or null",
      "dependencies": ["task name"]
    }
  ],
  "assumptions": ["string"]
}

Rules:
- Infer realistic dates when notes imply sequencing but not exact days.
- Prefer concise task names.
- Preserve important blockers and owners.
- Do not invent more than needed.
- Ensure each task has an end_date on or after start_date.
"""


def create_plan_from_notes(
    client: OpenAIChatClient,
    *,
    notes: str,
    title_hint: str | None = None,
) -> GanttPlan:
    prompt = f"""Create a project plan from these notes.

Title hint: {title_hint or "Use a concise project title"}

Notes:
{notes}
"""
    return client.complete_json(
        system_prompt=SYSTEM_PROMPT,
        user_prompt=prompt,
        response_model=GanttPlan,
    )


def update_plan_from_notes(
    client: OpenAIChatClient,
    *,
    current_plan: GanttPlan,
    notes: str,
) -> GanttPlan:
    prompt = f"""Update this existing plan using the new notes.

Current plan JSON:
{current_plan.model_dump_json(indent=2)}

New notes:
{notes}

Return the full revised plan, not a patch.
Keep tasks that still matter and adjust dates, status, owners, dependencies, and summary as needed.
"""
    return client.complete_json(
        system_prompt=SYSTEM_PROMPT,
        user_prompt=prompt,
        response_model=GanttPlan,
    )


def slide_to_metadata_text(slide: ManagedGanttSlide) -> str:
    return "PPT_AGENTIC_HELPER::" + json.dumps(slide.model_dump(mode="json"), separators=(",", ":"))


def metadata_text_to_slide(text: str) -> ManagedGanttSlide:
    prefix = "PPT_AGENTIC_HELPER::"
    if not text.startswith(prefix):
        raise RuntimeError("Slide metadata is missing the expected helper prefix")
    payload = text[len(prefix) :]
    parsed = json.loads(payload)
    if "plan" not in parsed:
        return ManagedGanttSlide(
            slide_title=parsed.get("title", "Project Gantt"),
            plan=GanttPlan.model_validate(parsed),
            style=GanttStyle(),
        )
    return ManagedGanttSlide.model_validate(parsed)


def plan_to_metadata_text(plan: GanttPlan) -> str:
    return slide_to_metadata_text(
        ManagedGanttSlide(
            slide_title=plan.title,
            plan=plan,
            style=GanttStyle(),
        )
    )


def metadata_text_to_plan(text: str) -> GanttPlan:
    return metadata_text_to_slide(text).plan
