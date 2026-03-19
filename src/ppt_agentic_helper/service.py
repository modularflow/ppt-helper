from __future__ import annotations

from .config import Settings
from .models import GanttPlan, GanttStyle, MCPResult, ManagedGanttSlide
from .openai_chat import OpenAIChatClient
from .planning import create_plan_from_notes, update_plan_from_notes
from .powerpoint import bootstrap_style_from_presentation, create_or_update_presentation, load_managed_slide


class GanttAgentService:
    def __init__(self, settings: Settings | None = None) -> None:
        self._settings = settings
        self._client: OpenAIChatClient | None = OpenAIChatClient(settings) if settings else None

    def _get_client(self) -> OpenAIChatClient:
        if self._client is None:
            self._settings = self._settings or Settings.from_env()
            self._client = OpenAIChatClient(self._settings)
        return self._client

    def create_gantt_from_notes(
        self,
        *,
        notes: str,
        presentation_path: str,
        slide_title: str,
        style: GanttStyle | None = None,
        output_path: str | None = None,
    ) -> MCPResult:
        plan = create_plan_from_notes(self._get_client(), notes=notes, title_hint=slide_title)
        effective_style = style or GanttStyle()
        saved_path = create_or_update_presentation(
            plan=plan,
            presentation_path=presentation_path,
            slide_title=slide_title,
            style=effective_style,
            output_path=output_path,
        )
        return MCPResult(
            message=f'Created or updated slide "{slide_title}" in {saved_path}',
            presentation_path=saved_path,
            slide_title=slide_title,
            plan=plan,
            style=effective_style,
        )

    def update_gantt_from_notes(
        self,
        *,
        notes: str,
        presentation_path: str,
        slide_title: str,
        style: GanttStyle | None = None,
        output_path: str | None = None,
    ) -> MCPResult:
        current_slide = load_managed_slide(
            presentation_path=presentation_path,
            slide_title=slide_title,
        )
        revised_plan = update_plan_from_notes(
            self._get_client(),
            current_plan=current_slide.plan,
            notes=notes,
        )
        effective_style = style or current_slide.style
        saved_path = create_or_update_presentation(
            plan=revised_plan,
            presentation_path=presentation_path,
            slide_title=slide_title,
            style=effective_style,
            output_path=output_path,
        )
        return MCPResult(
            message=f'Updated slide "{slide_title}" in {saved_path}',
            presentation_path=saved_path,
            slide_title=slide_title,
            plan=revised_plan,
            style=effective_style,
        )

    def plan_notes(self, *, notes: str, title_hint: str | None = None) -> GanttPlan:
        return create_plan_from_notes(self._get_client(), notes=notes, title_hint=title_hint)

    def inspect_slide(self, *, presentation_path: str, slide_title: str) -> ManagedGanttSlide:
        return load_managed_slide(presentation_path=presentation_path, slide_title=slide_title)

    def bootstrap_style(
        self,
        *,
        presentation_path: str,
        slide_title: str | None = None,
        slide_index: int | None = None,
        theme_name: str = "bootstrapped",
    ) -> GanttStyle:
        return bootstrap_style_from_presentation(
            presentation_path=presentation_path,
            slide_title=slide_title,
            slide_index=slide_index,
            theme_name=theme_name,
        )
