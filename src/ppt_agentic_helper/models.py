from __future__ import annotations

from datetime import date
from typing import Literal

from pydantic import BaseModel, Field, model_validator


TaskStatus = Literal["not_started", "in_progress", "blocked", "complete"]


class TaskVisualStyle(BaseModel):
    fill_color: str | None = None
    border_color: str | None = None
    text_color: str | None = None


class GanttStyle(BaseModel):
    theme_name: str = "default"
    slide_background_color: str = "#FFFFFF"
    title_color: str = "#1F2937"
    summary_color: str = "#4B5563"
    task_name_color: str = "#111827"
    owner_color: str = "#4B5563"
    header_fill_color: str = "#F3F4F6"
    header_border_color: str = "#D1D5DB"
    header_text_color: str = "#1F2937"
    row_fill_color: str = "#FFFFFF"
    alternate_row_fill_color: str = "#F9FAFB"
    grid_border_color: str = "#E5E7EB"
    status_styles: dict[TaskStatus, TaskVisualStyle] = Field(
        default_factory=lambda: {
            "not_started": TaskVisualStyle(
                fill_color="#9CA3AF", border_color="#6B7280", text_color="#FFFFFF"
            ),
            "in_progress": TaskVisualStyle(
                fill_color="#3B82F6", border_color="#2563EB", text_color="#FFFFFF"
            ),
            "blocked": TaskVisualStyle(
                fill_color="#EF4444", border_color="#B91C1C", text_color="#FFFFFF"
            ),
            "complete": TaskVisualStyle(
                fill_color="#10B981", border_color="#059669", text_color="#FFFFFF"
            ),
        }
    )
    task_styles: dict[str, TaskVisualStyle] = Field(default_factory=dict)


class GanttTask(BaseModel):
    name: str = Field(..., min_length=1)
    start_date: date
    end_date: date
    owner: str | None = None
    status: TaskStatus = "not_started"
    notes: str | None = None
    dependencies: list[str] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_dates(self) -> "GanttTask":
        if self.end_date < self.start_date:
            raise ValueError("end_date must be on or after start_date")
        return self


class GanttPlan(BaseModel):
    title: str = Field(..., min_length=1)
    summary: str | None = None
    tasks: list[GanttTask] = Field(default_factory=list)
    assumptions: list[str] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_tasks(self) -> "GanttPlan":
        if not self.tasks:
            raise ValueError("at least one task is required")
        return self

    @property
    def start_date(self) -> date:
        return min(task.start_date for task in self.tasks)

    @property
    def end_date(self) -> date:
        return max(task.end_date for task in self.tasks)


class ManagedGanttSlide(BaseModel):
    slide_title: str
    plan: GanttPlan
    style: GanttStyle = Field(default_factory=GanttStyle)


class MCPResult(BaseModel):
    message: str
    presentation_path: str | None = None
    slide_title: str | None = None
    plan: GanttPlan | None = None
    style: GanttStyle | None = None
