from __future__ import annotations

from datetime import date

from .models import GanttPlan, GanttStyle, GanttTask, TaskVisualStyle


def sample_plan() -> GanttPlan:
    return GanttPlan(
        title="Website Refresh Program",
        summary="Illustrative plan used for offline smoke testing.",
        assumptions=[
            "Dates are only for validating slide rendering.",
            "Owners are placeholders.",
        ],
        tasks=[
            GanttTask(
                name="Discovery",
                start_date=date(2026, 3, 2),
                end_date=date(2026, 3, 13),
                owner="PMO",
                status="complete",
            ),
            GanttTask(
                name="Design",
                start_date=date(2026, 3, 9),
                end_date=date(2026, 3, 27),
                owner="Design",
                status="in_progress",
                dependencies=["Discovery"],
            ),
            GanttTask(
                name="Build",
                start_date=date(2026, 3, 23),
                end_date=date(2026, 4, 17),
                owner="Engineering",
                status="not_started",
                dependencies=["Design"],
            ),
            GanttTask(
                name="UAT",
                start_date=date(2026, 4, 20),
                end_date=date(2026, 5, 1),
                owner="Business",
                status="not_started",
                dependencies=["Build"],
            ),
        ],
    )


def updated_sample_plan() -> GanttPlan:
    return GanttPlan(
        title="Website Refresh Program",
        summary="Updated version showing schedule movement and a blocker.",
        assumptions=[
            "Discovery and design are complete.",
            "Environment access is delaying build validation.",
        ],
        tasks=[
            GanttTask(
                name="Discovery",
                start_date=date(2026, 3, 2),
                end_date=date(2026, 3, 13),
                owner="PMO",
                status="complete",
            ),
            GanttTask(
                name="Design",
                start_date=date(2026, 3, 9),
                end_date=date(2026, 3, 27),
                owner="Design",
                status="complete",
                dependencies=["Discovery"],
            ),
            GanttTask(
                name="Build",
                start_date=date(2026, 3, 30),
                end_date=date(2026, 4, 24),
                owner="Engineering",
                status="blocked",
                notes="Access for the new publishing environment is still pending.",
                dependencies=["Design"],
            ),
            GanttTask(
                name="UAT",
                start_date=date(2026, 4, 27),
                end_date=date(2026, 5, 8),
                owner="Business",
                status="not_started",
                dependencies=["Build"],
            ),
            GanttTask(
                name="Launch",
                start_date=date(2026, 5, 11),
                end_date=date(2026, 5, 15),
                owner="Operations",
                status="not_started",
                dependencies=["UAT"],
            ),
        ],
    )


def sample_style() -> GanttStyle:
    return GanttStyle(
        theme_name="corporate-blue",
        slide_background_color="#F8FAFC",
        title_color="#0F172A",
        summary_color="#334155",
        task_name_color="#0F172A",
        owner_color="#475569",
        header_fill_color="#DBEAFE",
        header_border_color="#93C5FD",
        header_text_color="#1E3A8A",
        row_fill_color="#FFFFFF",
        alternate_row_fill_color="#F8FAFC",
        grid_border_color="#CBD5E1",
        task_styles={
            "Build": TaskVisualStyle(
                fill_color="#8B5CF6",
                border_color="#6D28D9",
                text_color="#FFFFFF",
            ),
            "Launch": TaskVisualStyle(
                fill_color="#F59E0B",
                border_color="#D97706",
                text_color="#111827",
            ),
        },
    )


def creation_notes() -> str:
    return (
        "Discovery starts immediately and lasts two weeks. "
        "Design overlaps the second discovery week and runs through March. "
        "Build starts after design, then UAT follows."
    )


def update_notes() -> str:
    return (
        "Design is now complete. Build slipped one week and is blocked by environment access. "
        "UAT moves accordingly, and launch is a new final milestone."
    )
