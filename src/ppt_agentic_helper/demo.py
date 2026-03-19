from __future__ import annotations

from pathlib import Path

from .files import write_json_file
from .powerpoint import bootstrap_style_from_presentation, create_or_update_presentation, load_managed_slide
from .sample_data import creation_notes, sample_plan, sample_style, update_notes, updated_sample_plan


def run_capability_demo(output_dir: str, *, slide_title: str = "Capability Demo") -> dict[str, str]:
    base = Path(output_dir)
    base.mkdir(parents=True, exist_ok=True)

    source_pptx = base / "00-source-reference.pptx"
    created_pptx = base / "01-created-gantt.pptx"
    updated_pptx = base / "02-updated-gantt.pptx"
    source_style_json = base / "00-source-style.json"
    bootstrapped_style_json = base / "01-bootstrapped-style.json"
    created_metadata_json = base / "01-created-metadata.json"
    updated_metadata_json = base / "02-updated-metadata.json"
    creation_notes_path = base / "creation-notes.txt"
    update_notes_path = base / "update-notes.txt"

    source_style = sample_style()
    plan = sample_plan()
    revised_plan = updated_sample_plan()

    create_or_update_presentation(
        plan=plan,
        presentation_path=str(source_pptx),
        slide_title=slide_title,
        style=source_style,
        include_metadata=False,
    )
    write_json_file(str(source_style_json), source_style.model_dump(mode="json"))

    bootstrapped_style = bootstrap_style_from_presentation(
        presentation_path=str(source_pptx),
        slide_title=slide_title,
        theme_name="bootstrapped-demo",
    )
    write_json_file(str(bootstrapped_style_json), bootstrapped_style.model_dump(mode="json"))

    create_or_update_presentation(
        plan=plan,
        presentation_path=str(created_pptx),
        slide_title=slide_title,
        style=bootstrapped_style,
    )
    write_json_file(
        str(created_metadata_json),
        load_managed_slide(
            presentation_path=str(created_pptx),
            slide_title=slide_title,
        ).model_dump(mode="json"),
    )
    creation_notes_path.write_text(creation_notes(), encoding="utf-8")

    create_or_update_presentation(
        plan=revised_plan,
        presentation_path=str(created_pptx),
        slide_title=slide_title,
        style=bootstrapped_style,
        output_path=str(updated_pptx),
    )
    write_json_file(
        str(updated_metadata_json),
        load_managed_slide(
            presentation_path=str(updated_pptx),
            slide_title=slide_title,
        ).model_dump(mode="json"),
    )
    update_notes_path.write_text(update_notes(), encoding="utf-8")

    return {
        "output_dir": str(base.resolve()),
        "source_presentation": str(source_pptx.resolve()),
        "created_presentation": str(created_pptx.resolve()),
        "updated_presentation": str(updated_pptx.resolve()),
        "source_style_json": str(source_style_json.resolve()),
        "bootstrapped_style_json": str(bootstrapped_style_json.resolve()),
        "created_metadata_json": str(created_metadata_json.resolve()),
        "updated_metadata_json": str(updated_metadata_json.resolve()),
        "creation_notes": str(creation_notes_path.resolve()),
        "update_notes": str(update_notes_path.resolve()),
    }
