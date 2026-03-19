from __future__ import annotations

import argparse
import json

from .demo import run_capability_demo
from .files import load_style_file, write_json_file
from .powerpoint import create_or_update_presentation
from .sample_data import sample_plan, sample_style
from .server import run as run_mcp_server
from .service import GanttAgentService


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="ppt-agentic-helper",
        description="Create and update PowerPoint Gantt charts from notes.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    mcp_parser = subparsers.add_parser("mcp", help="Run the MCP stdio server")
    mcp_parser.set_defaults(handler=_run_mcp)

    plan_parser = subparsers.add_parser("plan", help="Convert notes into a structured plan")
    plan_parser.add_argument("--notes", required=True, help="Free-form project notes")
    plan_parser.add_argument("--title-hint", help="Optional title hint for the plan")
    plan_parser.set_defaults(handler=_plan)

    create_parser = subparsers.add_parser("create", help="Create or replace a Gantt slide")
    _add_common_ppt_arguments(create_parser)
    create_parser.set_defaults(handler=_create)

    update_parser = subparsers.add_parser("update", help="Update an existing generated Gantt slide")
    _add_common_ppt_arguments(update_parser)
    update_parser.set_defaults(handler=_update)

    demo_parser = subparsers.add_parser(
        "demo",
        help="Create a sample PowerPoint Gantt slide without calling OpenAI",
    )
    demo_parser.add_argument(
        "--presentation-path",
        required=True,
        help="Where to write the demo PowerPoint file",
    )
    demo_parser.add_argument("--slide-title", default="Demo Gantt", help="Managed slide title")
    demo_parser.add_argument("--style-path", help="Optional JSON file describing the slide style")
    demo_parser.set_defaults(handler=_demo)

    demo_all_parser = subparsers.add_parser(
        "demo-all",
        help="Generate a full capability demo showing creation, parsing, and updating",
    )
    demo_all_parser.add_argument(
        "--output-dir",
        required=True,
        help="Directory where demo decks, notes, and metadata should be written",
    )
    demo_all_parser.add_argument("--slide-title", default="Capability Demo", help="Managed slide title")
    demo_all_parser.set_defaults(handler=_demo_all)

    inspect_parser = subparsers.add_parser(
        "inspect",
        help="Read a previously generated slide and print its stored plan and style metadata",
    )
    inspect_parser.add_argument("--presentation-path", required=True, help="PowerPoint file path")
    inspect_parser.add_argument("--slide-title", default="Project Gantt", help="Managed slide title")
    inspect_parser.set_defaults(handler=_inspect)

    bootstrap_parser = subparsers.add_parser(
        "bootstrap-style",
        help="Infer a reusable style JSON from an existing PowerPoint slide",
    )
    bootstrap_parser.add_argument("--presentation-path", required=True, help="PowerPoint file path")
    bootstrap_target = bootstrap_parser.add_mutually_exclusive_group(required=True)
    bootstrap_target.add_argument("--slide-title", help="Slide title to inspect")
    bootstrap_target.add_argument("--slide-index", type=int, help="1-based slide index to inspect")
    bootstrap_parser.add_argument("--theme-name", default="bootstrapped", help="Theme name to store in the style")
    bootstrap_parser.add_argument("--output-path", help="Optional JSON file path to write the inferred style")
    bootstrap_parser.set_defaults(handler=_bootstrap_style)

    args = parser.parse_args()
    args.handler(args)


def _add_common_ppt_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--notes", required=True, help="Free-form project notes")
    parser.add_argument("--presentation-path", required=True, help="Input PowerPoint file path")
    parser.add_argument("--slide-title", default="Project Gantt", help="Managed slide title")
    parser.add_argument("--style-path", help="Optional JSON file describing the slide style")
    parser.add_argument("--output-path", help="Optional output PowerPoint path")


def _run_mcp(_: argparse.Namespace) -> None:
    run_mcp_server()


def _plan(args: argparse.Namespace) -> None:
    service = GanttAgentService()
    plan = service.plan_notes(notes=args.notes, title_hint=args.title_hint)
    print(json.dumps(plan.model_dump(mode="json"), indent=2))


def _create(args: argparse.Namespace) -> None:
    service = GanttAgentService()
    result = service.create_gantt_from_notes(
        notes=args.notes,
        presentation_path=args.presentation_path,
        slide_title=args.slide_title,
        style=load_style_file(args.style_path),
        output_path=args.output_path,
    )
    print(json.dumps(result.model_dump(mode="json"), indent=2))


def _update(args: argparse.Namespace) -> None:
    service = GanttAgentService()
    result = service.update_gantt_from_notes(
        notes=args.notes,
        presentation_path=args.presentation_path,
        slide_title=args.slide_title,
        style=load_style_file(args.style_path),
        output_path=args.output_path,
    )
    print(json.dumps(result.model_dump(mode="json"), indent=2))


def _demo(args: argparse.Namespace) -> None:
    saved_path = create_or_update_presentation(
        plan=sample_plan(),
        presentation_path=args.presentation_path,
        slide_title=args.slide_title,
        style=load_style_file(args.style_path) or sample_style(),
    )
    print(json.dumps({"presentation_path": saved_path, "slide_title": args.slide_title}, indent=2))


def _inspect(args: argparse.Namespace) -> None:
    service = GanttAgentService()
    managed_slide = service.inspect_slide(
        presentation_path=args.presentation_path,
        slide_title=args.slide_title,
    )
    print(json.dumps(managed_slide.model_dump(mode="json"), indent=2))


def _demo_all(args: argparse.Namespace) -> None:
    result = run_capability_demo(args.output_dir, slide_title=args.slide_title)
    print(json.dumps(result, indent=2))


def _bootstrap_style(args: argparse.Namespace) -> None:
    service = GanttAgentService()
    style = service.bootstrap_style(
        presentation_path=args.presentation_path,
        slide_title=args.slide_title,
        slide_index=args.slide_index,
        theme_name=args.theme_name,
    )
    payload = style.model_dump(mode="json")
    if args.output_path:
        write_json_file(args.output_path, payload)
        payload = {**payload, "output_path": args.output_path}
    print(json.dumps(payload, indent=2))
