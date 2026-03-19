## PowerPoint Gantt Helper

`ppt-agentic-helper` is a Python package and MCP server for turning project notes into PowerPoint Gantt charts, or updating an existing generated Gantt chart with new notes.

It is designed around two constraints:

- It uses the legacy OpenAI-compatible `chat/completions` endpoint instead of the newer Responses API.
- It can run as an MCP server over stdio so teammates can plug it into MCP-capable clients.

## What It Does

- Converts free-form notes into a structured project plan.
- Creates a managed Gantt slide inside a `.pptx` file.
- Stores hidden slide metadata so the same slide can be updated later.
- Exposes the core workflows as MCP tools.

## Install

```bash
uv sync
```

## Configuration

Set these environment variables:

```bash
OPENAI_API_KEY=...
OPENAI_CHAT_MODEL=gpt-4o-mini
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_CHAT_TIMEOUT_SECONDS=60
```

`OPENAI_BASE_URL` can point at an internal OpenAI-compatible gateway if your workplace proxies the old chat endpoint.

## CLI Usage

Plan notes without writing PowerPoint output:

```bash
uv run ppt-agentic-helper plan --notes "Kickoff this week. Build starts next Monday. QA after build."
```

Create or replace a Gantt slide:

```bash
uv run ppt-agentic-helper create \
  --notes "Kickoff this week. Build starts next Monday. QA after build." \
  --presentation-path ./examples/project-plan.pptx \
  --slide-title "Program Timeline"
```

Update a previously generated slide:

```bash
uv run ppt-agentic-helper update \
  --notes "QA slipped by one week and infra is blocked waiting on access." \
  --presentation-path ./examples/project-plan.pptx \
  --slide-title "Program Timeline"
```

Run the MCP server over stdio:

```bash
uv run ppt-agentic-helper mcp
```

You can also use the dedicated entrypoint:

```bash
uv run ppt-agentic-mcp
```

## Testing Without PowerPoint

You do not need Microsoft PowerPoint installed to generate or update `.pptx` files with this project. The package uses `python-pptx`, which writes Office Open XML files directly rather than automating the PowerPoint desktop app.

Useful local checks:

Generate a sample deck without calling OpenAI:

```bash
uv run ppt-agentic-helper demo --presentation-path ./scratch/demo-gantt.pptx
```

Inspect the stored metadata from a generated slide:

```bash
uv run ppt-agentic-helper inspect \
  --presentation-path ./scratch/demo-gantt.pptx \
  --slide-title "Demo Gantt"
```

Generate a full end-to-end capability demo:

```bash
uv run ppt-agentic-helper demo-all --output-dir ./scratch/capability-demo
```

Bootstrap a reusable style from an existing slide:

```bash
uv run ppt-agentic-helper bootstrap-style \
  --presentation-path ./existing-deck.pptx \
  --slide-index 3 \
  --theme-name "existing-program-style" \
  --output-path ./scratch/existing-style.json
```

Run the automated tests:

```bash
uv run pytest
```

If you want to visually inspect the file and do not have PowerPoint, you can usually open the generated `.pptx` in LibreOffice Impress or upload it into Google Slides.

## Styling

You can lock a slide's appearance with a style JSON file. The selected style is stored in the slide metadata so later updates can preserve it.

Pass a style file during creation:

```bash
uv run ppt-agentic-helper create \
  --notes "Kickoff this week. Build starts next Monday. QA after build." \
  --presentation-path ./examples/project-plan.pptx \
  --slide-title "Program Timeline" \
  --style-path ./scratch/capability-demo/style-template.json
```

If you update a generated slide without `--style-path`, the tool reuses the style stored in the slide metadata. If you pass a new style file, it overrides the previous one.

If you already have a slide you like, `bootstrap-style` will infer a best-effort style JSON from that deck so future generated slides can match it more closely. It works best when the reference slide already has a Gantt-like layout with task rows on the left and timeline bars on the right.

Style file shape:

```json
{
  "theme_name": "corporate-blue",
  "slide_background_color": "#F8FAFC",
  "title_color": "#0F172A",
  "summary_color": "#334155",
  "header_fill_color": "#DBEAFE",
  "header_border_color": "#93C5FD",
  "header_text_color": "#1E3A8A",
  "row_fill_color": "#FFFFFF",
  "alternate_row_fill_color": "#F8FAFC",
  "grid_border_color": "#CBD5E1",
  "status_styles": {
    "not_started": { "fill_color": "#9CA3AF", "border_color": "#6B7280", "text_color": "#FFFFFF" },
    "in_progress": { "fill_color": "#3B82F6", "border_color": "#2563EB", "text_color": "#FFFFFF" },
    "blocked": { "fill_color": "#EF4444", "border_color": "#B91C1C", "text_color": "#FFFFFF" },
    "complete": { "fill_color": "#10B981", "border_color": "#059669", "text_color": "#FFFFFF" }
  },
  "task_styles": {
    "Build": { "fill_color": "#8B5CF6", "border_color": "#6D28D9", "text_color": "#FFFFFF" }
  }
}
```

## MCP Tools

The server exposes:

- `create_gantt_from_notes`
- `update_gantt_from_notes`
- `plan_notes`
- `inspect_gantt_slide`
- `bootstrap_gantt_style`

## Example MCP Client Config

Example Cursor/Claude Desktop style command:

```json
{
  "mcpServers": {
    "ppt-agentic-helper": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "C:/Users/robel/ppt-helper",
        "ppt-agentic-mcp"
      ],
      "env": {
        "OPENAI_API_KEY": "YOUR_KEY",
        "OPENAI_CHAT_MODEL": "gpt-4o-mini",
        "OPENAI_BASE_URL": "https://api.openai.com/v1"
      }
    }
  }
}
```

## Development

Run tests with:

```bash
uv run pytest
```
