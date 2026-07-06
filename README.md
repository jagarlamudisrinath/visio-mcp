# visio-mcp

An MCP server that drives **Microsoft Visio desktop** via COM automation to draw diagrams for you: flowcharts, block diagrams, and Azure/AWS architecture diagrams using the official stencils. Claude (or any MCP client) drops real Visio masters, glues dynamic connectors, runs Visio's auto-layout, and **exports a PNG of the page so the AI can see the result and iterate**.

> **Runs on Windows only** (it automates the Visio desktop app through COM). The codebase is developed and unit-tested cross-platform, but the server must run on a Windows machine with Visio installed.

## Requirements

- Windows 10/11
- Microsoft Visio desktop (any recent version; must be installed and activated)
- [Python 3.10+](https://www.python.org/downloads/) and [uv](https://docs.astral.sh/uv/getting-started/installation/)
- Claude Code (or Claude Desktop) on the same machine

## Setup (on the Windows machine)

```powershell
git clone <this repo> C:\tools\visio_mcp   # or copy the folder
cd C:\tools\visio_mcp
uv sync
```

### Optional: install Azure / AWS stencils

For cloud architecture diagrams with official icons, download the stencil packs and unzip them into your **My Shapes** folder (usually `Documents\My Shapes`):

- **Azure**: search "Microsoft Azure stencils Visio download" (Microsoft publishes `.vssx` icon sets)
- **AWS**: search "AWS shapes Visio toolkit" (AWS publishes `.vssx` sets)

The server finds them automatically — `visio_status` reports the exact My Shapes path, and `open_stencil("azure")` fuzzy-matches files there.

### Smoke test

Verifies the full pipeline end-to-end (launches Visio, builds a small styled flowchart, auto-lays it out, exports a PNG, saves/reopens the .vsdx):

```powershell
uv run python scripts/smoke_test.py
```

### Register with Claude Code

```powershell
claude mcp add visio -- uv --directory C:\tools\visio_mcp run visio-mcp
```

Then ask Claude something like: *"Draw a 5-step login flowchart in Visio and show me the result."*

## Tools

| Tool | Purpose |
|---|---|
| `visio_status` | Version, open docs/stencils, My Shapes path; launches Visio if needed |
| `create_document` | New drawing, optionally from a template (`BASFLO_U.VSTX` = basic flowchart) |
| `open_document` / `save_document` | Open/save `.vsdx` files |
| `export_page_png` | Export the page as PNG **and return the image** for visual iteration |
| `open_stencil` | Open a stencil by built-in name, path, or fuzzy My Shapes match |
| `find_masters` | Search droppable shapes across open stencils |
| `drop_shape` / `drop_shapes` | Drop masters at (x, y) inches with optional text/size |
| `update_shape` / `style_shape` / `delete_shapes` | Edit text/position/size, colors/fonts, delete |
| `connect_shapes` | Glued dynamic connector with label, arrows, right-angle/straight/curved routing |
| `pages` | List/add/activate pages |
| `auto_layout` | Visio's automatic layout (flowchart top-bottom/left-right, tree, radial, circular) |
| `get_page_state` | Everything on a page: ids, masters, positions, text, connector endpoints |

**Coordinate system**: inches, origin at the page **bottom-left**, y grows upward, and drop coordinates are the shape's **center**. Rough placement is fine — finish with `auto_layout`.

## Architecture

```
Claude (stdio) → FastMCP server (asyncio; no COM imports)
                    ↓ submits callables
                 ComWorker — single STA thread owns ALL COM objects
                    ↓
                 VisioClient — the only module that touches win32com
                    ↓
                 Visio.Application
```

COM objects are apartment-threaded, so every COM call is funneled through one dedicated STA worker thread (`runtime.py`); tool bodies just `await` results. This also serializes concurrent tool calls safely.

## Development (any OS)

Unit tests run against a duck-typed fake of the Visio COM surface — no Windows needed:

```bash
uv sync
uv run pytest
```

## Troubleshooting

| Symptom | Fix |
|---|---|
| "Visio is busy — close any open dialog" | A modal dialog is open in the Visio window; close it and retry |
| "Visio desktop is not installed" | Install Visio desktop; Visio for the web cannot be automated |
| "Could not open a stencil matching …" | Download the stencil pack into `Documents\My Shapes`, then retry; use `visio_status` to confirm the path |
| First run shows Visio setup/license dialogs | Start Visio manually once, dismiss the dialogs, then retry |

## Roadmap

- Agent skills encoding Azure/AWS diagram conventions (zones/containers, brand colors, icon naming)
- Optional HTTP transport to drive a Windows Visio box from another machine
- Containers/groups via Visio's built-in container stencils
