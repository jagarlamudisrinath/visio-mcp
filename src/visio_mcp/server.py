"""visio-mcp: MCP server that drives Microsoft Visio desktop via COM.

Tool bodies never touch COM directly: every call is submitted to the
ComWorker's single STA thread, which owns the VisioClient and all COM
objects (see runtime.py for why).
"""

from __future__ import annotations

import functools
import os
import tempfile
import time
from typing import Literal, Optional

from mcp.server.fastmcp import FastMCP, Image
from mcp.server.fastmcp.exceptions import ToolError

from .errors import VisioMcpError
from .models import DropSpec
from .runtime import ComWorker
from .visio_client import VisioClient

mcp = FastMCP(
    "visio",
    instructions=(
        "Draw diagrams in a live Microsoft Visio desktop instance (Windows). "
        "Typical workflow: visio_status -> create_document (e.g. template "
        "'BASFLO_U.VSTX' for flowcharts) -> open_stencil / find_masters to "
        "discover shapes -> drop_shapes -> connect_shapes -> auto_layout -> "
        "export_page_png to SEE the result and iterate -> save_document. "
        "Coordinate system: units are INCHES, origin at the page BOTTOM-LEFT, "
        "y grows UPWARD, and the drop point (x, y) is the shape's CENTER. "
        "Rough placement is fine when you finish with auto_layout. For Azure/"
        "AWS architecture diagrams, open the official stencil packs installed "
        "in the My Shapes folder (see visio_status.my_shapes_path) and search "
        "them with find_masters."
    ),
)

_worker = ComWorker()
_client = VisioClient()
_export_counter = 0


def _tool_errors(fn):
    """Convert VisioMcpError (and unexpected errors) into clean ToolErrors."""

    @functools.wraps(fn)
    async def wrapper(*args, **kwargs):
        try:
            return await fn(*args, **kwargs)
        except (VisioMcpError, ValueError) as exc:
            raise ToolError(str(exc)) from exc

    return wrapper


@mcp.tool()
@_tool_errors
async def visio_status() -> dict:
    """Report Visio state: version, open documents and stencils, active page,
    and the My Shapes folder path (where Azure/AWS stencil packs belong).
    Launches Visio if it is not already running. Call this first."""
    return await _worker.run(_client.status)


@mcp.tool()
@_tool_errors
async def create_document(
    template: Optional[str] = None,
    measurement: Literal["us", "metric"] = "us",
) -> dict:
    """Create a new drawing document in Visio.

    Args:
        template: Visio template name or absolute path. Examples:
            'BASFLO_U.VSTX' (Basic Flowchart, US units) or bare 'BASFLO'
            (units suffix added from `measurement`). Omit for a blank drawing.
        measurement: 'us' (inches) or 'metric' — used when template has no
            _U/_M suffix.
    """
    return await _worker.run(_client.create_document, template, measurement)


@mcp.tool()
@_tool_errors
async def open_document(path: str) -> dict:
    """Open an existing Visio file (.vsdx/.vsd) for viewing or editing."""
    return await _worker.run(_client.open_document, path)


@mcp.tool()
@_tool_errors
async def save_document(path: Optional[str] = None) -> dict:
    """Save the active drawing. Pass an absolute path ending in .vsdx the
    first time; afterwards you can omit it to save in place."""
    return await _worker.run(_client.save_document, path)


@mcp.tool()
@_tool_errors
async def export_page_png(
    path: Optional[str] = None,
    page: Optional[str] = None,
) -> list:
    """Export a page as PNG and RETURN THE IMAGE so you can visually inspect
    the diagram. Use this after dropping/connecting/layout to verify the
    result and iterate.

    Args:
        path: Output .png path; omit for an auto-named temp file.
        page: Page name; defaults to the active page.
    """
    global _export_counter
    if path is None:
        _export_counter += 1
        path = os.path.join(
            tempfile.gettempdir(),
            f"visio_mcp_{os.getpid()}_{_export_counter}_{int(time.time())}.png",
        )
    result = await _worker.run(_client.export_page_png, path, page)
    return [f"Exported page {result['page']!r} to {result['path']}", Image(path=result["path"])]


@mcp.tool()
@_tool_errors
async def open_stencil(name_or_path: str) -> dict:
    """Open a stencil (shape library) so its masters can be dropped.

    Accepts a built-in name ('BASFLO_U.VSSX' basic flowchart, 'PERIPH_U.VSSX'
    peripherals...), an absolute path, or a fuzzy name matched against files
    in the My Shapes folder (where downloaded Azure/AWS stencil packs live).
    """
    return await _worker.run(_client.open_stencil, name_or_path)


@mcp.tool()
@_tool_errors
async def find_masters(query: Optional[str] = None, stencil: Optional[str] = None) -> dict:
    """Search masters (droppable shapes) in the open stencils by name
    substring. Omit `query` to list everything (capped at 100)."""
    return await _worker.run(_client.find_masters, query, stencil)


@mcp.tool()
@_tool_errors
async def drop_shape(
    master: str,
    x: float,
    y: float,
    stencil: Optional[str] = None,
    text: Optional[str] = None,
    width_in: Optional[float] = None,
    height_in: Optional[float] = None,
    page: Optional[str] = None,
) -> dict:
    """Drop one master onto the page. (x, y) is the shape CENTER in inches
    from the bottom-left. Returns the shape_id used by other tools."""
    spec = DropSpec(
        master=master, x=x, y=y, stencil=stencil, text=text,
        width_in=width_in, height_in=height_in,
    ).model_dump()
    result = await _worker.run(_client.drop_shapes, [spec], page)
    return result["shapes"][0]


@mcp.tool()
@_tool_errors
async def drop_shapes(shapes: list[DropSpec], page: Optional[str] = None) -> dict:
    """Drop many masters in one call (preferred when building a diagram).
    Executed in a single undo scope. Returns a shape_id per item, in order."""
    specs = [s.model_dump() for s in shapes]
    return await _worker.run(_client.drop_shapes, specs, page)


@mcp.tool()
@_tool_errors
async def update_shape(
    shape_id: int,
    text: Optional[str] = None,
    x: Optional[float] = None,
    y: Optional[float] = None,
    width_in: Optional[float] = None,
    height_in: Optional[float] = None,
    page: Optional[str] = None,
) -> dict:
    """Change a shape's text, position (center, inches), or size. Only the
    provided fields are changed."""
    return await _worker.run(
        _client.update_shape, shape_id, page, text, x, y, width_in, height_in
    )


@mcp.tool()
@_tool_errors
async def style_shape(
    shape_id: int,
    fill_color: Optional[str] = None,
    line_color: Optional[str] = None,
    text_color: Optional[str] = None,
    line_weight_pt: Optional[float] = None,
    font_size_pt: Optional[float] = None,
    bold: Optional[bool] = None,
    page: Optional[str] = None,
) -> dict:
    """Style a shape. Colors are hex like '#0078D4' (Azure blue) or '#FF9900'
    (AWS orange). Only the provided fields are changed."""
    return await _worker.run(
        _client.style_shape, shape_id, page, fill_color, line_color,
        text_color, line_weight_pt, font_size_pt, bold,
    )


@mcp.tool()
@_tool_errors
async def delete_shapes(shape_ids: list[int], page: Optional[str] = None) -> dict:
    """Delete shapes (or connectors) by id."""
    return await _worker.run(_client.delete_shapes, shape_ids, page)


@mcp.tool()
@_tool_errors
async def connect_shapes(
    from_id: int,
    to_id: int,
    label: Optional[str] = None,
    route: Literal["right_angle", "straight", "curved"] = "right_angle",
    end_arrow: bool = True,
    begin_arrow: bool = False,
    page: Optional[str] = None,
) -> dict:
    """Connect two shapes with a dynamic connector glued to both (Visio
    auto-picks the best sides and re-routes when shapes move).

    Args:
        label: Optional text on the connector (e.g. 'Yes' / 'No').
        route: 'right_angle' (default), 'straight', or 'curved'.
    """
    return await _worker.run(
        _client.connect_shapes, from_id, to_id, label, route,
        end_arrow, begin_arrow, page,
    )


@mcp.tool()
@_tool_errors
async def pages(
    action: Literal["list", "add", "activate"],
    name: Optional[str] = None,
) -> dict:
    """List pages, add a page (optionally named), or activate a page by name."""
    return await _worker.run(_client.pages, action, name)


@mcp.tool()
@_tool_errors
async def auto_layout(
    style: Literal[
        "flowchart_tb", "flowchart_lr", "tree_tb", "tree_lr", "radial", "circular"
    ] = "flowchart_tb",
    spacing_in: float = 0.75,
    resize_page: bool = True,
    page: Optional[str] = None,
) -> dict:
    """Run Visio's automatic layout on the page — arranges shapes and
    re-routes connectors. Great after dropping shapes at rough positions.
    'flowchart_tb' = top-to-bottom flowchart, '_lr' = left-to-right."""
    return await _worker.run(_client.auto_layout, style, spacing_in, resize_page, page)


@mcp.tool()
@_tool_errors
async def get_page_state(page: Optional[str] = None) -> dict:
    """List everything on a page: shape ids, master names, positions (inches,
    center), sizes, text, and connector endpoints (from_id/to_id). Call this
    to (re)discover shape ids, e.g. after open_document."""
    return await _worker.run(_client.get_page_state, page)


def main() -> None:
    try:
        mcp.run()  # stdio transport
    finally:
        _worker.shutdown(release=_client.release)


if __name__ == "__main__":
    main()
