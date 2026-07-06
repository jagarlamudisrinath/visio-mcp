"""visio-mcp: MCP server that drives Microsoft Visio desktop via COM.

Tool bodies never touch COM directly: every call is submitted to the
ComWorker's single STA thread, which owns the VisioClient and all COM
objects (see runtime.py for why).
"""

from __future__ import annotations

import asyncio
import functools
import os
import tempfile
import time
import urllib.parse
import urllib.request
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
        "Rough placement is fine when you finish with auto_layout. For wide "
        "architecture diagrams, call set_page_size FIRST (default page is "
        "8.5x11 and the PNG export crops to it). Group shapes into zones "
        "(VNet/VPC/subnet/trust boundary) with add_container — innermost "
        "containers first. Use drop_text for titles and legends, and dashed "
        "connectors (line_pattern) for control-plane/auth flows. For Azure/"
        "AWS icons, Visio ships built-in cloud stencils — visio_status lists "
        "them in builtin_cloud_stencils and they open by bare filename (e.g. "
        "open_stencil('AZURESTORAGE_U.VSSX')); extra vendor packs can go in "
        "the My Shapes folder. Search open stencils with find_masters. For an "
        "icon Visio has no master for (e.g. the Azure 'Subnet' glyph), save a "
        "labeled image (subnet.svg) in the local icon folder (list_local_icons "
        "shows the path) so drop_shape('subnet', ...) places it, or insert one "
        "on the fly with import_image (a local file path or a direct image "
        "URL)."
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


_IMAGE_EXTS = {".svg", ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".emf", ".wmf"}
_CTYPE_EXT = {
    "image/svg+xml": ".svg",
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
}


def _download_image_to_temp(url: str) -> str:
    """Download a direct image URL to a temp file for Visio to Import.

    Guards against the common failure where an icon *gallery* URL returns an
    HTML single-page app instead of the raw image bytes.
    """
    req = urllib.request.Request(url, headers={"User-Agent": "visio-mcp"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            ctype = (resp.headers.get("Content-Type") or "").split(";")[0].strip().lower()
            data = resp.read()
    except Exception as exc:
        raise VisioMcpError(f"Failed to download image from {url!r}: {exc}") from exc
    head = data[:64].lstrip().lower()
    if ctype == "text/html" or head.startswith(b"<!doctype html") or head.startswith(b"<html"):
        raise VisioMcpError(
            f"{url!r} returned an HTML page, not an image. Icon galleries like "
            "az-icons.com are single-page apps whose /icon URLs serve HTML — "
            "use the site's Download button and pass the saved file path, or a "
            "direct raw image URL (one that ends in .svg/.png)."
        )
    ext = os.path.splitext(urllib.parse.urlparse(url).path)[1].lower()
    if ext not in _IMAGE_EXTS:
        ext = _CTYPE_EXT.get(ctype, ".img")
    fd, tmp = tempfile.mkstemp(prefix="visio_mcp_icon_", suffix=ext)
    os.close(fd)
    with open(tmp, "wb") as fh:
        fh.write(data)
    return tmp


@mcp.tool()
@_tool_errors
async def visio_status() -> dict:
    """Report Visio state: version, open documents and stencils, active page,
    the My Shapes folder path, and the built-in stencil library — including
    builtin_cloud_stencils, the Azure/AWS stencil files Visio ships with
    (openable by bare filename, no download needed). Launches Visio if it is
    not already running. Call this first."""
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
    substring. Omit `query` to list everything (capped at 100). Matching
    custom icons from the local icon folder are included too, marked with
    stencil '(local icon)'."""
    return await _worker.run(_client.find_masters, query, stencil)


@mcp.tool()
@_tool_errors
async def list_local_icons() -> dict:
    """List the local custom-icon folder and the labeled image files in it.

    Users manually download icons (e.g. from az-icons.com) and save them into
    this folder; each file's name without its extension is its label. Labels
    are discoverable via find_masters and can be placed with
    drop_shape('<label>', x, y) when Visio has no built-in master. The folder
    defaults to ~/.visio-mcp/icons and moves with the VISIO_MCP_ICONS_DIR
    environment variable."""
    return await _worker.run(_client.list_local_icons)


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
async def import_image(
    source: str,
    x: float,
    y: float,
    width_in: Optional[float] = None,
    height_in: Optional[float] = None,
    page: Optional[str] = None,
) -> dict:
    """Insert an external image (SVG, PNG, JPG, EMF, ...) onto the page as a
    shape — for custom icons that have NO Visio master (e.g. the newer Azure
    'Subnet' glyph) or vendor logos.

    Args:
        source: a LOCAL file path OR a direct http(s) IMAGE URL. A URL must
            point at the raw image bytes; icon *web pages* (single-page-app
            galleries such as az-icons.com) return HTML and are rejected with
            guidance — download the file in the browser and pass its path.
        x, y: shape CENTER in inches from the bottom-left.
        width_in / height_in: size in inches. Give just one to scale the other
            automatically and preserve the image's aspect ratio; omit both to
            keep the image's native size.

    Prefer built-in masters (drop_shape) when Visio has the icon; use this for
    the gaps. Returns the shape_id.
    """
    path = source
    if source.strip().lower().startswith(("http://", "https://")):
        loop = asyncio.get_running_loop()
        path = await loop.run_in_executor(None, _download_image_to_temp, source.strip())
    return await _worker.run(
        _client.import_image, path, x, y, width_in, height_in, page
    )


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
    line_pattern: Optional[Literal["solid", "dashed", "dotted", "dash_dot"]] = None,
    page: Optional[str] = None,
) -> dict:
    """Style a shape, connector, or container. Colors are hex like '#0078D4'
    (Azure blue) or '#FF9900' (AWS orange). Only the provided fields are
    changed. Works on containers too: theme-guarded cells are force-
    overridden, so recoloring a VNet/subscription boundary is supported."""
    return await _worker.run(
        _client.style_shape, shape_id, page, fill_color, line_color,
        text_color, line_weight_pt, font_size_pt, bold, line_pattern,
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
    line_pattern: Literal["solid", "dashed", "dotted", "dash_dot"] = "solid",
    line_weight_pt: Optional[float] = None,
    line_color: Optional[str] = None,
    page: Optional[str] = None,
) -> dict:
    """Connect two shapes with a dynamic connector glued to both (Visio
    auto-picks the best sides and re-routes when shapes move).

    Args:
        label: Optional text on the connector (e.g. 'Yes' / 'No').
        route: 'right_angle' (default), 'straight', or 'curved'.
        line_pattern: 'dashed'/'dotted' are the convention for control-plane,
            auth, or reference flows in architecture diagrams.
        line_weight_pt: Thicker lines (e.g. 2) suit peering/trust boundaries.
        line_color: Hex like '#808080'.
    """
    return await _worker.run(
        _client.connect_shapes, from_id, to_id, label, route,
        end_arrow, begin_arrow, page, line_pattern, line_weight_pt, line_color,
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
async def set_page_size(
    width_in: Optional[float] = None,
    height_in: Optional[float] = None,
    orientation: Optional[Literal["portrait", "landscape"]] = None,
    fit_to_contents: bool = False,
    page: Optional[str] = None,
) -> dict:
    """Resize the page. Do this BEFORE building a wide architecture diagram —
    the default page is US Letter (8.5 x 11 in) and export_page_png crops to
    the page bounds, so shapes placed beyond them are invisible.

    Args:
        width_in / height_in: New page size in inches (e.g. 20 x 12 for a
            hybrid-cloud reference diagram).
        orientation: Print orientation; inferred from the size if omitted.
        fit_to_contents: True = shrink/grow the page to fit what's on it
            (alternative to explicit width/height).
    """
    return await _worker.run(
        _client.set_page_size, width_in, height_in, orientation, fit_to_contents, page
    )


@mcp.tool()
@_tool_errors
async def drop_text(
    text: str,
    x: float,
    y: float,
    width_in: Optional[float] = None,
    height_in: Optional[float] = None,
    font_size_pt: float = 10.0,
    bold: bool = False,
    text_color: Optional[str] = None,
    align: Literal["left", "center", "right"] = "center",
    page: Optional[str] = None,
) -> dict:
    """Drop a text-only label (no border, no fill) — for diagram titles,
    legends, and callouts. (x, y) is the CENTER in inches. Size is estimated
    from the text if omitted. Auto-layout ignores unconnected text shapes."""
    return await _worker.run(
        _client.drop_text, text, x, y, width_in, height_in,
        font_size_pt, bold, text_color, align, page,
    )


@mcp.tool()
@_tool_errors
async def add_container(
    label: str,
    member_ids: list[int],
    master: Optional[str] = None,
    padding_in: float = 0.4,
    page: Optional[str] = None,
) -> dict:
    """Wrap existing shapes in a real Visio container (for VNets, VPCs,
    subnets, resource groups, trust zones). The container is sized around its
    members plus padding, and members MOVE WITH the container afterwards.
    Nest zones by creating inner containers first, then outer ones.

    Args:
        label: Container heading text.
        member_ids: shape_ids to place inside.
        master: Optional container style name from Visio's built-in container
            stencil (defaults to the first available style).

    To recolor the container afterwards (e.g. brand colors per zone), use
    style_shape on the returned shape_id — theme guards are overridden.
    """
    return await _worker.run(
        _client.add_container, label, member_ids, master, padding_in, page
    )


@mcp.tool()
@_tool_errors
async def container_members(
    action: Literal["add", "remove"],
    container_id: int,
    member_ids: list[int],
    page: Optional[str] = None,
) -> dict:
    """Add shapes to or remove shapes from an existing container."""
    return await _worker.run(
        _client.container_members, action, container_id, member_ids, page
    )


@mcp.tool()
@_tool_errors
async def get_page_state(page: Optional[str] = None) -> dict:
    """List everything on a page: shape ids, master names, positions (inches,
    center), sizes, text, connector endpoints (from_id/to_id), and container
    membership (container_ids). Call this to (re)discover shape ids, e.g.
    after open_document."""
    return await _worker.run(_client.get_page_state, page)


def main() -> None:
    try:
        mcp.run()  # stdio transport
    finally:
        _worker.shutdown(release=_client.release)


if __name__ == "__main__":
    main()
