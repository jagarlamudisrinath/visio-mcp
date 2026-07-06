"""The 16 MCP tools register with correct names and schemas."""

import sys

import pytest

from visio_mcp import server

EXPECTED_TOOLS = {
    "visio_status",
    "create_document",
    "open_document",
    "save_document",
    "export_page_png",
    "open_stencil",
    "find_masters",
    "drop_shape",
    "drop_shapes",
    "update_shape",
    "style_shape",
    "delete_shapes",
    "connect_shapes",
    "pages",
    "auto_layout",
    "get_page_state",
}


async def test_all_sixteen_tools_registered():
    tools = await server.mcp.list_tools()
    names = {t.name for t in tools}
    assert names == EXPECTED_TOOLS
    assert len(EXPECTED_TOOLS) == 16


async def test_every_tool_has_description_mentioning_behavior():
    for tool in await server.mcp.list_tools():
        assert tool.description and len(tool.description) > 20, tool.name


async def test_drop_shape_schema_marks_required_fields():
    tools = {t.name: t for t in await server.mcp.list_tools()}
    schema = tools["drop_shape"].inputSchema
    assert set(schema["required"]) == {"master", "x", "y"}


async def test_connect_shapes_route_enum():
    tools = {t.name: t for t in await server.mcp.list_tools()}
    schema = tools["connect_shapes"].inputSchema
    assert schema["properties"]["route"]["enum"] == ["right_angle", "straight", "curved"]


async def test_auto_layout_style_enum():
    tools = {t.name: t for t in await server.mcp.list_tools()}
    schema = tools["auto_layout"].inputSchema
    assert schema["properties"]["style"]["enum"] == [
        "flowchart_tb", "flowchart_lr", "tree_tb", "tree_lr", "radial", "circular",
    ]


async def test_instructions_document_coordinate_system():
    text = server.mcp.instructions
    assert "INCHES" in text
    assert "BOTTOM-LEFT" in text


@pytest.mark.skipif(sys.platform == "win32", reason="non-Windows guard message")
async def test_tool_error_surface_on_this_platform():
    """On macOS/Linux the tools fail with a clean, actionable ToolError
    ('must run on Windows') instead of crashing the protocol."""
    from mcp.server.fastmcp.exceptions import ToolError

    with pytest.raises(ToolError, match="must run on Windows"):
        await server.mcp.call_tool("visio_status", {})
