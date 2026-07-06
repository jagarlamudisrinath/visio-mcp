"""VisioClient end-to-end against FakeVisio."""

import pytest

from visio_mcp import constants as C
from visio_mcp.errors import VisioMcpError
from visio_mcp.visio_client import VisioClient

from .fake_visio import FakeApplication


@pytest.fixture
def app():
    return FakeApplication()


@pytest.fixture
def client(app):
    return VisioClient(app_factory=lambda: app)


def _build_flowchart(client):
    """create doc + stencil, drop start/process/decision, return drop result."""
    client.create_document("BASFLO_U.VSTX")
    client.open_stencil("BASFLO_U.VSSX")
    return client.drop_shapes([
        {"master": "Start/End", "x": 4, "y": 9, "text": "Start"},
        {"master": "Process", "x": 4, "y": 7, "text": "Do work"},
        {"master": "Decision", "x": 4, "y": 5, "text": "OK?"},
    ])


def test_status_launches_visio_and_reports_state(client, app):
    result = client.status()
    assert app.Visible is True
    assert app.AlertResponse == 7
    assert result["visio_version"] == "16.0 (fake)"
    assert result["documents"] == []
    assert result["my_shapes_path"] == app.MyShapesPath


def test_create_document_with_bare_template_name_gets_units_suffix(client, app):
    result = client.create_document("BASFLO", measurement="metric")
    assert result["document"] == "BASFLO_M.VSTX"
    assert result["pages"] == ["Page-1"]


def test_create_document_unknown_template_is_actionable(client):
    with pytest.raises(VisioMcpError, match="NOPE_U.VSTX"):
        client.create_document("NOPE")


def test_drop_requires_document(client):
    with pytest.raises(VisioMcpError, match="create_document"):
        client.drop_shapes([{"master": "Process", "x": 1, "y": 1}])


def test_find_masters_requires_stencil(client):
    client.create_document()
    with pytest.raises(VisioMcpError, match="open_stencil"):
        client.find_masters("process")


def test_stencil_open_and_master_search(client):
    client.create_document()
    info = client.open_stencil("BASFLO_U.VSSX")
    assert info["master_count"] == 5
    found = client.find_masters("dec")
    assert [m["master"] for m in found["masters"]] == ["Decision"]
    everything = client.find_masters()
    assert len(everything["masters"]) == 5


def test_master_not_found_suggests_close_matches(client):
    client.create_document()
    client.open_stencil("BASFLO_U.VSSX")
    with pytest.raises(VisioMcpError, match="Decision"):
        client.drop_shapes([{"master": "Decsion", "x": 1, "y": 1}])


def test_drop_shapes_positions_text_and_undo_scope(client, app):
    result = _build_flowchart(client)
    shapes = result["shapes"]
    assert [s["shape_id"] for s in shapes] == [1, 2, 3]
    assert shapes[0]["text"] == "Start"
    assert shapes[1]["master"] == "Process"
    assert (shapes[2]["x"], shapes[2]["y"]) == (4, 5)
    assert app.undo_scopes == [(1, True)], "bulk drop must commit one undo scope"


def test_drop_with_size_override(client):
    client.create_document()
    client.open_stencil("BASFLO_U.VSSX")
    result = client.drop_shapes([
        {"master": "Process", "x": 2, "y": 2, "width_in": 2.5, "height_in": 1.25},
    ])
    assert result["shapes"][0]["width_in"] == 2.5
    assert result["shapes"][0]["height_in"] == 1.25


def test_connect_shapes_glues_routes_and_arrows(client, app):
    _build_flowchart(client)
    result = client.connect_shapes(1, 2, label="next", route="right_angle")
    page = app.ActivePage
    conn = page.shapes[-1]
    assert result["connector_id"] == conn.ID
    assert conn.OneD == 1
    assert conn.Text == "next"
    assert page.glue_log == [
        (conn.ID, "BeginX", 1, "PinX"),
        (conn.ID, "EndX", 2, "PinX"),
    ]
    assert conn.CellsU("ShapeRouteStyle").ResultIU == C.ROUTE_RIGHT_ANGLE
    assert conn.CellsU("ConLineRouteExt").ResultIU == C.ROUTE_EXT_STRAIGHT
    assert conn.CellsU("EndArrow").FormulaU == "5"
    assert conn.CellsU("BeginArrow").FormulaU == "0"


def test_connect_curved_route(client, app):
    _build_flowchart(client)
    client.connect_shapes(2, 3, route="curved", end_arrow=False)
    conn = app.ActivePage.shapes[-1]
    assert conn.CellsU("ConLineRouteExt").ResultIU == C.ROUTE_EXT_NURBS
    assert conn.CellsU("EndArrow").FormulaU == "0"


def test_connect_unknown_route_rejected(client):
    _build_flowchart(client)
    with pytest.raises(VisioMcpError, match="route"):
        client.connect_shapes(1, 2, route="zigzag")


def test_style_shape_sets_formulas(client, app):
    _build_flowchart(client)
    client.style_shape(2, fill_color="#0078D4", line_color="#004578",
                       text_color="#FFFFFF", line_weight_pt=2, font_size_pt=12, bold=True)
    shape = app.ActivePage.shapes[1]
    assert shape.CellsU("FillForegnd").FormulaU == "RGB(0,120,212)"
    assert shape.CellsU("LineColor").FormulaU == "RGB(0,69,120)"
    assert shape.CellsU("Char.Color").FormulaU == "RGB(255,255,255)"
    assert shape.CellsU("LineWeight").FormulaU == "2 pt"
    assert shape.CellsU("Char.Size").FormulaU == "12 pt"
    assert int(shape.CellsU("Char.Style").ResultIU) & 1


def test_style_shape_bad_color_rejected(client):
    _build_flowchart(client)
    with pytest.raises(ValueError, match="hex"):
        client.style_shape(1, fill_color="blue")


def test_update_and_delete_shape(client, app):
    _build_flowchart(client)
    updated = client.update_shape(2, text="Renamed", x=6.0, width_in=3.0)
    assert updated["text"] == "Renamed"
    assert updated["x"] == 6.0
    assert updated["width_in"] == 3.0
    client.delete_shapes([2])
    with pytest.raises(VisioMcpError, match="get_page_state"):
        client.update_shape(2, text="gone")


def test_stale_shape_id_is_actionable(client):
    _build_flowchart(client)
    with pytest.raises(VisioMcpError, match="get_page_state"):
        client.connect_shapes(1, 99)


def test_pages_list_add_activate(client, app):
    client.create_document()
    added = client.pages("add", "Details")
    assert added["page"] == "Details"
    client.pages("activate", "Details")
    listing = client.pages("list")
    assert [(p["name"], p["active"]) for p in listing["pages"]] == [
        ("Page-1", False), ("Details", True),
    ]
    with pytest.raises(VisioMcpError, match="list, add, or activate"):
        client.pages("rename", "X")


def test_auto_layout_sets_cells_and_runs_layout(client, app):
    _build_flowchart(client)
    client.auto_layout("flowchart_tb", spacing_in=0.5)
    page = app.ActivePage
    sheet = page.PageSheet
    assert sheet.CellsU("PlaceStyle").ResultIU == C.PLACE_TOP_TO_BOTTOM
    assert sheet.CellsU("RouteStyle").ResultIU == C.ROUTE_FLOWCHART_TB
    assert sheet.CellsU("AvenueSizeX").ResultIU == 0.5
    assert page.layout_calls == 1
    assert page.resize_calls == 1
    with pytest.raises(VisioMcpError, match="style"):
        client.auto_layout("diagonal")


def test_get_page_state_reports_shapes_and_connector_endpoints(client):
    _build_flowchart(client)
    client.connect_shapes(1, 2, label="go")
    state = client.get_page_state()
    assert state["page"] == "Page-1"
    assert state["page_size_in"] == [8.5, 11.0]
    by_id = {s["shape_id"]: s for s in state["shapes"]}
    assert by_id[1]["master"] == "Start/End"
    assert by_id[1]["is_connector"] is False
    conn = by_id[4]
    assert conn["is_connector"] is True
    assert (conn["from_id"], conn["to_id"]) == (1, 2)
    assert conn["text"] == "go"


def test_save_and_export_png(client, app, tmp_path):
    _build_flowchart(client)
    with pytest.raises(VisioMcpError, match="never been saved"):
        client.save_document()
    saved = client.save_document(str(tmp_path / "flow"))
    assert saved["path"].endswith("flow.vsdx")
    assert client.save_document()["path"] == saved["path"]

    exported = client.export_page_png(str(tmp_path / "flow"))
    assert exported["path"].endswith("flow.png")
    with open(exported["path"], "rb") as f:
        assert f.read(8) == b"\x89PNG\r\n\x1a\n"


def test_open_document_roundtrip(client, app, tmp_path):
    _build_flowchart(client)
    path = str(tmp_path / "saved.vsdx")
    client.save_document(path)
    result = client.open_document(path)
    assert result["document"] == "saved.vsdx"
    with pytest.raises(VisioMcpError, match="File not found"):
        client.open_document(str(tmp_path / "missing.vsdx"))


def test_reattaches_if_user_quit_visio(app):
    """If the cached app proxy dies, the client transparently reattaches."""
    apps = [app, FakeApplication()]
    client = VisioClient(app_factory=lambda: apps.pop(0))
    client.status()

    class Dead:
        def __getattr__(self, name):
            raise OSError("RPC server unavailable")

    client._app_obj = Dead()
    assert client.status()["visio_version"] == "16.0 (fake)"
