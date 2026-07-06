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


def test_connector_line_pattern_weight_color(client, app):
    _build_flowchart(client)
    client.connect_shapes(1, 2, line_pattern="dashed", line_weight_pt=2,
                          line_color="#808080")
    conn = app.ActivePage.shapes[-1]
    assert conn.CellsU("LinePattern").ResultIU == 2
    assert conn.CellsU("LineWeight").FormulaU == "2 pt"
    assert conn.CellsU("LineColor").FormulaU == "RGB(128,128,128)"
    with pytest.raises(VisioMcpError, match="line_pattern"):
        client.connect_shapes(1, 2, line_pattern="wavy")


def test_style_shape_line_pattern(client, app):
    _build_flowchart(client)
    client.style_shape(1, line_pattern="dotted")
    assert app.ActivePage.shapes[0].CellsU("LinePattern").ResultIU == 3
    with pytest.raises(VisioMcpError, match="line_pattern"):
        client.style_shape(1, line_pattern="zigzag")


def test_set_page_size_explicit_and_orientation(client, app):
    client.create_document()
    result = client.set_page_size(20, 12)
    assert result["page_size_in"] == [20, 12]
    sheet = app.ActivePage.PageSheet
    assert sheet.CellsU("PrintPageOrientation").ResultIU == 2  # landscape inferred
    client.set_page_size(orientation="portrait")
    assert sheet.CellsU("PrintPageOrientation").ResultIU == 1


def test_set_page_size_fit_to_contents(client, app):
    _build_flowchart(client)
    client.set_page_size(fit_to_contents=True)
    assert app.ActivePage.resize_calls == 1


def test_drop_text_creates_borderless_unfilled_label(client, app):
    client.create_document()
    result = client.drop_text("Hybrid Cloud Reference", 5, 10.5,
                              font_size_pt=18, bold=True, text_color="#333333",
                              align="left")
    shape = app.ActivePage.shapes[0]
    assert result["is_text"] is True
    assert result["text"] == "Hybrid Cloud Reference"
    assert (result["x"], result["y"]) == (5, 10.5)
    assert shape.CellsU("LinePattern").ResultIU == 0
    assert shape.CellsU("FillPattern").ResultIU == 0
    assert shape.CellsU("Char.Size").FormulaU == "18 pt"
    assert int(shape.CellsU("Char.Style").ResultIU) & 1
    assert shape.CellsU("Char.Color").FormulaU == "RGB(51,51,51)"
    assert shape.CellsU("Para.HorzAlign").ResultIU == 0
    with pytest.raises(VisioMcpError, match="align"):
        client.drop_text("x", 1, 1, align="justified")
    with pytest.raises(VisioMcpError, match="empty"):
        client.drop_text("   ", 1, 1)


def test_add_container_wraps_members_and_reports_membership(client, app):
    _build_flowchart(client)
    result = client.add_container("VNet", [1, 2])
    # SendToBack: the container must sit behind its members in z-order
    container = app.ActivePage.shapes[0]
    assert result["shape_id"] == container.ID
    assert result["member_ids"] == [1, 2]
    assert container.Text == "VNet"
    assert container.ContainerProperties is not None
    assert app.undo_scopes[-1][1] is True, "add_container must commit its undo scope"
    # bbox of members (both 1.0x0.75 at x=4, y in {9,7}) + 0.4 padding
    assert result["x"] == 4
    assert result["y"] == 8
    assert result["width_in"] == pytest.approx(1.0 + 0.8)
    assert result["height_in"] == pytest.approx(2.75 + 0.8)
    state = client.get_page_state()
    by_id = {s["shape_id"]: s for s in state["shapes"]}
    assert by_id[1]["container_ids"] == [container.ID]
    assert by_id[3]["container_ids"] == []
    assert by_id[container.ID]["is_container"] is True


def test_container_members_add_remove(client, app):
    _build_flowchart(client)
    container_id = client.add_container("Zone", [1])["shape_id"]
    client.container_members("add", container_id, [2, 3])
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    assert state[2]["container_ids"] == [container_id]
    client.container_members("remove", container_id, [2])
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    assert state[2]["container_ids"] == []
    with pytest.raises(VisioMcpError, match="not a container"):
        client.container_members("add", 1, [2])
    with pytest.raises(VisioMcpError, match="add.*remove"):
        client.container_members("move", container_id, [2])


def test_add_container_requires_members(client):
    _build_flowchart(client)
    with pytest.raises(VisioMcpError, match="at least one"):
        client.add_container("Empty", [])


def test_add_container_reuses_open_container_stencil(client, app):
    _build_flowchart(client)
    client.add_container("Outer", [1])
    client.add_container("Inner", [2])
    container_opens = [n for n in app.Documents.open_ex_calls if "CONTAINER" in n.upper()]
    assert len(container_opens) == 1, "built-in container stencil must be opened once"


def test_add_container_failure_rolls_back_undo_scope(client, app):
    _build_flowchart(client)
    # make the built-in container stencil produce NON-container masters so
    # the post-drop 'is it really a container?' check fails mid-build
    app.known_stencils["BOXES_U.VSSX"] = ["Plain box"]
    app.GetBuiltInStencilFile = lambda t, m: "BOXES_U.VSSX"
    with pytest.raises(VisioMcpError, match="not a Visio container"):
        client.add_container("Zone", [1])
    assert app.undo_scopes[-1][1] is False, "failed add_container must roll back"


def test_add_container_ambiguous_master_is_rejected(client):
    _build_flowchart(client)
    with pytest.raises(VisioMcpError, match="ambiguous"):
        client.add_container("Zone", [1], master="container")
    result = client.add_container("Zone", [1], master="Container 1")
    assert result["member_ids"] == [1]


def test_add_container_bbox_respects_off_center_pin(client, app):
    _build_flowchart(client)
    # simulate an icon master anchored at bottom-left: LocPin (0, 0) means the
    # shape body extends from Pin to Pin+Width/Height
    shape = app.ActivePage.shapes[0]  # 1.0 x 0.75 at pin (4, 9)
    shape.CellsU("LocPinX").ResultIU = 0
    shape.CellsU("LocPinY").ResultIU = 0
    result = client.add_container("Zone", [1], padding_in=0.5)
    assert result["x"] == pytest.approx(4.5)   # (4 + 5) / 2
    assert result["y"] == pytest.approx(9.375)  # (9 + 9.75) / 2
    assert result["width_in"] == pytest.approx(2.0)
    assert result["height_in"] == pytest.approx(1.75)


def test_set_page_size_validates_before_mutating(client, app):
    client.create_document()
    with pytest.raises(VisioMcpError, match="orientation"):
        client.set_page_size(20, 12, orientation="diagonal")
    sheet = app.ActivePage.PageSheet
    assert sheet.CellsU("PageWidth").ResultIU == 8.5, "page must not be resized on invalid input"
    with pytest.raises(VisioMcpError, match="positive"):
        client.set_page_size(-3)


def test_drop_text_rejects_zero_size(client):
    client.create_document()
    with pytest.raises(VisioMcpError, match="positive"):
        client.drop_text("Title", 1, 1, width_in=0)


def test_connector_explicit_solid_sets_line_pattern(client, app):
    _build_flowchart(client)
    client.connect_shapes(1, 2, line_pattern="solid")
    conn = app.ActivePage.shapes[-1]
    assert conn.CellsU("LinePattern").ResultIU == 1, \
        "explicit solid must be written to override themed defaults"


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
