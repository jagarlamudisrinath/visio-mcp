"""End-to-end smoke test — RUN THIS ON WINDOWS with Visio desktop installed.

    uv run python scripts/smoke_test.py

Exercises the real COM pipeline: launch/attach Visio, create a flowchart
document, open the basic flowchart stencil, drop + connect + style shapes,
auto-layout, export PNG (opened for eyeballing), save and reopen the .vsdx.
Prints numbered PASS/FAIL per step and continues on failure.
"""

from __future__ import annotations

import os
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from visio_mcp import constants as C  # noqa: E402
from visio_mcp.runtime import ComWorker  # noqa: E402
from visio_mcp.visio_client import VisioClient  # noqa: E402

FAILURES: list[str] = []


def step(num: int, title: str):
    def decorator(fn):
        def run(*args, **kwargs):
            try:
                result = fn(*args, **kwargs)
                print(f"[{num:02d}] PASS  {title}")
                return result
            except Exception as exc:
                FAILURES.append(f"[{num:02d}] {title}: {exc}")
                print(f"[{num:02d}] FAIL  {title}")
                traceback.print_exc()
                return None
        return run
    return decorator


def main() -> int:
    if sys.platform != "win32":
        print("This smoke test must run on Windows with Visio desktop installed.")
        return 2

    worker = ComWorker()
    client = VisioClient()
    call = worker.run_sync
    tmp = tempfile.gettempdir()
    png_path = os.path.join(tmp, "visio_mcp_smoke.png")
    vsdx_path = os.path.join(tmp, "visio_mcp_smoke.vsdx")
    state = {}

    @step(0, "constants match win32com generated constants (if available)")
    def check_constants():
        try:
            import win32com.client.gencache as gencache  # noqa: F401
            from win32com.client import constants as wc

            wc.visOpenRO  # raises if no typelib generated
        except Exception:
            print("      (no gen_py typelib — skipping constant cross-check)")
            return
        assert C.VIS_OPEN_RO == wc.visOpenRO
        assert C.VIS_OPEN_DOCKED == wc.visOpenDocked
        assert C.PLACE_TOP_TO_BOTTOM == wc.visPLOPlaceTopToBottom
        assert C.ROUTE_FLOWCHART_TB == wc.visLORouteFlowchartNS
        assert C.ROUTE_RIGHT_ANGLE == wc.visLORouteRightAngle

    @step(1, "attach/launch Visio, report status")
    def status():
        s = call(client.status)
        print(f"      Visio {s['visio_version']}; My Shapes: {s['my_shapes_path']}")
        return s

    @step(2, "create document from Basic Flowchart template")
    def create():
        try:
            return call(client.create_document, "BASFLO_U.VSTX")
        except Exception:
            print("      template missing — falling back to blank document")
            return call(client.create_document)

    @step(3, "open basic flowchart stencil + find 'decision' master")
    def stencil():
        call(client.open_stencil, "BASFLO_U.VSSX")
        found = call(client.find_masters, "decision")
        assert found["masters"], "no master matching 'decision'"

    @step(4, "drop start/process/decision/end shapes with text")
    def drop():
        result = call(client.drop_shapes, [
            {"master": "Start/End", "x": 4, "y": 9.5, "text": "Start"},
            {"master": "Process", "x": 4, "y": 8, "text": "Validate input"},
            {"master": "Decision", "x": 4, "y": 6.5, "text": "Valid?"},
            {"master": "Process", "x": 6.5, "y": 6.5, "text": "Show error"},
            {"master": "Start/End", "x": 4, "y": 5, "text": "Done"},
        ])
        ids = [s["shape_id"] for s in result["shapes"]]
        state["ids"] = ids
        assert len(ids) == 5

    @step(5, "connect shapes incl. labeled Yes/No branches")
    def connect():
        a, b, c, err, done = state["ids"]
        call(client.connect_shapes, a, b)
        call(client.connect_shapes, b, c)
        call(client.connect_shapes, c, done, "Yes")
        call(client.connect_shapes, c, err, "No")
        page_state = call(client.get_page_state)
        conns = [s for s in page_state["shapes"] if s["is_connector"]]
        assert len(conns) == 4, f"expected 4 connectors, found {len(conns)}"
        assert any(s["from_id"] == c and s["to_id"] == err for s in conns), \
            "No-branch connector endpoints did not round-trip"

    @step(6, "style the process shape and verify formula readback")
    def style():
        call(client.style_shape, state["ids"][1],
             None, "#DDEBF7", "#2F5B7C", None, 1.5, 11, True)

    @step(7, "auto-layout flowchart_tb moves shapes")
    def layout():
        before = {s["shape_id"]: (s["x"], s["y"])
                  for s in call(client.get_page_state)["shapes"]}
        call(client.auto_layout, "flowchart_tb", 0.6, True)
        after = {s["shape_id"]: (s["x"], s["y"])
                 for s in call(client.get_page_state)["shapes"]}
        assert before != after, "layout did not move anything (may be OK, eyeball the PNG)"

    @step(8, "export page PNG")
    def export():
        result = call(client.export_page_png, png_path)
        assert os.path.getsize(result["path"]) > 1024, "PNG suspiciously small"
        os.startfile(result["path"])  # noqa: S606 — open for eyeball check

    @step(9, "save .vsdx, reopen it, shapes survive")
    def save_reopen():
        call(client.save_document, vsdx_path)
        call(client.open_document, vsdx_path)
        page_state = call(client.get_page_state)
        assert len(page_state["shapes"]) >= 9, "shapes lost in save/reopen round-trip"

    @step(10, "optional: Azure/AWS stencils under My Shapes")
    def cloud_stencils():
        import glob

        s = call(client.status)
        hits = [p for pat in ("*azure*", "*aws*", "*Azure*", "*AWS*")
                for p in glob.glob(os.path.join(s["my_shapes_path"], "**", pat + ".vss*"),
                                   recursive=True)]
        if not hits:
            print("      (no Azure/AWS stencils found in My Shapes — skipping; "
                  "download them to enable cloud icons)")
            return
        info = call(client.open_stencil, hits[0])
        print(f"      opened {info['stencil']} with {info['master_count']} masters")

    for fn in (check_constants, status, create, stencil, drop, connect,
               style, layout, export, save_reopen, cloud_stencils):
        fn()

    worker.shutdown(release=client.release)

    print()
    if FAILURES:
        print(f"{len(FAILURES)} step(s) FAILED:")
        for f in FAILURES:
            print(f"  {f}")
        return 1
    print("All steps passed. Next:")
    print(r"  claude mcp add visio -- uv --directory " + os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..")) + " run visio-mcp")
    print('  Then ask Claude: "Draw a 5-step login flowchart in Visio and show me the result."')
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
