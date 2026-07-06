"""VisioClient: every COM call in the codebase lives in this module.

Rules that keep this correct:
- All methods are called on the ComWorker's STA thread (see runtime.py).
- Methods return plain dicts/lists only — COM proxies must never escape.
- win32com is imported lazily inside the default app factory, so this module
  imports fine on macOS and tests inject a FakeVisio app factory instead.
- COM collections are 1-based and iterated explicitly via Count/Item so a
  duck-typed fake works without implementing COM enumerators.
"""

from __future__ import annotations

import difflib
import os
import sys
from typing import Any, Callable, Optional

from . import constants as C
from .errors import VisioMcpError, hresult_of, translate_com_error
from .models import DropSpec, hex_to_rgb_formula

_CONTROL_CHARS = dict.fromkeys(i for i in range(32) if i not in (9, 10, 13))


def _clean_text(raw: Any) -> str:
    """Visio shape text can embed field-escape control chars; strip them."""
    return str(raw).translate(_CONTROL_CHARS) if raw else ""


def _is_com_error(exc: BaseException) -> bool:
    return type(exc).__name__ == "com_error"


def _items(collection):
    """Iterate a 1-based COM collection via Count/Item."""
    for i in range(1, int(collection.Count) + 1):
        yield collection.Item(i)


def _stencil_files_under(root: str) -> list[str]:
    """All .vss/.vssx files under root, case-insensitive (stencil filenames
    are conventionally uppercase)."""
    hits = []
    for dirpath, _, filenames in os.walk(root):
        for fn in filenames:
            if fn.lower().endswith((".vssx", ".vss")):
                hits.append(os.path.join(dirpath, fn))
    return sorted(hits)


def default_app_factory():
    """Attach to a running Visio instance, or launch a new visible one."""
    if sys.platform != "win32":
        raise VisioMcpError(
            "This server drives Microsoft Visio via COM and must run on Windows "
            "with Visio desktop installed. It cannot control Visio from "
            f"{sys.platform!r}."
        )
    import win32com.client  # noqa: PLC0415 — Windows-only import by design

    try:
        return win32com.client.GetActiveObject("Visio.Application")
    except Exception as exc:
        if not _is_com_error(exc) or hresult_of(exc) != C.MK_E_UNAVAILABLE:
            raise translate_com_error(exc, "attaching to Visio") from exc
    try:
        return win32com.client.DispatchEx("Visio.Application")
    except Exception as exc:
        if _is_com_error(exc):
            raise translate_com_error(exc, "launching Visio") from exc
        raise


class VisioClient:
    def __init__(self, app_factory: Optional[Callable[[], Any]] = None):
        self._app_factory = app_factory or default_app_factory
        self._app_obj: Any = None

    # -- plumbing -----------------------------------------------------------

    def release(self) -> None:
        """Drop COM references. Never quits Visio — the user keeps their app."""
        self._app_obj = None

    def _app(self):
        if self._app_obj is not None:
            try:
                _ = self._app_obj.Version  # liveness probe: user may have quit Visio
            except Exception:
                self._app_obj = None
        if self._app_obj is None:
            app = self._app_factory()
            app.Visible = True
            try:
                # 7 = IDNO: auto-answer modal prompts so COM calls never hang
                app.AlertResponse = 7
            except Exception:
                pass
            self._app_obj = app
        return self._app_obj

    def _guard(self, context: str, fn: Callable[[], Any]) -> Any:
        """Run fn, translating com_error into an actionable VisioMcpError."""
        try:
            return fn()
        except VisioMcpError:
            raise
        except Exception as exc:
            if _is_com_error(exc):
                raise translate_com_error(exc, context) from exc
            raise

    def _drawing_doc(self):
        app = self._app()
        doc = None
        try:
            doc = app.ActiveDocument
        except Exception:
            doc = None
        if doc is not None and int(doc.Type) == C.VIS_TYPE_DRAWING:
            return doc
        for d in _items(app.Documents):
            if int(d.Type) == C.VIS_TYPE_DRAWING:
                return d
        raise VisioMcpError(
            "No drawing document is open in Visio — call create_document or "
            "open_document first."
        )

    def _page(self, page_name: Optional[str] = None):
        app = self._app()
        doc = self._drawing_doc()
        if page_name is None:
            try:
                page = app.ActivePage
                if page is not None:
                    return page
            except Exception:
                pass
            return doc.Pages.Item(1)
        for p in _items(doc.Pages):
            if str(p.Name) == page_name or str(p.NameU) == page_name:
                return p
        names = [str(p.Name) for p in _items(doc.Pages)]
        raise VisioMcpError(f"No page named {page_name!r}. Pages: {names}")

    def _shape_by_id(self, page, shape_id: int):
        try:
            return page.Shapes.ItemFromID(int(shape_id))
        except Exception:
            raise VisioMcpError(
                f"Shape {shape_id} not found on page {str(page.Name)!r} — call "
                "get_page_state to list current shape ids."
            ) from None

    def _stencil_docs(self) -> list:
        app = self._app()
        return [d for d in _items(app.Documents) if int(d.Type) == C.VIS_TYPE_STENCIL]

    # -- status / documents --------------------------------------------------

    def _builtin_content_dir(self) -> Optional[str]:
        """Visio ships ~130 stencils (incl. Azure/AWS) under
        <install>\\Visio Content\\<locale>\\ — resolvable by bare filename."""
        try:
            base = str(self._app().Path)
        except Exception:
            return None
        root = os.path.join(base, "Visio Content")
        return root if os.path.isdir(root) else None

    def _builtin_stencil_files(self) -> list[str]:
        root = self._builtin_content_dir()
        return _stencil_files_under(root) if root else []

    def status(self) -> dict:
        app = self._app()
        docs, stencils = [], []
        for d in _items(app.Documents):
            entry = {"name": str(d.Name), "path": str(d.FullName)}
            if int(d.Type) == C.VIS_TYPE_STENCIL:
                stencils.append(entry)
            else:
                docs.append(entry)
        active_doc = active_page = None
        try:
            if app.ActiveDocument is not None:
                active_doc = str(app.ActiveDocument.Name)
            if app.ActivePage is not None:
                active_page = str(app.ActivePage.Name)
        except Exception:
            pass
        builtin_files = self._builtin_stencil_files()
        cloud = [
            os.path.basename(f)
            for f in builtin_files
            if any(k in os.path.basename(f).lower() for k in ("azure", "aws"))
        ]
        return {
            "visio_version": str(app.Version),
            "documents": docs,
            "stencils": stencils,
            "active_document": active_doc,
            "active_page": active_page,
            "my_shapes_path": str(app.MyShapesPath),
            "builtin_stencils_path": self._builtin_content_dir(),
            "builtin_stencil_count": len(builtin_files),
            "builtin_cloud_stencils": cloud[:60],
            "note": (
                "Built-in stencils open by bare filename, e.g. "
                "open_stencil('AZURESTORAGE_U.VSSX') — no My Shapes install needed."
                if cloud else
                "No built-in Azure/AWS stencils detected; download stencil packs "
                "into the My Shapes folder for cloud icons."
            ),
        }

    def create_document(self, template: Optional[str] = None, measurement: str = "us") -> dict:
        app = self._app()
        name = template or ""
        if name and "." not in os.path.basename(name):
            # bare family name like 'BASFLO' -> BASFLO_U.VSTX / BASFLO_M.VSTX
            name = f"{name}_{'M' if measurement == 'metric' else 'U'}.VSTX"

        def _add():
            return app.Documents.Add(name)

        doc = self._guard(
            f"creating document from template {name!r}" if name else "creating blank document",
            _add,
        )
        return {
            "document": str(doc.Name),
            "pages": [str(p.Name) for p in _items(doc.Pages)],
            "note": "Coordinates are in inches, origin at the page bottom-left; "
                    "drop point is the shape center.",
        }

    def open_document(self, path: str) -> dict:
        app = self._app()
        abs_path = os.path.abspath(os.path.expanduser(path))
        if not os.path.exists(abs_path):
            raise VisioMcpError(f"File not found: {abs_path}")
        doc = self._guard(f"opening {abs_path}", lambda: app.Documents.Open(abs_path))
        return {
            "document": str(doc.Name),
            "path": str(doc.FullName),
            "pages": [str(p.Name) for p in _items(doc.Pages)],
            "note": "Call get_page_state to learn the shape ids on each page.",
        }

    def save_document(self, path: Optional[str] = None) -> dict:
        doc = self._drawing_doc()
        if path is None:
            existing = str(doc.Path or "")
            if not existing:
                raise VisioMcpError(
                    "This document has never been saved — pass an absolute file "
                    "path ending in .vsdx to save_document."
                )
            self._guard("saving document", doc.Save)
            return {"path": str(doc.FullName)}
        abs_path = os.path.abspath(os.path.expanduser(path))
        if not abs_path.lower().endswith((".vsdx", ".vsdm", ".vsd")):
            abs_path += ".vsdx"
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        self._guard(f"saving document to {abs_path}", lambda: doc.SaveAs(abs_path))
        return {"path": str(doc.FullName)}

    def export_page_png(self, path: str, page_name: Optional[str] = None) -> dict:
        page = self._page(page_name)
        abs_path = os.path.abspath(os.path.expanduser(path))
        if not abs_path.lower().endswith(".png"):
            abs_path += ".png"
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        self._guard(f"exporting page to {abs_path}", lambda: page.Export(abs_path))
        if not os.path.exists(abs_path):
            raise VisioMcpError(f"Visio reported success but no file was written at {abs_path}")
        return {"path": abs_path, "page": str(page.Name)}

    # -- stencils / masters ---------------------------------------------------

    def open_stencil(self, name_or_path: str) -> dict:
        app = self._app()
        query = name_or_path.strip()
        candidates: list[str] = []
        if os.path.isabs(query):
            candidates.append(query)
        else:
            candidates.append(query)  # Visio resolves bare names via its search paths
            my_shapes = str(app.MyShapesPath)
            candidates.append(os.path.join(my_shapes, query))
            if "." not in os.path.basename(query):
                candidates.append(os.path.join(my_shapes, query + ".vssx"))
                candidates.append(os.path.join(my_shapes, query + ".vss"))
            # last resort: fuzzy filename match under My Shapes, then Visio's
            # built-in content folder (which ships Azure/AWS stencils)
            fuzzy_pool = _stencil_files_under(my_shapes) + self._builtin_stencil_files()
            for hit in fuzzy_pool:
                if query.lower() in os.path.basename(hit).lower():
                    candidates.append(hit)

        errors: list[str] = []
        for cand in candidates:
            try:
                doc = app.Documents.OpenEx(cand, C.STENCIL_OPEN_FLAGS)
                return {
                    "stencil": str(doc.Name),
                    "path": str(doc.FullName),
                    "master_count": int(doc.Masters.Count),
                    "note": "Use find_masters to list its shapes.",
                }
            except Exception as exc:
                errors.append(f"{cand}: {exc}")
        raise VisioMcpError(
            f"Could not open a stencil matching {name_or_path!r}. Tried Visio's "
            f"search paths, {str(app.MyShapesPath)!r}, and the built-in Visio "
            "Content folder. Call visio_status to see builtin_cloud_stencils "
            "(Visio ships Azure/AWS stencils out of the box); for other vendor "
            "packs, download and unzip into the My Shapes folder. "
            f"Attempts: {errors[:3]}"
        )

    def _master_sources(self) -> list[tuple[str, Any]]:
        """(source_name, masters_collection) for all open stencils + the doc itself."""
        sources = [(str(d.Name), d.Masters) for d in self._stencil_docs()]
        try:
            doc = self._drawing_doc()
            sources.append((f"(document) {doc.Name}", doc.Masters))
        except VisioMcpError:
            pass
        return sources

    def find_masters(self, query: Optional[str] = None, stencil: Optional[str] = None) -> dict:
        sources = self._master_sources()
        if stencil is not None:
            sources = [s for s in sources if stencil.lower() in s[0].lower()]
            if not sources:
                raise VisioMcpError(
                    f"No open stencil matches {stencil!r} — call open_stencil first, "
                    "or omit the stencil argument to search all open stencils."
                )
        total_masters = sum(int(m.Count) for _, m in sources)
        if total_masters == 0:
            raise VisioMcpError(
                "No stencils are open — call open_stencil first (e.g. "
                "open_stencil('BASFLO_U.VSSX') for basic flowchart shapes)."
            )
        rows = []
        q = (query or "").lower()
        for source_name, masters in sources:
            for m in _items(masters):
                name = str(m.Name)
                if q and q not in name.lower():
                    continue
                rows.append({"master": name, "stencil": source_name})
                if len(rows) >= 100:
                    return {"masters": rows, "truncated": True}
        return {"masters": rows, "truncated": False}

    def _find_master(self, master: str, stencil: Optional[str] = None, _alias_depth: int = 0):
        sources = self._master_sources()
        if stencil is not None:
            sources = [s for s in sources if stencil.lower() in s[0].lower()]
        all_names: list[str] = []
        substring_hits = []
        for _, masters in sources:
            for m in _items(masters):
                name = str(m.Name)
                all_names.append(name)
                if name.lower() == master.lower():
                    return m
                if master.lower() in name.lower():
                    substring_hits.append(m)
        if len(substring_hits) == 1:
            return substring_hits[0]
        if len(substring_hits) > 1:
            names = sorted({str(m.Name) for m in substring_hits})[:8]
            raise VisioMcpError(
                f"Master {master!r} is ambiguous — matches {names}. Use the exact name."
            )
        key = master.strip().lower()
        if key in C.MASTER_HINTS:
            raise VisioMcpError(f"No master named {master!r}: {C.MASTER_HINTS[key]}")
        alias = C.MASTER_ALIASES.get(key)
        if alias and _alias_depth == 0:
            try:
                return self._find_master(alias, stencil, _alias_depth=1)
            except VisioMcpError:
                pass  # alias target not in the open stencils either
        suggestions = difflib.get_close_matches(master, all_names, n=5, cutoff=0.4)
        hint = f" Closest matches: {suggestions}." if suggestions else ""
        alias_hint = (
            f" This is usually drawn with the {alias!r} master — open the "
            "stencil that contains it." if alias else ""
        )
        raise VisioMcpError(
            f"No master named {master!r} found in the open stencils.{alias_hint}{hint} "
            "Call find_masters to browse, or open_stencil to load the stencil "
            "that contains it."
        )

    # -- shapes ---------------------------------------------------------------

    def _shape_info(self, shape) -> dict:
        return {
            "shape_id": int(shape.ID),
            "name": str(shape.NameU),
            "x": round(float(shape.CellsU("PinX").ResultIU), 4),
            "y": round(float(shape.CellsU("PinY").ResultIU), 4),
            "width_in": round(float(shape.CellsU("Width").ResultIU), 4),
            "height_in": round(float(shape.CellsU("Height").ResultIU), 4),
            "text": _clean_text(shape.Text),
        }

    def drop_shapes(self, specs: list[dict], page_name: Optional[str] = None) -> dict:
        app = self._app()
        page = self._page(page_name)
        parsed = [DropSpec.model_validate(s) for s in specs]
        results = []
        scope = self._guard("starting undo scope", lambda: app.BeginUndoScope("Drop shapes"))
        try:
            for spec in parsed:
                m = self._find_master(spec.master, spec.stencil)
                shape = self._guard(
                    f"dropping master {spec.master!r}",
                    lambda m=m, s=spec: page.Drop(m, s.x, s.y),
                )
                if spec.width_in is not None:
                    shape.CellsU("Width").ResultIU = spec.width_in
                if spec.height_in is not None:
                    shape.CellsU("Height").ResultIU = spec.height_in
                if spec.text is not None:
                    shape.Text = spec.text
                results.append({**self._shape_info(shape), "master": str(m.Name)})
            self._guard("closing undo scope", lambda: app.EndUndoScope(scope, True))
        except Exception:
            try:
                app.EndUndoScope(scope, False)
            except Exception:
                pass
            raise
        return {"page": str(page.Name), "shapes": results}

    def update_shape(
        self,
        shape_id: int,
        page_name: Optional[str] = None,
        text: Optional[str] = None,
        x: Optional[float] = None,
        y: Optional[float] = None,
        width_in: Optional[float] = None,
        height_in: Optional[float] = None,
    ) -> dict:
        page = self._page(page_name)
        shape = self._shape_by_id(page, shape_id)

        def _apply():
            if text is not None:
                shape.Text = text
            if x is not None:
                shape.CellsU("PinX").ResultIU = x
            if y is not None:
                shape.CellsU("PinY").ResultIU = y
            if width_in is not None:
                shape.CellsU("Width").ResultIU = width_in
            if height_in is not None:
                shape.CellsU("Height").ResultIU = height_in

        self._guard(f"updating shape {shape_id}", _apply)
        return self._shape_info(shape)

    def style_shape(
        self,
        shape_id: int,
        page_name: Optional[str] = None,
        fill_color: Optional[str] = None,
        line_color: Optional[str] = None,
        text_color: Optional[str] = None,
        line_weight_pt: Optional[float] = None,
        font_size_pt: Optional[float] = None,
        bold: Optional[bool] = None,
        line_pattern: Optional[str] = None,
    ) -> dict:
        if line_pattern is not None and line_pattern not in C.LINE_PATTERNS:
            raise VisioMcpError(
                f"line_pattern must be one of {sorted(C.LINE_PATTERNS)}, got {line_pattern!r}"
            )
        page = self._page(page_name)
        shape = self._shape_by_id(page, shape_id)

        def _apply():
            # FormulaForceU writes through cell guards — Visio's built-in
            # container/theme masters guard LineColor/Char.* cells, and plain
            # FormulaU raises 'Cell is guarded' on them
            if line_pattern is not None:
                shape.CellsU("LinePattern").FormulaForceU = str(C.LINE_PATTERNS[line_pattern])
            if fill_color is not None:
                shape.CellsU("FillForegnd").FormulaForceU = hex_to_rgb_formula(fill_color)
            if line_color is not None:
                shape.CellsU("LineColor").FormulaForceU = hex_to_rgb_formula(line_color)
            if text_color is not None:
                shape.CellsU("Char.Color").FormulaForceU = hex_to_rgb_formula(text_color)
            if line_weight_pt is not None:
                shape.CellsU("LineWeight").FormulaForceU = f"{line_weight_pt} pt"
            if font_size_pt is not None:
                shape.CellsU("Char.Size").FormulaForceU = f"{font_size_pt} pt"
            if bold is not None:
                current = int(shape.CellsU("Char.Style").ResultIU)
                shape.CellsU("Char.Style").FormulaForceU = str(
                    (current | 1) if bold else (current & ~1)
                )

        self._guard(f"styling shape {shape_id}", _apply)
        return {"shape_id": int(shape.ID), "styled": True}

    def delete_shapes(self, shape_ids: list[int], page_name: Optional[str] = None) -> dict:
        page = self._page(page_name)
        for sid in shape_ids:
            shape = self._shape_by_id(page, sid)
            self._guard(f"deleting shape {sid}", shape.Delete)
        return {"deleted": len(shape_ids)}

    # -- connectors -----------------------------------------------------------

    def connect_shapes(
        self,
        from_id: int,
        to_id: int,
        label: Optional[str] = None,
        route: str = "right_angle",
        end_arrow: bool = True,
        begin_arrow: bool = False,
        page_name: Optional[str] = None,
        line_pattern: str = "solid",
        line_weight_pt: Optional[float] = None,
        line_color: Optional[str] = None,
    ) -> dict:
        if route not in C.CONNECTOR_ROUTES:
            raise VisioMcpError(
                f"route must be one of {sorted(C.CONNECTOR_ROUTES)}, got {route!r}"
            )
        if line_pattern not in C.LINE_PATTERNS:
            raise VisioMcpError(
                f"line_pattern must be one of {sorted(C.LINE_PATTERNS)}, got {line_pattern!r}"
            )
        app = self._app()
        page = self._page(page_name)
        shape_from = self._shape_by_id(page, from_id)
        shape_to = self._shape_by_id(page, to_id)

        def _connect():
            conn = page.Drop(app.ConnectorToolDataObject, 0.0, 0.0)
            # Gluing to PinX = dynamic glue: Visio picks the best sides and
            # re-routes as the layout changes.
            conn.CellsU("BeginX").GlueTo(shape_from.CellsU("PinX"))
            conn.CellsU("EndX").GlueTo(shape_to.CellsU("PinX"))
            route_style, route_ext = C.CONNECTOR_ROUTES[route]
            conn.CellsU("ShapeRouteStyle").ResultIU = route_style
            conn.CellsU("ConLineRouteExt").ResultIU = route_ext
            conn.CellsU("EndArrow").FormulaU = str(
                C.ARROW_FILLED_TRIANGLE if end_arrow else C.ARROW_NONE
            )
            conn.CellsU("BeginArrow").FormulaU = str(
                C.ARROW_FILLED_TRIANGLE if begin_arrow else C.ARROW_NONE
            )
            # always set: an explicit 'solid' must override themed defaults
            conn.CellsU("LinePattern").ResultIU = C.LINE_PATTERNS[line_pattern]
            if line_weight_pt is not None:
                conn.CellsU("LineWeight").FormulaU = f"{line_weight_pt} pt"
            if line_color is not None:
                conn.CellsU("LineColor").FormulaU = hex_to_rgb_formula(line_color)
            if label:
                conn.Text = label
            return conn

        conn = self._guard(f"connecting shape {from_id} -> {to_id}", _connect)
        return {
            "connector_id": int(conn.ID),
            "from_id": from_id,
            "to_id": to_id,
            "route": route,
        }

    # -- pages ----------------------------------------------------------------

    def pages(self, action: str, name: Optional[str] = None) -> dict:
        app = self._app()
        doc = self._drawing_doc()
        if action == "list":
            active = None
            try:
                active = str(app.ActivePage.Name) if app.ActivePage is not None else None
            except Exception:
                pass
            return {
                "pages": [
                    {"name": str(p.Name), "index": int(p.Index), "active": str(p.Name) == active}
                    for p in _items(doc.Pages)
                ]
            }
        if action == "add":
            page = self._guard("adding page", doc.Pages.Add)
            if name:
                page.Name = name
            return {"page": str(page.Name), "index": int(page.Index)}
        if action == "activate":
            if not name:
                raise VisioMcpError("pages(action='activate') requires a page name")
            page = self._page(name)
            self._guard(
                f"activating page {name!r}",
                lambda: setattr(app.ActiveWindow, "Page", page),
            )
            return {"active_page": str(page.Name)}
        raise VisioMcpError(f"Unknown pages action {action!r}; use list, add, or activate.")

    # -- page size / text / containers ------------------------------------------

    def set_page_size(
        self,
        width_in: Optional[float] = None,
        height_in: Optional[float] = None,
        orientation: Optional[str] = None,
        fit_to_contents: bool = False,
        page_name: Optional[str] = None,
    ) -> dict:
        # validate everything BEFORE mutating the document
        if orientation not in (None, "portrait", "landscape"):
            raise VisioMcpError(
                f"orientation must be 'portrait' or 'landscape', got {orientation!r}"
            )
        for name, value in (("width_in", width_in), ("height_in", height_in)):
            if value is not None and value <= 0:
                raise VisioMcpError(f"{name} must be positive, got {value}")
        page = self._page(page_name)

        def _apply():
            sheet = page.PageSheet
            if fit_to_contents:
                page.ResizeToFitContents()
            if width_in is not None:
                sheet.CellsU("PageWidth").ResultIU = width_in
            if height_in is not None:
                sheet.CellsU("PageHeight").ResultIU = height_in
            w = float(sheet.CellsU("PageWidth").ResultIU)
            h = float(sheet.CellsU("PageHeight").ResultIU)
            if orientation == "landscape" or (orientation is None and w > h):
                sheet.CellsU("PrintPageOrientation").ResultIU = C.ORIENTATION_LANDSCAPE
            else:
                sheet.CellsU("PrintPageOrientation").ResultIU = C.ORIENTATION_PORTRAIT
            return [round(w, 3), round(h, 3)]

        size = self._guard("setting page size", _apply)
        return {"page": str(page.Name), "page_size_in": size}

    def drop_text(
        self,
        text: str,
        x: float,
        y: float,
        width_in: Optional[float] = None,
        height_in: Optional[float] = None,
        font_size_pt: float = 10.0,
        bold: bool = False,
        text_color: Optional[str] = None,
        align: str = "center",
        page_name: Optional[str] = None,
    ) -> dict:
        aligns = {"left": C.ALIGN_LEFT, "center": C.ALIGN_CENTER, "right": C.ALIGN_RIGHT}
        if align not in aligns:
            raise VisioMcpError(f"align must be one of {sorted(aligns)}, got {align!r}")
        if not text.strip():
            raise VisioMcpError("text must not be empty")
        for name, value in (("width_in", width_in), ("height_in", height_in)):
            if value is not None and value <= 0:
                raise VisioMcpError(f"{name} must be positive, got {value}")
        page = self._page(page_name)
        # rough auto-size from the longest line when no explicit size given
        longest = max(len(line) for line in text.splitlines())
        w = width_in if width_in is not None else max(1.0, longest * font_size_pt * 0.009)
        h = height_in if height_in is not None else max(0.3, len(text.splitlines()) * font_size_pt * 0.02)

        def _drop():
            # Visio's standard text-only shape: a borderless, unfilled rectangle
            shape = page.DrawRectangle(x - w / 2, y - h / 2, x + w / 2, y + h / 2)
            shape.CellsU("LinePattern").ResultIU = 0
            shape.CellsU("FillPattern").ResultIU = 0
            shape.Text = text
            shape.CellsU("Char.Size").FormulaU = f"{font_size_pt} pt"
            if bold:
                cell = shape.CellsU("Char.Style")
                cell.ResultIU = int(cell.ResultIU) | 1
            if text_color is not None:
                shape.CellsU("Char.Color").FormulaU = hex_to_rgb_formula(text_color)
            shape.CellsU("Para.HorzAlign").ResultIU = aligns[align]
            return shape

        shape = self._guard("dropping text", _drop)
        return {**self._shape_info(shape), "is_text": True}

    def _container_master(self, master_name: Optional[str]):
        app = self._app()
        path = self._guard(
            "locating the built-in container stencil",
            lambda: app.GetBuiltInStencilFile(C.VIS_STENCIL_CONTAINERS, C.VIS_MS_US),
        )
        # reuse the stencil if already open — reopening an open stencil can
        # raise on real Visio
        base = os.path.basename(str(path))
        stencil = next(
            (d for d in self._stencil_docs() if str(d.Name) == base), None
        )
        if stencil is None:
            stencil = self._guard(
                "opening the built-in container stencil",
                lambda: app.Documents.OpenEx(path, C.STENCIL_OPEN_FLAGS),
            )
        if master_name:
            return self._find_master(master_name, stencil=str(stencil.Name))
        return stencil.Masters.Item(1)

    def add_container(
        self,
        label: str,
        member_ids: list[int],
        master: Optional[str] = None,
        padding_in: float = 0.4,
        page_name: Optional[str] = None,
    ) -> dict:
        if not member_ids:
            raise VisioMcpError("member_ids must list at least one shape to contain")
        app = self._app()
        page = self._page(page_name)
        members = [self._shape_by_id(page, sid) for sid in member_ids]
        container_master = self._container_master(master)

        def _build():
            # bounding box of the members, in page coordinates; the local pin
            # is not necessarily the shape center (icon masters often anchor
            # bottom-center), so subtract LocPin rather than assuming Width/2
            lefts, rights, bottoms, tops = [], [], [], []
            for m in members:
                left = float(m.CellsU("PinX").ResultIU) - float(m.CellsU("LocPinX").ResultIU)
                bottom = float(m.CellsU("PinY").ResultIU) - float(m.CellsU("LocPinY").ResultIU)
                lefts.append(left)
                rights.append(left + float(m.CellsU("Width").ResultIU))
                bottoms.append(bottom)
                tops.append(bottom + float(m.CellsU("Height").ResultIU))
            cx = (min(lefts) + max(rights)) / 2
            cy = (min(bottoms) + max(tops)) / 2
            container = page.Drop(container_master, cx, cy)
            container.CellsU("Width").ResultIU = max(rights) - min(lefts) + 2 * padding_in
            container.CellsU("Height").ResultIU = max(tops) - min(bottoms) + 2 * padding_in
            container.Text = label
            props = container.ContainerProperties
            if props is None:
                raise VisioMcpError(
                    "The dropped shape is not a Visio container — try a different "
                    "container master."
                )
            for m in members:
                props.AddMember(m, C.VIS_MEMBER_ADD_EXPAND_CONTAINER)
            # Drop puts the container on top of its members; push it behind
            # them so an opaque container fill can't hide the shapes
            container.SendToBack()
            return container

        # one undo scope so a mid-build failure can't leave an orphan container
        scope = self._guard("starting undo scope", lambda: app.BeginUndoScope("Add container"))
        try:
            container = self._guard(f"creating container {label!r}", _build)
            self._guard("closing undo scope", lambda: app.EndUndoScope(scope, True))
        except Exception:
            try:
                app.EndUndoScope(scope, False)  # False = roll the drop back
            except Exception:
                pass
            raise
        return {**self._shape_info(container), "member_ids": member_ids}

    def container_members(
        self,
        action: str,
        container_id: int,
        member_ids: list[int],
        page_name: Optional[str] = None,
    ) -> dict:
        if action not in ("add", "remove"):
            raise VisioMcpError(f"action must be 'add' or 'remove', got {action!r}")
        if not member_ids:
            raise VisioMcpError("member_ids must not be empty")
        page = self._page(page_name)
        container = self._shape_by_id(page, container_id)
        members = [self._shape_by_id(page, sid) for sid in member_ids]

        def _apply():
            props = container.ContainerProperties
            if props is None:
                raise VisioMcpError(
                    f"Shape {container_id} is not a container — create one with "
                    "add_container."
                )
            for m in members:
                if action == "add":
                    props.AddMember(m, C.VIS_MEMBER_ADD_EXPAND_CONTAINER)
                else:
                    props.RemoveMember(m)

        self._guard(f"{'adding' if action == 'add' else 'removing'} container members", _apply)
        return {"container_id": container_id, "action": action, "member_ids": member_ids}

    # -- layout / introspection ------------------------------------------------

    def auto_layout(
        self,
        style: str = "flowchart_tb",
        spacing_in: float = 0.75,
        resize_page: bool = True,
        page_name: Optional[str] = None,
    ) -> dict:
        if style not in C.LAYOUT_STYLES:
            raise VisioMcpError(
                f"style must be one of {sorted(C.LAYOUT_STYLES)}, got {style!r}"
            )
        page = self._page(page_name)
        place_style, route_style = C.LAYOUT_STYLES[style]

        def _layout():
            sheet = page.PageSheet
            sheet.CellsU("PlaceStyle").ResultIU = place_style
            sheet.CellsU("RouteStyle").ResultIU = route_style
            sheet.CellsU("AvenueSizeX").ResultIU = spacing_in
            sheet.CellsU("AvenueSizeY").ResultIU = spacing_in
            page.Layout()
            if resize_page:
                page.ResizeToFitContents()

        self._guard(f"auto-layout ({style})", _layout)
        return {"page": str(page.Name), "style": style, "note": "Export a PNG to inspect the result."}

    def get_page_state(self, page_name: Optional[str] = None) -> dict:
        page = self._page(page_name)
        shapes = []
        for shape in _items(page.Shapes):
            info = self._shape_info(shape)
            try:
                info["master"] = str(shape.Master.NameU) if shape.Master is not None else None
            except Exception:
                info["master"] = None
            is_connector = False
            try:
                is_connector = bool(int(shape.OneD))
            except Exception:
                pass
            info["is_connector"] = is_connector
            try:
                containers = shape.MemberOfContainers
                info["container_ids"] = [int(c) for c in containers] if containers else []
            except Exception:
                info["container_ids"] = []
            try:
                info["is_container"] = shape.ContainerProperties is not None
            except Exception:
                info["is_container"] = False
            if is_connector:
                from_id = to_id = None
                try:
                    for connect in _items(shape.Connects):
                        cell_name = str(connect.FromCell.Name)
                        target = int(connect.ToSheet.ID)
                        if cell_name.startswith("Begin"):
                            from_id = target
                        elif cell_name.startswith("End"):
                            to_id = target
                except Exception:
                    pass
                info["from_id"] = from_id
                info["to_id"] = to_id
            shapes.append(info)
        sheet = page.PageSheet
        return {
            "page": str(page.Name),
            "page_size_in": [
                round(float(sheet.CellsU("PageWidth").ResultIU), 3),
                round(float(sheet.CellsU("PageHeight").ResultIU), 3),
            ],
            "shapes": shapes,
        }
