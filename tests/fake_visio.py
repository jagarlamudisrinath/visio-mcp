"""Duck-typed fake of the Visio COM surface used by VisioClient.

Implements exactly the members VisioClient touches: Application (Documents,
ActiveDocument, ActivePage, ActiveWindow, MyShapesPath, Version, undo scopes,
ConnectorToolDataObject), Documents (Add/Open/OpenEx/Count/Item), Document
(Type, Name, FullName, Path, Pages, Masters, Save/SaveAs), Page (Drop, Shapes,
PageSheet, Layout, Export, ...), Shape (ID, Text, CellsU, GlueTo plumbing,
Connects), Cell (ResultIU, FormulaU, GlueTo).

COM collections are 1-based (Count/Item), matching how VisioClient iterates.
"""

from __future__ import annotations

import os

# A tiny valid 1x1 transparent PNG so export tests produce a real image file.
PNG_1X1 = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049"
    "454e44ae426082"
)


class com_error(Exception):  # noqa: N801 — must be named like pywintypes.com_error
    def __init__(self, hresult: int, text: str = "fake com error"):
        super().__init__(hresult, text, (0, None, text, None, 0, hresult))
        self.hresult = hresult


class FakeCollection:
    """1-based COM-style collection."""

    def __init__(self, items=None):
        self._list = list(items or [])

    @property
    def Count(self):
        return len(self._list)

    def Item(self, i):
        return self._list[i - 1]

    def append(self, item):
        self._list.append(item)

    def remove(self, item):
        self._list.remove(item)


class FakeCell:
    def __init__(self, shape, name, value=0.0):
        self.shape = shape
        self.Name = name
        self.ResultIU = value
        self.FormulaU = ""
        self.glued_to = None

    def GlueTo(self, other_cell):
        self.glued_to = other_cell
        self.shape.page.glue_log.append(
            (self.shape.ID, self.Name, other_cell.shape.ID, other_cell.Name)
        )
        self.shape.connects.append(
            FakeConnect(from_cell=self, to_sheet=other_cell.shape)
        )


class FakeConnect:
    def __init__(self, from_cell, to_sheet):
        self.FromCell = from_cell
        self.ToSheet = to_sheet


class FakeShape:
    def __init__(self, page, shape_id, master=None, x=0.0, y=0.0):
        self.page = page
        self.ID = shape_id
        self.Master = master
        self.NameU = f"{master.NameU if master else 'Connector'}.{shape_id}"
        self.Text = ""
        self.OneD = 0
        self.connects: list[FakeConnect] = []
        self._cells: dict[str, FakeCell] = {}
        self.CellsU("PinX").ResultIU = x
        self.CellsU("PinY").ResultIU = y
        self.CellsU("Width").ResultIU = 1.0
        self.CellsU("Height").ResultIU = 0.75

    def CellsU(self, name):
        if name not in self._cells:
            self._cells[name] = FakeCell(self, name)
        return self._cells[name]

    @property
    def Connects(self):
        return FakeCollection(self.connects)

    def Delete(self):
        self.page.shapes.remove(self)


class FakeMaster:
    def __init__(self, name):
        self.Name = name
        self.NameU = name


class FakePageSheet:
    def __init__(self):
        self._cells = {}

    def CellsU(self, name):
        if name not in self._cells:
            cell = FakeCell.__new__(FakeCell)
            cell.Name = name
            cell.ResultIU = 8.5 if name == "PageWidth" else 11.0 if name == "PageHeight" else 0.0
            cell.FormulaU = ""
            self._cells[name] = cell
        return self._cells[name]


class FakePage:
    def __init__(self, doc, name, index):
        self.doc = doc
        self.Name = name
        self.NameU = name
        self.Index = index
        self.shapes: list[FakeShape] = []
        self.glue_log: list[tuple] = []
        self.layout_calls = 0
        self.resize_calls = 0
        self.PageSheet = FakePageSheet()
        self._next_id = 1

    @property
    def Shapes(self):
        coll = FakeCollection(self.shapes)
        page = self

        def item_from_id(sid):
            for s in page.shapes:
                if s.ID == sid:
                    return s
            raise com_error(-2147352567, f"no shape with ID {sid}")

        coll.ItemFromID = item_from_id
        return coll

    def Drop(self, master_or_dataobject, x, y):
        shape = FakeShape(self, self._next_id, master=None, x=x, y=y)
        self._next_id += 1
        if isinstance(master_or_dataobject, FakeMaster):
            shape.Master = master_or_dataobject
            shape.NameU = f"{master_or_dataobject.NameU}.{shape.ID}"
        else:  # ConnectorToolDataObject
            shape.OneD = 1
            shape.Master = None
        self.shapes.append(shape)
        return shape

    def Layout(self):
        self.layout_calls += 1

    def ResizeToFitContents(self):
        self.resize_calls += 1

    def Export(self, path):
        with open(path, "wb") as f:
            f.write(PNG_1X1)


class FakeDocument:
    def __init__(self, app, name, doc_type=1, path=""):
        self.app = app
        self.Type = doc_type
        self.Name = name
        self.Path = path
        self.FullName = os.path.join(path, name) if path else name
        self._pages = [FakePage(self, "Page-1", 1)]
        self.Masters = FakeCollection()
        self.saved_to = None

    @property
    def Pages(self):
        coll = FakeCollection(self._pages)
        doc = self

        def add():
            page = FakePage(doc, f"Page-{len(doc._pages) + 1}", len(doc._pages) + 1)
            doc._pages.append(page)
            return page

        coll.Add = add
        return coll

    def Save(self):
        self.saved_to = self.FullName

    def SaveAs(self, path):
        with open(path, "wb") as f:  # real Visio writes a file; tests rely on it
            f.write(b"PK-fake-vsdx")
        self.Path = os.path.dirname(path)
        self.Name = os.path.basename(path)
        self.FullName = path
        self.saved_to = path


class FakeWindow:
    def __init__(self, app):
        self._app = app

    @property
    def Page(self):
        return self._app.ActivePage

    @Page.setter
    def Page(self, page):
        self._app.ActivePage = page


class FakeDocuments:
    def __init__(self, app):
        self._app = app
        self._coll = FakeCollection()

    @property
    def Count(self):
        return self._coll.Count

    def Item(self, i):
        return self._coll.Item(i)

    def Add(self, template=""):
        if template and template not in self._app.known_templates:
            raise com_error(-2032465756, f"template not found: {template}")
        doc = FakeDocument(self._app, template or f"Drawing{self.Count + 1}", doc_type=1)
        self._coll.append(doc)
        self._app.ActiveDocument = doc
        self._app.ActivePage = doc._pages[0]
        return doc

    def Open(self, path):
        if not os.path.exists(path):
            raise com_error(-2032465756, f"file not found: {path}")
        doc = FakeDocument(self._app, os.path.basename(path), doc_type=1, path=os.path.dirname(path))
        self._coll.append(doc)
        self._app.ActiveDocument = doc
        self._app.ActivePage = doc._pages[0]
        return doc

    def OpenEx(self, name, flags):
        stencil_masters = self._app.known_stencils.get(name) or self._app.known_stencils.get(
            os.path.basename(name)
        )
        if stencil_masters is None:
            raise com_error(-2032465756, f"stencil not found: {name}")
        doc = FakeDocument(self._app, os.path.basename(name), doc_type=2, path=os.path.dirname(name))
        for master_name in stencil_masters:
            doc.Masters.append(FakeMaster(master_name))
        self._coll.append(doc)
        # stencils open docked: they do NOT become the active document
        return doc


class FakeApplication:
    def __init__(self, my_shapes_path="/tmp/MyShapes"):
        self.Version = "16.0 (fake)"
        self.Visible = False
        self.AlertResponse = 0
        self.MyShapesPath = my_shapes_path
        self.Documents = FakeDocuments(self)
        self.ActiveDocument = None
        self.ActivePage = None
        self.ActiveWindow = FakeWindow(self)
        self.ConnectorToolDataObject = object()
        self.undo_scopes: list[tuple[int, bool]] = []
        self._next_scope = 1
        # what Documents.Add / OpenEx will accept:
        self.known_templates = {"BASFLO_U.VSTX", "BASFLO_M.VSTX"}
        self.known_stencils = {
            "BASFLO_U.VSSX": ["Process", "Decision", "Start/End", "Document", "Data"],
        }

    def BeginUndoScope(self, name):
        scope = self._next_scope
        self._next_scope += 1
        return scope

    def EndUndoScope(self, scope, commit):
        self.undo_scopes.append((scope, bool(commit)))
