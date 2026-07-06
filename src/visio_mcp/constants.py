"""Visio COM constants used by this server.

We ship integer values instead of relying on win32com's makepy-generated
constants (gencache), which avoids gen_py cache corruption on user machines
and lets the non-COM code import these on any platform.

scripts/smoke_test.py re-asserts the load-bearing values against
win32com.client.constants at runtime when a generated typelib is available.
"""

# --- VisOpenSaveArgs (Documents.OpenEx flags) ------------------------------
VIS_OPEN_COPY = 0x1
VIS_OPEN_RO = 0x2
VIS_OPEN_DOCKED = 0x4
VIS_OPEN_DONT_LIST = 0x8
VIS_OPEN_HIDDEN = 0x40
VIS_OPEN_MACROS_DISABLED = 0x80

STENCIL_OPEN_FLAGS = VIS_OPEN_RO | VIS_OPEN_DOCKED  # = 6

# --- VisDocumentTypes ------------------------------------------------------
VIS_TYPE_DRAWING = 1
VIS_TYPE_STENCIL = 2

# --- VisMeasurementSystem --------------------------------------------------
VIS_MS_DEFAULT = 0
VIS_MS_US = 1
VIS_MS_METRIC = 2

# --- VisBuiltInStencilTypes (Application.GetBuiltInStencilFile) ------------
VIS_STENCIL_BACKGROUNDS = 0
VIS_STENCIL_BORDERS = 1
VIS_STENCIL_CONTAINERS = 2
VIS_STENCIL_CALLOUTS = 3
VIS_STENCIL_LEGENDS = 4

# --- RouteStyle / ShapeRouteStyle cell values (visLORoute*) ----------------
ROUTE_DEFAULT = 0
ROUTE_RIGHT_ANGLE = 1
ROUTE_STRAIGHT = 2
ROUTE_ORGCHART_TB = 3
ROUTE_ORGCHART_LR = 4
ROUTE_FLOWCHART_TB = 5
ROUTE_FLOWCHART_LR = 6
ROUTE_TREE_TB = 7
ROUTE_TREE_LR = 8
ROUTE_NETWORK = 9
ROUTE_CENTER_TO_CENTER = 16

# --- PlaceStyle cell values (visPLOPlace*) ---------------------------------
PLACE_DEFAULT = 0
PLACE_TOP_TO_BOTTOM = 1
PLACE_LEFT_TO_RIGHT = 2
PLACE_RADIAL = 3
PLACE_BOTTOM_TO_TOP = 4
PLACE_RIGHT_TO_LEFT = 5
PLACE_CIRCULAR = 6
PLACE_COMPACT_DOWN_RIGHT = 7

# --- ConLineRouteExt (curved vs straight connector rendering) --------------
ROUTE_EXT_DEFAULT = 0
ROUTE_EXT_STRAIGHT = 1
ROUTE_EXT_NURBS = 2  # curved

# --- VisAutoConnectDir -----------------------------------------------------
AUTOCONNECT_NONE = 0

# --- Arrowheads ------------------------------------------------------------
ARROW_NONE = 0
ARROW_FILLED_TRIANGLE = 5

# --- LinePattern cell values ------------------------------------------------
LINE_PATTERNS: dict[str, int] = {
    "solid": 1,
    "dashed": 2,
    "dotted": 3,
    "dash_dot": 4,
}

# --- PrintPageOrientation cell values ----------------------------------------
ORIENTATION_PORTRAIT = 1
ORIENTATION_LANDSCAPE = 2

# --- VisMemberAddOptions (ContainerProperties.AddMember) ---------------------
VIS_MEMBER_ADD_EXPAND_CONTAINER = 1

# --- Paragraph HorzAlign cell values -----------------------------------------
ALIGN_LEFT = 0
ALIGN_CENTER = 1
ALIGN_RIGHT = 2

# --- HRESULTs we branch on -------------------------------------------------
MK_E_UNAVAILABLE = -2147221021  # GetActiveObject: no running instance
CO_E_CLASSSTRING = -2147221005  # invalid class string: Visio not installed
RPC_E_CALL_REJECTED = -2147418111  # Visio busy (modal dialog etc.); retry

# --- Master name aliases -----------------------------------------------------
# Names agents commonly guess -> the master that actually exists in Visio's
# built-in Azure/AWS stencils. Consulted by _find_master when a direct lookup
# misses; keys are lowercase.
MASTER_ALIASES: dict[str, str] = {
    "private endpoint": "Private Link",
    "amazon s3": "Bucket with Objects",
    "s3": "Bucket with Objects",
    "s3 bucket": "Bucket with Objects",
    "iam": "Security Identity and Compliance",
    "iam user": "Role",
    "self-hosted ir": "Virtual Machine",
    "self-hosted integration runtime": "Virtual Machine",
    "azure ad": "Azure Active Directory",
    "entra id": "Azure Active Directory",
}

# Guessed names that are not masters at all -> guidance for the agent.
MASTER_HINTS: dict[str, str] = {
    "subnet": "Subnets are drawn as containers, not masters — use add_container.",
    "vnet": "VNets are drawn as containers, not masters — use add_container.",
    "vpc": "VPCs are drawn as containers, not masters — use add_container.",
    "resource group": "Resource groups are drawn as containers — use add_container.",
    "availability zone": "Availability zones are drawn as containers — use add_container.",
}

# --- auto_layout style -> (PlaceStyle, RouteStyle) -------------------------
LAYOUT_STYLES: dict[str, tuple[int, int]] = {
    "flowchart_tb": (PLACE_TOP_TO_BOTTOM, ROUTE_FLOWCHART_TB),
    "flowchart_lr": (PLACE_LEFT_TO_RIGHT, ROUTE_FLOWCHART_LR),
    "tree_tb": (PLACE_TOP_TO_BOTTOM, ROUTE_TREE_TB),
    "tree_lr": (PLACE_LEFT_TO_RIGHT, ROUTE_TREE_LR),
    "radial": (PLACE_RADIAL, ROUTE_NETWORK),
    "circular": (PLACE_CIRCULAR, ROUTE_CENTER_TO_CENTER),
}

# --- connect_shapes route option -> (ShapeRouteStyle, ConLineRouteExt) -----
CONNECTOR_ROUTES: dict[str, tuple[int, int]] = {
    "right_angle": (ROUTE_RIGHT_ANGLE, ROUTE_EXT_STRAIGHT),
    "straight": (ROUTE_CENTER_TO_CENTER, ROUTE_EXT_STRAIGHT),
    "curved": (ROUTE_RIGHT_ANGLE, ROUTE_EXT_NURBS),
}
