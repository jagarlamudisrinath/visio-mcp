"""Error types and COM-error translation.

Every message is written for the calling agent: it should say what went
wrong AND what tool call to make next.
"""

from __future__ import annotations

from . import constants as C


class VisioMcpError(Exception):
    """An error with an agent-actionable message."""


def hresult_of(exc: BaseException) -> int | None:
    """Extract the HRESULT from a pywintypes.com_error without importing pywin32."""
    if exc.args and isinstance(exc.args[0], int):
        return exc.args[0]
    return getattr(exc, "hresult", None)


def com_error_detail(exc: BaseException) -> str:
    """Pull Visio's own error text out of a com_error's excepinfo, if present."""
    try:
        excepinfo = exc.args[2]
        if excepinfo and excepinfo[2]:
            return str(excepinfo[2]).strip()
    except (IndexError, TypeError):
        pass
    return str(exc)


def translate_com_error(exc: BaseException, context: str = "") -> VisioMcpError:
    """Map a com_error to an actionable VisioMcpError."""
    hr = hresult_of(exc)
    prefix = f"{context}: " if context else ""
    if hr == C.CO_E_CLASSSTRING:
        return VisioMcpError(
            f"{prefix}Microsoft Visio desktop is not installed on this machine "
            "(COM class 'Visio.Application' not registered). Install Visio desktop "
            "and try again."
        )
    if hr == C.RPC_E_CALL_REJECTED:
        return VisioMcpError(
            f"{prefix}Visio rejected the call because it is busy — most likely a "
            "modal dialog is open in the Visio window. Ask the user to close any "
            "open dialog in Visio, then retry."
        )
    hr_text = f" (HRESULT {hr & 0xFFFFFFFF:#010x})" if hr is not None else ""
    return VisioMcpError(f"{prefix}Visio COM error{hr_text}: {com_error_detail(exc)}")
