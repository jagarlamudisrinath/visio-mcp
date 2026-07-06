"""Pydantic models shared by tools and the Visio client."""

from __future__ import annotations

import re
from typing import Optional

from pydantic import BaseModel, Field, field_validator

_HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


def hex_to_rgb_formula(color: str) -> str:
    """'#4472C4' -> 'RGB(68,114,196)' (the formula Visio cells expect)."""
    v = color.strip()
    if not _HEX_RE.match(v):
        raise ValueError(f"Color must be hex like '#4472C4', got {color!r}")
    return f"RGB({int(v[1:3], 16)},{int(v[3:5], 16)},{int(v[5:7], 16)})"


class DropSpec(BaseModel):
    """One shape to drop, used by drop_shape / drop_shapes."""

    master: str = Field(description="Master (stencil shape) name, e.g. 'Process' or 'Decision'")
    x: float = Field(description="Center X in inches from the page's left edge")
    y: float = Field(description="Center Y in inches from the page's BOTTOM edge (y grows upward)")
    stencil: Optional[str] = Field(
        default=None,
        description="Stencil to look in; omit to search all open stencils",
    )
    text: Optional[str] = Field(default=None, description="Label text for the shape")
    width_in: Optional[float] = Field(default=None, gt=0, description="Override width in inches")
    height_in: Optional[float] = Field(default=None, gt=0, description="Override height in inches")

    @field_validator("master")
    @classmethod
    def _master_not_empty(cls, v: str) -> str:
        if not v.strip():
            raise ValueError("master must not be empty")
        return v
