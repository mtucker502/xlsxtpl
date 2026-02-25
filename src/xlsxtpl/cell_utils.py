from __future__ import annotations

import re
from datetime import date, datetime, time
from typing import Any

from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

# Matches a cell whose entire value is a single {{ expression }}
# Uses negative lookahead to prevent matching across multiple {{ }} pairs
_PURE_EXPR_RE = re.compile(r"^\s*\{\{\s*((?:(?!\}\}).)+?)\s*\}\}\s*$", re.DOTALL)

# Matches any {{ ... }} or {% ... %} tag
_HAS_TAG_RE = re.compile(r"\{\{.*?\}\}|\{%.*?%\}", re.DOTALL)

# Matches a cell whose entire value is a single {% block directive %}
_BLOCK_TAG_RE = re.compile(r"^\s*\{%[-\s]*(.*?)[-\s]*%\}\s*$", re.DOTALL)

# Parses block directives: for, endfor, if, endif
_FOR_RE = re.compile(r"^for\s+(\w+)\s+in\s+(.+)$")
_ENDFOR_RE = re.compile(r"^endfor$")
_IF_RE = re.compile(r"^if\s+(.+)$")
_ENDIF_RE = re.compile(r"^endif$")


def copy_cell(source: Cell, target: Cell) -> None:
    """Copy value and style from source cell to target cell."""
    target.value = source.value
    if source.has_style:
        target._style = source._style


def copy_row_dimensions(ws: Worksheet, src_row: int, tgt_row: int) -> None:
    """Copy row height, hidden, and outline level from one row to another."""
    src_dim = ws.row_dimensions.get(src_row)
    if src_dim is None:
        return
    tgt_dim = ws.row_dimensions[tgt_row]
    tgt_dim.height = src_dim.height
    tgt_dim.hidden = src_dim.hidden
    tgt_dim.outlineLevel = src_dim.outlineLevel


def cell_has_only_expression(value: str) -> re.Match | None:
    """Return match if cell value is purely {{ expr }}, else None."""
    if not isinstance(value, str):
        return None
    return _PURE_EXPR_RE.match(value)


def has_template_tag(value: Any) -> bool:
    """Return True if the value contains any {{ }} or {% %} tag."""
    if not isinstance(value, str):
        return False
    return bool(_HAS_TAG_RE.search(value))


def is_block_tag(value: Any) -> bool:
    """Return True if the cell value is a block directive ({% ... %})."""
    if not isinstance(value, str):
        return False
    return bool(_BLOCK_TAG_RE.match(value))


def extract_block_directive(value: str) -> dict[str, Any] | None:
    """Parse a block directive and return its components.

    Returns a dict with 'type' key ('for', 'endfor', 'if', 'endif') and
    additional keys depending on type, or None if not a block directive.
    """
    m = _BLOCK_TAG_RE.match(value)
    if not m:
        return None
    inner = m.group(1).strip()

    fm = _FOR_RE.match(inner)
    if fm:
        return {"type": "for", "var": fm.group(1), "iterable": fm.group(2).strip()}

    if _ENDFOR_RE.match(inner):
        return {"type": "endfor"}

    im = _IF_RE.match(inner)
    if im:
        return {"type": "if", "condition": im.group(1).strip()}

    if _ENDIF_RE.match(inner):
        return {"type": "endif"}

    return None


def coerce_type(rendered: str, original_value: Any) -> Any:
    """Attempt to convert a rendered string back to the original cell's data type."""
    if not isinstance(rendered, str) or not rendered:
        return rendered

    # Try int
    if isinstance(original_value, (int, float)):
        try:
            as_float = float(rendered)
            as_int = int(as_float)
            if as_float == as_int and "." not in rendered:
                return as_int
            return as_float
        except (ValueError, OverflowError):
            pass

    # Try bool
    if isinstance(original_value, bool):
        lower = rendered.strip().lower()
        if lower in ("true", "1"):
            return True
        if lower in ("false", "0"):
            return False

    # Try date/datetime
    if isinstance(original_value, (date, datetime)):
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(rendered, fmt)
            except ValueError:
                continue

    return rendered
