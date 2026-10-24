"""Helpers for comparing worksheets by value and by serialized cell style."""

from openpyxl.worksheet.worksheet import Worksheet

STYLE_ATTRS = ("number_format", "font", "fill", "border", "alignment", "protection")


def cell_signature(cell) -> tuple:
    return (cell.value,) + tuple(repr(getattr(cell, a)) for a in STYLE_ATTRS)


def style_signature(cell) -> tuple:
    return tuple(repr(getattr(cell, a)) for a in STYLE_ATTRS)


def sheet_signature(ws: Worksheet, max_row: int | None = None, max_col: int = 15) -> dict:
    limit = max_row or ws.max_row
    cells = {
        c.coordinate: cell_signature(c)
        for row in ws.iter_rows(max_row=limit, max_col=max_col)
        for c in row
    }
    merges = sorted(str(m) for m in ws.merged_cells.ranges if m.min_row <= limit)
    widths = {k: v.width for k, v in ws.column_dimensions.items()}
    heights = {k: v.height for k, v in ws.row_dimensions.items() if v.height and k <= limit}
    return {"cells": cells, "merges": merges, "widths": widths, "heights": heights}


def first_difference(a: dict, b: dict):
    """Describe the first differing cell between two sheet signatures (for assertion messages)."""
    for k in sorted(set(a["cells"]) | set(b["cells"])):
        if a["cells"].get(k) != b["cells"].get(k):
            return k, a["cells"].get(k), b["cells"].get(k)
    return None
