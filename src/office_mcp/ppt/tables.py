"""TableMixin — table lifecycle, cell formatting, sizing, merge, row/col insert.

Mixed into PowerPointController. Relies on base for:
    self._validate_index(), self._require_shape(), self._parse_color(), self.pres

PowerPoint table model:
  Shape (Type=msoTable)
    └─ .Table
         ├─ .Rows.Count / .Columns.Count
         ├─ .Cell(row, col)
         │     ├─ .Shape  (the cell as a shape — has TextFrame, Fill, …)
         │     └─ .Merge(MergeTo=other_cell)
         ├─ .Rows(r).Height
         ├─ .Columns(c).Width
         ├─ .Rows.Add(BeforeRow=n)
         └─ .Columns.Add(BeforeColumn=n)
"""

from __future__ import annotations

import re

import pywintypes

from .constants import (
    PARAGRAPH_ALIGN_MAP,
    TABLE_STYLE_MAP,
    VERTICAL_ALIGN_MAP,
)
from .errors import PowerPointError

_GUID_RE = re.compile(r"^\{[0-9A-Fa-f-]{36}\}$")


class TableMixin:
    def _require_table(self, slide_index: int, shape_id: int):
        """Return (slide, shape, table) — raises if shape is not a table."""
        slide, shape = self._require_shape(slide_index, shape_id)
        try:
            table = shape.Table
        except pywintypes.com_error as e:
            raise PowerPointError(f"Shape id={shape_id} on slide {slide_index} is not a table: {e}")
        return slide, shape, table

    @staticmethod
    def _validate_cell(table, row: int, col: int) -> None:
        rows = int(table.Rows.Count)
        cols = int(table.Columns.Count)
        if row < 1 or row > rows or col < 1 or col > cols:
            raise PowerPointError(f"Cell ({row},{col}) out of range — table is {rows}x{cols}.")

    def table_add(
        self,
        slide_index: int,
        rows: int,
        cols: int,
        left: float,
        top: float,
        width: float,
        height: float,
    ) -> dict:
        """Add a new table to a slide."""
        self._validate_index(slide_index)
        if rows < 1 or cols < 1:
            raise PowerPointError(f"rows and cols must be ≥ 1, got rows={rows}, cols={cols}")
        slide = self.pres.Slides(slide_index)
        try:
            shape = slide.Shapes.AddTable(
                int(rows),
                int(cols),
                float(left),
                float(top),
                float(width),
                float(height),
            )
        except pywintypes.com_error as e:
            raise PowerPointError(f"AddTable failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "shape_id": int(shape.Id),
            "name": shape.Name,
            "rows": int(rows),
            "cols": int(cols),
            "left": float(shape.Left),
            "top": float(shape.Top),
            "width": float(shape.Width),
            "height": float(shape.Height),
        }

    def table_set_cell_text(
        self,
        slide_index: int,
        shape_id: int,
        row: int,
        col: int,
        text: str,
    ) -> dict:
        """Overwrite a single cell's text."""
        _, _, table = self._require_table(slide_index, shape_id)
        self._validate_cell(table, row, col)
        try:
            cell = table.Cell(int(row), int(col))
            cell.Shape.TextFrame.TextRange.Text = text
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set cell text failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "row": int(row),
            "col": int(col),
            "text_length": len(text),
        }

    def table_set_cell_fill(
        self,
        slide_index: int,
        shape_id: int,
        row: int,
        col: int,
        color,
    ) -> dict:
        """Set a single cell's solid fill."""
        _, _, table = self._require_table(slide_index, shape_id)
        self._validate_cell(table, row, col)
        rgb = self._parse_color(color)
        try:
            cell_shape = table.Cell(int(row), int(col)).Shape
            cell_shape.Fill.Solid()
            cell_shape.Fill.ForeColor.RGB = rgb
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set cell fill failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "row": int(row),
            "col": int(col),
            "rgb": rgb,
        }

    def table_set_column_width(
        self,
        slide_index: int,
        shape_id: int,
        col: int,
        width: float,
    ) -> dict:
        """Set a column width in points."""
        _, _, table = self._require_table(slide_index, shape_id)
        cols = int(table.Columns.Count)
        if col < 1 or col > cols:
            raise PowerPointError(f"col {col} out of range (1-{cols})")
        if float(width) <= 0:
            raise PowerPointError(f"width must be > 0, got {width}")
        try:
            table.Columns(int(col)).Width = float(width)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set column width failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "col": int(col),
            "width": float(width),
        }

    def table_set_row_height(
        self,
        slide_index: int,
        shape_id: int,
        row: int,
        height: float,
    ) -> dict:
        """Set a row height in points."""
        _, _, table = self._require_table(slide_index, shape_id)
        rows = int(table.Rows.Count)
        if row < 1 or row > rows:
            raise PowerPointError(f"row {row} out of range (1-{rows})")
        if float(height) <= 0:
            raise PowerPointError(f"height must be > 0, got {height}")
        try:
            table.Rows(int(row)).Height = float(height)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set row height failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "row": int(row),
            "height": float(height),
        }

    def table_merge_cells(
        self,
        slide_index: int,
        shape_id: int,
        row1: int,
        col1: int,
        row2: int,
        col2: int,
    ) -> dict:
        """Merge a rectangular range of cells."""
        _, _, table = self._require_table(slide_index, shape_id)
        self._validate_cell(table, row1, col1)
        self._validate_cell(table, row2, col2)
        try:
            a = table.Cell(int(row1), int(col1))
            b = table.Cell(int(row2), int(col2))
            a.Merge(MergeTo=b)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Merge failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "from": (int(row1), int(col1)),
            "to": (int(row2), int(col2)),
        }

    def table_add_row(
        self,
        slide_index: int,
        shape_id: int,
        position: int | None = None,
    ) -> dict:
        """Add a row. ``position`` is the 1-based index of an existing row
        BEFORE which the new row is inserted. Omit to append."""
        _, _, table = self._require_table(slide_index, shape_id)
        rows = int(table.Rows.Count)
        try:
            if position is None:
                table.Rows.Add()
            else:
                if position < 1 or position > rows + 1:
                    raise PowerPointError(f"position {position} out of range (1-{rows + 1})")
                # BeforeRow=n inserts before that row. Appending requires no arg.
                if position == rows + 1:
                    table.Rows.Add()
                else:
                    table.Rows.Add(BeforeRow=int(position))
        except pywintypes.com_error as e:
            raise PowerPointError(f"Add row failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "rows": int(table.Rows.Count),
            "added_at": position or (rows + 1),
        }

    def table_add_column(
        self,
        slide_index: int,
        shape_id: int,
        position: int | None = None,
    ) -> dict:
        """Add a column. ``position`` is the 1-based index BEFORE which the
        new column is inserted. Omit to append."""
        _, _, table = self._require_table(slide_index, shape_id)
        cols = int(table.Columns.Count)
        try:
            if position is None:
                table.Columns.Add()
            else:
                if position < 1 or position > cols + 1:
                    raise PowerPointError(f"position {position} out of range (1-{cols + 1})")
                if position == cols + 1:
                    table.Columns.Add()
                else:
                    table.Columns.Add(BeforeColumn=int(position))
        except pywintypes.com_error as e:
            raise PowerPointError(f"Add column failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "cols": int(table.Columns.Count),
            "added_at": position or (cols + 1),
        }

    def table_apply_style(
        self,
        slide_index: int,
        shape_id: int,
        style: str,
    ) -> dict:
        """Apply a built-in table style. ``style`` accepts a friendly name
        (e.g. ``"medium_style_2_accent_1"``) or a raw GUID string. This
        replaces dozens of per-cell formatting ops with a single call —
        banded rows, branded header, theme borders all come for free."""
        _, shape, table = self._require_table(slide_index, shape_id)
        if _GUID_RE.match(style):
            guid = style
        else:
            key = style.lower().strip()
            if key not in TABLE_STYLE_MAP:
                raise PowerPointError(
                    f"Unknown table style '{style}'. Use a GUID or one of {sorted(TABLE_STYLE_MAP)}"
                )
            guid = TABLE_STYLE_MAP[key]
        try:
            table.ApplyStyle(guid, True)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Apply table style failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "style": style,
            "style_id": guid,
        }

    def table_set_cell_font(
        self,
        slide_index: int,
        shape_id: int,
        row: int,
        col: int,
        name: str | None = None,
        size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color=None,
    ) -> dict:
        """Set font attributes for a single cell. Only the kwargs you pass
        are applied — everything else is left untouched."""
        _, _, table = self._require_table(slide_index, shape_id)
        self._validate_cell(table, row, col)
        applied: dict = {}
        try:
            rng = table.Cell(int(row), int(col)).Shape.TextFrame.TextRange
            font = rng.Font
            if name is not None:
                font.Name = str(name)
                applied["name"] = str(name)
            if size is not None:
                if float(size) <= 0:
                    raise PowerPointError(f"size must be > 0, got {size}")
                font.Size = float(size)
                applied["size"] = float(size)
            if bold is not None:
                font.Bold = -1 if bold else 0
                applied["bold"] = bool(bold)
            if italic is not None:
                font.Italic = -1 if italic else 0
                applied["italic"] = bool(italic)
            if underline is not None:
                font.Underline = -1 if underline else 0
                applied["underline"] = bool(underline)
            if color is not None:
                rgb = self._parse_color(color)
                font.Color.RGB = rgb
                applied["rgb"] = rgb
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set cell font failed: {e}")
        if not applied:
            raise PowerPointError(
                "table_set_cell_font called with no attributes. Pass at "
                "least one of name/size/bold/italic/underline/color."
            )
        return {
            "success": True,
            "shape_id": int(shape_id),
            "row": int(row),
            "col": int(col),
            "applied": applied,
        }

    def table_set_cell_alignment(
        self,
        slide_index: int,
        shape_id: int,
        row: int,
        col: int,
        horizontal: str | None = None,
        vertical: str | None = None,
    ) -> dict:
        """Align text inside a single cell. ``horizontal`` ∈
        {left, center, right, justify}; ``vertical`` ∈ {top, middle, bottom}.

        Critical for readable financial tables — numbers right-aligned,
        headers centered, labels left-aligned."""
        _, _, table = self._require_table(slide_index, shape_id)
        self._validate_cell(table, row, col)
        applied: dict = {}
        try:
            cell_shape = table.Cell(int(row), int(col)).Shape
            tf = cell_shape.TextFrame
            if horizontal is not None:
                key = horizontal.lower().strip()
                if key not in PARAGRAPH_ALIGN_MAP:
                    raise PowerPointError(
                        f"Unknown horizontal alignment '{horizontal}'. "
                        f"Choose from {sorted(PARAGRAPH_ALIGN_MAP)}"
                    )
                tf.TextRange.ParagraphFormat.Alignment = PARAGRAPH_ALIGN_MAP[key]
                applied["horizontal"] = key
            if vertical is not None:
                key = vertical.lower().strip()
                if key not in VERTICAL_ALIGN_MAP:
                    raise PowerPointError(
                        f"Unknown vertical alignment '{vertical}'. "
                        f"Choose from {sorted(VERTICAL_ALIGN_MAP)}"
                    )
                tf.VerticalAnchor = VERTICAL_ALIGN_MAP[key]
                applied["vertical"] = key
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set cell alignment failed: {e}")
        if not applied:
            raise PowerPointError(
                "table_set_cell_alignment called with no attributes. Pass "
                "horizontal and/or vertical."
            )
        return {
            "success": True,
            "shape_id": int(shape_id),
            "row": int(row),
            "col": int(col),
            "applied": applied,
        }

    def table_get_info(self, slide_index: int, shape_id: int) -> dict:
        """Inspect a table — dimensions + per-cell text."""
        _, _, table = self._require_table(slide_index, shape_id)
        rows = int(table.Rows.Count)
        cols = int(table.Columns.Count)
        grid: list[list[str]] = []
        for r in range(1, rows + 1):
            row_data: list[str] = []
            for c in range(1, cols + 1):
                try:
                    cell = table.Cell(r, c)
                    txt = cell.Shape.TextFrame.TextRange.Text or ""
                except pywintypes.com_error:
                    txt = ""
                row_data.append(txt)
            grid.append(row_data)
        return {
            "shape_id": int(shape_id),
            "rows": rows,
            "cols": cols,
            "column_widths": [float(table.Columns(c).Width) for c in range(1, cols + 1)],
            "row_heights": [float(table.Rows(r).Height) for r in range(1, rows + 1)],
            "cells": grid,
        }
