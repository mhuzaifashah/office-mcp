"""ShapeMixin — autoshape lifecycle, fill, line, rotation, z-order, group, align.

Mixed into PowerPointController. Relies on the following from the base:
    self._validate_index(), self._find_shape_by_id(), self._require_shape(),
    self._parse_color(), self.pres
"""

from __future__ import annotations

import pywintypes

from .constants import (
    ALIGN_MAP,
    AUTOSHAPE_NAME_MAP,
    DASH_STYLE_MAP,
    DISTRIBUTE_MAP,
    SHAPE_TYPE_NAMES,
    TEXT_ORIENTATION_MAP,
    Z_ORDER_MAP,
    msoFalse,
    msoTextOrientationHorizontal,
    msoTrue,
)
from .errors import PowerPointError


class ShapeMixin:
    """Stage 2a (lifecycle) + Stage 2b (fill/line/rotate/group/align/z-order)."""

    # ─────────────────────────────────────────────
    # Lifecycle — list / add / delete / move / resize
    # ─────────────────────────────────────────────

    def _resolve_autoshape_type(self, shape_type: int | str) -> int:
        """Map a friendly name or int to a MsoAutoShapeType value."""
        if isinstance(shape_type, int):
            return shape_type
        key = shape_type.lower().strip().replace(" ", "_")
        if key not in AUTOSHAPE_NAME_MAP:
            raise PowerPointError(
                f"Unknown shape_type '{shape_type}'. Friendly names: {sorted(AUTOSHAPE_NAME_MAP)}"
            )
        return AUTOSHAPE_NAME_MAP[key]

    def shape_list(self, slide_index: int) -> dict:
        """List every shape on a slide with its bbox + type."""
        self._validate_index(slide_index)
        s = self.pres.Slides(slide_index)
        shapes: list[dict] = []
        for i in range(1, s.Shapes.Count + 1):
            try:
                shapes.append(self._shape_summary(s.Shapes(i)))
            except pywintypes.com_error:
                continue
        return {
            "slide_index": slide_index,
            "shapes_count": s.Shapes.Count,
            "shapes": shapes,
        }

    def shape_add(
        self,
        slide_index: int,
        shape_type: int | str,
        left: float,
        top: float,
        width: float,
        height: float,
        text: str | None = None,
    ) -> dict:
        """Add an autoshape to a slide. Coordinates/sizes are in points."""
        self._validate_index(slide_index)
        mso_type = self._resolve_autoshape_type(shape_type)
        slide = self.pres.Slides(slide_index)
        try:
            shape = slide.Shapes.AddShape(
                mso_type,
                float(left),
                float(top),
                float(width),
                float(height),
            )
        except pywintypes.com_error as e:
            raise PowerPointError(f"AddShape failed (type={shape_type}): {e}")
        if text:
            try:
                shape.TextFrame.TextRange.Text = text
            except pywintypes.com_error as e:
                raise PowerPointError(f"AddShape succeeded but setting text failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "shape_id": int(shape.Id),
            "name": shape.Name,
            "shape_type": int(shape.Type),
            "shape_type_name": SHAPE_TYPE_NAMES.get(int(shape.Type), "unknown"),
            "left": float(shape.Left),
            "top": float(shape.Top),
            "width": float(shape.Width),
            "height": float(shape.Height),
        }

    def shape_add_text_box(
        self,
        slide_index: int,
        left: float,
        top: float,
        width: float,
        height: float,
        text: str | None = None,
        orientation: str | None = None,
    ) -> dict:
        """Add a text box to a slide. ``orientation`` defaults to horizontal."""
        self._validate_index(slide_index)
        if orientation:
            key = orientation.lower().strip()
            if key not in TEXT_ORIENTATION_MAP:
                raise PowerPointError(
                    f"Unknown orientation '{orientation}'. "
                    f"Choose from {sorted(TEXT_ORIENTATION_MAP)}"
                )
            orient = TEXT_ORIENTATION_MAP[key]
        else:
            orient = msoTextOrientationHorizontal
        slide = self.pres.Slides(slide_index)
        try:
            shape = slide.Shapes.AddTextbox(
                orient,
                float(left),
                float(top),
                float(width),
                float(height),
            )
        except pywintypes.com_error as e:
            raise PowerPointError(f"AddTextbox failed: {e}")
        if text:
            try:
                shape.TextFrame.TextRange.Text = text
            except pywintypes.com_error as e:
                raise PowerPointError(f"Text box created but setting text failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "shape_id": int(shape.Id),
            "name": shape.Name,
            "shape_type": int(shape.Type),
            "shape_type_name": SHAPE_TYPE_NAMES.get(int(shape.Type), "unknown"),
            "left": float(shape.Left),
            "top": float(shape.Top),
            "width": float(shape.Width),
            "height": float(shape.Height),
            "text": text or "",
        }

    def shape_delete(self, slide_index: int, shape_id: int) -> dict:
        """Delete a shape by its Id."""
        self._validate_index(slide_index)
        slide = self.pres.Slides(slide_index)
        shape = self._find_shape_by_id(slide, shape_id)
        if shape is None:
            raise PowerPointError(f"Shape id={shape_id} not found on slide {slide_index}.")
        name = shape.Name
        try:
            shape.Delete()
        except pywintypes.com_error as e:
            raise PowerPointError(f"Delete failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "deleted_shape_id": int(shape_id),
            "deleted_name": name,
            "remaining_count": slide.Shapes.Count,
        }

    def shape_move(
        self,
        slide_index: int,
        shape_id: int,
        left: float,
        top: float,
    ) -> dict:
        """Move a shape to (left, top) in points (absolute, from slide origin)."""
        _, shape = self._require_shape(slide_index, shape_id)
        try:
            shape.Left = float(left)
            shape.Top = float(top)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Move failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "shape_id": int(shape_id),
            "left": float(shape.Left),
            "top": float(shape.Top),
        }

    def shape_resize(
        self,
        slide_index: int,
        shape_id: int,
        width: float,
        height: float,
    ) -> dict:
        """Resize a shape to (width, height) in points."""
        _, shape = self._require_shape(slide_index, shape_id)
        if float(width) <= 0 or float(height) <= 0:
            raise PowerPointError(
                f"width/height must be positive, got width={width}, height={height}"
            )
        try:
            shape.Width = float(width)
            shape.Height = float(height)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Resize failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "shape_id": int(shape_id),
            "width": float(shape.Width),
            "height": float(shape.Height),
        }

    # ─────────────────────────────────────────────
    # Fill
    # ─────────────────────────────────────────────

    def shape_set_fill_color(
        self,
        slide_index: int,
        shape_id: int,
        color,
    ) -> dict:
        """Set a solid fill color on a shape."""
        _, shape = self._require_shape(slide_index, shape_id)
        rgb = self._parse_color(color)
        try:
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = rgb
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set fill color failed: {e}")
        return {"success": True, "shape_id": int(shape_id), "rgb": rgb}

    def shape_set_fill_transparency(
        self,
        slide_index: int,
        shape_id: int,
        transparency: float,
    ) -> dict:
        """Set fill transparency (0.0 = opaque, 1.0 = fully transparent)."""
        _, shape = self._require_shape(slide_index, shape_id)
        t = float(transparency)
        if t < 0.0 or t > 1.0:
            raise PowerPointError(f"transparency must be in [0.0, 1.0], got {transparency}")
        try:
            shape.Fill.Transparency = t
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set transparency failed: {e}")
        return {"success": True, "shape_id": int(shape_id), "transparency": t}

    def shape_remove_fill(self, slide_index: int, shape_id: int) -> dict:
        """Remove the fill (transparent / no fill)."""
        _, shape = self._require_shape(slide_index, shape_id)
        try:
            shape.Fill.Visible = msoFalse
        except pywintypes.com_error as e:
            raise PowerPointError(f"Remove fill failed: {e}")
        return {"success": True, "shape_id": int(shape_id)}

    # ─────────────────────────────────────────────
    # Line
    # ─────────────────────────────────────────────

    def shape_set_line_color(
        self,
        slide_index: int,
        shape_id: int,
        color,
    ) -> dict:
        """Set the outline color."""
        _, shape = self._require_shape(slide_index, shape_id)
        rgb = self._parse_color(color)
        try:
            shape.Line.Visible = msoTrue
            shape.Line.ForeColor.RGB = rgb
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set line color failed: {e}")
        return {"success": True, "shape_id": int(shape_id), "rgb": rgb}

    def shape_set_line_weight(
        self,
        slide_index: int,
        shape_id: int,
        weight: float,
    ) -> dict:
        """Set the outline thickness in points."""
        _, shape = self._require_shape(slide_index, shape_id)
        if float(weight) <= 0:
            raise PowerPointError(f"weight must be > 0, got {weight}")
        try:
            shape.Line.Visible = msoTrue
            shape.Line.Weight = float(weight)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set line weight failed: {e}")
        return {"success": True, "shape_id": int(shape_id), "weight": float(weight)}

    def shape_set_line_dash_style(
        self,
        slide_index: int,
        shape_id: int,
        dash_style: int | str,
    ) -> dict:
        """Set the outline dash style. Accepts an MsoLineDashStyle int or
        a friendly name (solid, dash, dot, dash_dot, long_dash, ...)."""
        _, shape = self._require_shape(slide_index, shape_id)
        if isinstance(dash_style, int):
            val = dash_style
        else:
            key = dash_style.lower().strip().replace(" ", "_")
            if key not in DASH_STYLE_MAP:
                raise PowerPointError(
                    f"Unknown dash_style '{dash_style}'. Choose from {sorted(DASH_STYLE_MAP)}"
                )
            val = DASH_STYLE_MAP[key]
        try:
            shape.Line.Visible = msoTrue
            shape.Line.DashStyle = val
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set line dash failed: {e}")
        return {"success": True, "shape_id": int(shape_id), "dash_style": val}

    # ─────────────────────────────────────────────
    # Rotation / z-order
    # ─────────────────────────────────────────────

    def shape_set_rotation(
        self,
        slide_index: int,
        shape_id: int,
        rotation: float,
    ) -> dict:
        """Set absolute rotation in degrees (clockwise)."""
        _, shape = self._require_shape(slide_index, shape_id)
        try:
            shape.Rotation = float(rotation)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set rotation failed: {e}")
        return {"success": True, "shape_id": int(shape_id), "rotation": float(shape.Rotation)}

    def shape_z_order(
        self,
        slide_index: int,
        shape_id: int,
        cmd: int | str,
    ) -> dict:
        """Adjust shape stacking. Accepts MsoZOrderCmd int or friendly name
        ('bring_to_front', 'send_to_back', 'bring_forward', 'send_backward')."""
        _, shape = self._require_shape(slide_index, shape_id)
        if isinstance(cmd, int):
            val = cmd
        else:
            key = cmd.lower().strip().replace(" ", "_")
            if key not in Z_ORDER_MAP:
                raise PowerPointError(
                    f"Unknown z_order cmd '{cmd}'. Choose from {sorted(Z_ORDER_MAP)}"
                )
            val = Z_ORDER_MAP[key]
        try:
            shape.ZOrder(val)
        except pywintypes.com_error as e:
            raise PowerPointError(f"ZOrder failed: {e}")
        return {"success": True, "shape_id": int(shape_id), "z_order_cmd": val}

    # ─────────────────────────────────────────────
    # Group / align / distribute
    # ─────────────────────────────────────────────

    def _shape_range(self, slide, shape_ids: list[int]):
        """Build a ShapeRange from a list of shape Ids (resolved to Names)."""
        names: list[str] = []
        missing: list[int] = []
        for sid in shape_ids:
            sh = self._find_shape_by_id(slide, sid)
            if sh is None:
                missing.append(int(sid))
            else:
                names.append(sh.Name)
        if missing:
            raise PowerPointError(f"Shape ids not found on slide: {missing}")
        if not names:
            raise PowerPointError("shape_ids must be non-empty")
        return slide.Shapes.Range(names)

    def shape_group(self, slide_index: int, shape_ids: list[int]) -> dict:
        """Group multiple shapes into one."""
        self._validate_index(slide_index)
        if not shape_ids or len(shape_ids) < 2:
            raise PowerPointError("shape_group requires at least 2 shape_ids.")
        slide = self.pres.Slides(slide_index)
        rng = self._shape_range(slide, shape_ids)
        try:
            grouped = rng.Group()
        except pywintypes.com_error as e:
            raise PowerPointError(f"Group failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "group_shape_id": int(grouped.Id),
            "group_name": grouped.Name,
            "child_count": len(shape_ids),
        }

    def shape_ungroup(self, slide_index: int, shape_id: int) -> dict:
        """Ungroup a previously grouped shape. Returns the new child ids."""
        _, shape = self._require_shape(slide_index, shape_id)
        try:
            children = shape.Ungroup()
        except pywintypes.com_error as e:
            raise PowerPointError(f"Ungroup failed: {e}")
        child_ids: list[int] = []
        try:
            for i in range(1, children.Count + 1):
                child_ids.append(int(children(i).Id))
        except pywintypes.com_error:
            pass
        return {
            "success": True,
            "slide_index": slide_index,
            "ungrouped_id": int(shape_id),
            "child_ids": child_ids,
        }

    def shape_align(
        self,
        slide_index: int,
        shape_ids: list[int],
        alignment: int | str,
        relative_to_slide: bool = False,
    ) -> dict:
        """Align a set of shapes. ``alignment`` accepts MsoAlignCmd int or a
        friendly name (left, center, right, top, middle, bottom).
        ``relative_to_slide`` aligns each shape to the slide instead of to the
        other selected shapes."""
        self._validate_index(slide_index)
        if not shape_ids:
            raise PowerPointError("shape_align requires at least 1 shape_id.")
        if isinstance(alignment, int):
            cmd = alignment
        else:
            key = alignment.lower().strip()
            if key not in ALIGN_MAP:
                raise PowerPointError(
                    f"Unknown alignment '{alignment}'. Choose from {sorted(ALIGN_MAP)}"
                )
            cmd = ALIGN_MAP[key]
        slide = self.pres.Slides(slide_index)
        rng = self._shape_range(slide, shape_ids)
        try:
            rng.Align(cmd, msoTrue if relative_to_slide else msoFalse)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Align failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "shape_ids": [int(s) for s in shape_ids],
            "alignment": cmd,
            "relative_to_slide": bool(relative_to_slide),
        }

    def shape_distribute(
        self,
        slide_index: int,
        shape_ids: list[int],
        direction: int | str,
        relative_to_slide: bool = False,
    ) -> dict:
        """Distribute shapes evenly. ``direction`` is horizontal / vertical."""
        self._validate_index(slide_index)
        if not shape_ids or len(shape_ids) < 2:
            raise PowerPointError("shape_distribute requires at least 2 shape_ids.")
        if isinstance(direction, int):
            cmd = direction
        else:
            key = direction.lower().strip()
            if key not in DISTRIBUTE_MAP:
                raise PowerPointError(
                    f"Unknown direction '{direction}'. Choose from {sorted(DISTRIBUTE_MAP)}"
                )
            cmd = DISTRIBUTE_MAP[key]
        slide = self.pres.Slides(slide_index)
        rng = self._shape_range(slide, shape_ids)
        try:
            rng.Distribute(cmd, msoTrue if relative_to_slide else msoFalse)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Distribute failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "shape_ids": [int(s) for s in shape_ids],
            "direction": cmd,
            "relative_to_slide": bool(relative_to_slide),
        }

    # ─────────────────────────────────────────────────────────
    # Polish — format copy + shadow. The format-copy op lets the
    # agent style one shape perfectly then replicate to many,
    # cutting the per-shape op cost from ~5 calls to 1.
    # ─────────────────────────────────────────────────────────

    def shape_copy_format(
        self,
        slide_index: int,
        source_id: int,
        target_id: int | list[int],
    ) -> dict:
        """Replicate fill/line/font/effects from one shape onto one or more
        targets. Uses PowerPoint's PickUp/Apply pair (the COM equivalent of
        the Format Painter tool). All shapes must be on the same slide."""
        _, source = self._require_shape(slide_index, source_id)
        if isinstance(target_id, int):
            targets_in = [target_id]
        else:
            targets_in = list(target_id)
        if not targets_in:
            raise PowerPointError("shape_copy_format needs at least one target")
        applied: list[int] = []
        try:
            source.PickUp()
            for tid in targets_in:
                _, tshape = self._require_shape(slide_index, tid)
                tshape.Apply()
                applied.append(int(tid))
        except pywintypes.com_error as e:
            raise PowerPointError(f"Copy format failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "source_id": int(source_id),
            "applied_to": applied,
        }

    def shape_set_shadow(
        self,
        slide_index: int,
        shape_id: int,
        visible: bool | None = None,
        color=None,
        transparency: float | None = None,
        blur: float | None = None,
        offset_x: float | None = None,
        offset_y: float | None = None,
    ) -> dict:
        """Configure a shape's drop shadow. Soft shadows are the difference
        between a flat rectangle and a "card" / KPI tile."""
        _, shape = self._require_shape(slide_index, shape_id)
        applied: dict = {}
        try:
            shadow = shape.Shadow
            if visible is False:
                shadow.Visible = msoFalse
                applied["visible"] = False
            elif (
                visible is True
                or color is not None
                or transparency is not None
                or blur is not None
                or offset_x is not None
                or offset_y is not None
            ):
                shadow.Visible = msoTrue
                applied["visible"] = True
            if color is not None:
                rgb = self._parse_color(color)
                shadow.ForeColor.RGB = rgb
                applied["rgb"] = rgb
            if transparency is not None:
                t = float(transparency)
                if not 0.0 <= t <= 1.0:
                    raise PowerPointError(f"transparency must be 0.0–1.0, got {transparency}")
                shadow.Transparency = t
                applied["transparency"] = t
            if blur is not None:
                if float(blur) < 0:
                    raise PowerPointError(f"blur must be ≥ 0, got {blur}")
                try:
                    shadow.Blur = float(blur)
                    applied["blur"] = float(blur)
                except pywintypes.com_error:
                    # .Blur is not exposed on every PowerPoint build — ignore.
                    pass
            if offset_x is not None:
                shadow.OffsetX = float(offset_x)
                applied["offset_x"] = float(offset_x)
            if offset_y is not None:
                shadow.OffsetY = float(offset_y)
                applied["offset_y"] = float(offset_y)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set shadow failed: {e}")
        if not applied:
            raise PowerPointError(
                "shape_set_shadow called with no attributes. Pass at least "
                "one of visible/color/transparency/blur/offset_x/offset_y."
            )
        return {
            "success": True,
            "shape_id": int(shape_id),
            "applied": applied,
        }
