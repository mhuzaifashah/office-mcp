"""PictureMixin — insert / replace / crop / border for picture shapes.

Mixed into PowerPointController. Relies on base for:
    self._validate_index(), self._require_shape(), self._parse_color(), self.pres
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import pywintypes

from .constants import DASH_STYLE_MAP, msoFalse, msoTrue
from .errors import PowerPointError


class PictureMixin:
    def picture_add_from_file(
        self,
        slide_index: int,
        file_path: str,
        left: float,
        top: float,
        width: float | None = None,
        height: float | None = None,
        link_to_file: bool = False,
        save_with_document: bool = True,
    ) -> dict:
        """Insert a picture from disk. If width/height omitted, uses native size."""
        self._validate_index(slide_index)
        p = Path(file_path).resolve()
        if not p.exists():
            raise PowerPointError(f"Picture file not found: {p}")
        slide = self.pres.Slides(slide_index)
        # AddPicture signature: FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height
        w = -1 if width is None else float(width)
        h = -1 if height is None else float(height)
        try:
            shape = slide.Shapes.AddPicture(
                FileName=str(p),
                LinkToFile=msoTrue if link_to_file else msoFalse,
                SaveWithDocument=msoTrue if save_with_document else msoFalse,
                Left=float(left),
                Top=float(top),
                Width=w,
                Height=h,
            )
        except pywintypes.com_error as e:
            raise PowerPointError(f"AddPicture failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "shape_id": int(shape.Id),
            "name": shape.Name,
            "left": float(shape.Left),
            "top": float(shape.Top),
            "width": float(shape.Width),
            "height": float(shape.Height),
            "linked": bool(link_to_file),
        }

    def picture_add_from_url(
        self,
        slide_index: int,
        url: str,
        left: float,
        top: float,
        width: float | None = None,
        height: float | None = None,
    ) -> dict:
        """Download a picture from a URL into a temp file, then insert it.

        Image stays embedded — the temp file is deleted after insertion since
        ``save_with_document=True`` means PowerPoint has already copied it
        into the .pptx package.
        """
        self._validate_index(slide_index)
        import tempfile
        import urllib.parse
        import urllib.request

        parsed = urllib.parse.urlparse(url)
        if parsed.scheme not in ("http", "https"):
            raise PowerPointError(f"Unsupported URL scheme '{parsed.scheme}'. Use http or https.")
        # Infer extension from URL path; default .png
        ext = Path(parsed.path).suffix.lower() or ".png"
        if ext not in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"):
            ext = ".png"
        try:
            req = urllib.request.Request(
                url,
                headers={"User-Agent": "office-mcp/0.1"},
            )
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = resp.read()
        except (OSError, ValueError) as e:
            raise PowerPointError(f"Failed to download {url}: {e}")

        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tf:
            tf.write(data)
            tmp_path = tf.name
        try:
            result = self.picture_add_from_file(
                slide_index=slide_index,
                file_path=tmp_path,
                left=left,
                top=top,
                width=width,
                height=height,
                link_to_file=False,
                save_with_document=True,
            )
            result["source_url"] = url
            return result
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    def picture_replace(
        self,
        slide_index: int,
        shape_id: int,
        file_path: str,
    ) -> dict:
        """Replace a picture's image while keeping its position/size/effects.

        PowerPoint has no atomic 'replace picture' API on most versions, so
        we delete and re-insert at the same bbox, returning the new id.
        """
        _, shape = self._require_shape(slide_index, shape_id)
        bbox = (
            float(shape.Left),
            float(shape.Top),
            float(shape.Width),
            float(shape.Height),
        )
        try:
            shape.Delete()
        except pywintypes.com_error as e:
            raise PowerPointError(f"Replace (delete) failed: {e}")
        result = self.picture_add_from_file(
            slide_index=slide_index,
            file_path=file_path,
            left=bbox[0],
            top=bbox[1],
            width=bbox[2],
            height=bbox[3],
        )
        result["replaced_shape_id"] = int(shape_id)
        return result

    def picture_crop(
        self,
        slide_index: int,
        shape_id: int,
        left: float | None = None,
        top: float | None = None,
        right: float | None = None,
        bottom: float | None = None,
    ) -> dict:
        """Set crop offsets in points. None = leave that edge unchanged."""
        _, shape = self._require_shape(slide_index, shape_id)
        try:
            crop = shape.PictureFormat.Crop
        except pywintypes.com_error as e:
            raise PowerPointError(f"Shape has no PictureFormat: {e}")
        applied: dict[str, float] = {}
        try:
            if left is not None:
                crop.PictureOffsetX = float(left)
                applied["left"] = float(left)
            if top is not None:
                crop.PictureOffsetY = float(top)
                applied["top"] = float(top)
            if right is not None:
                crop.ShapeWidth = max(1.0, float(shape.Width) - float(right))
                applied["right"] = float(right)
            if bottom is not None:
                crop.ShapeHeight = max(1.0, float(shape.Height) - float(bottom))
                applied["bottom"] = float(bottom)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Crop failed: {e}")
        if not applied:
            raise PowerPointError(
                "picture_crop called with no edges. Pass at least one of "
                "left/top/right/bottom (points)."
            )
        return {"success": True, "shape_id": int(shape_id), "applied": applied}

    def picture_set_border(
        self,
        slide_index: int,
        shape_id: int,
        color=None,
        weight: float | None = None,
        dash_style: int | str | None = None,
        visible: bool | None = None,
    ) -> dict:
        """Set the picture's outline (delegates to the same Line COM as shapes)."""
        _, shape = self._require_shape(slide_index, shape_id)
        applied: dict[str, Any] = {}
        try:
            line = shape.Line
            if visible is False:
                line.Visible = msoFalse
                applied["visible"] = False
            elif (
                color is not None or weight is not None or dash_style is not None or visible is True
            ):
                line.Visible = msoTrue
                applied["visible"] = True
            if color is not None:
                rgb = self._parse_color(color)
                line.ForeColor.RGB = rgb
                applied["rgb"] = rgb
            if weight is not None:
                if float(weight) <= 0:
                    raise PowerPointError(f"weight must be > 0, got {weight}")
                line.Weight = float(weight)
                applied["weight"] = float(weight)
            if dash_style is not None:
                if isinstance(dash_style, int):
                    val = dash_style
                else:
                    key = dash_style.lower().strip().replace(" ", "_")
                    if key not in DASH_STYLE_MAP:
                        raise PowerPointError(
                            f"Unknown dash_style '{dash_style}'. "
                            f"Choose from {sorted(DASH_STYLE_MAP)}"
                        )
                    val = DASH_STYLE_MAP[key]
                line.DashStyle = val
                applied["dash_style"] = val
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set border failed: {e}")
        if not applied:
            raise PowerPointError(
                "picture_set_border called with no attributes. Pass color, weight, "
                "dash_style, and/or visible."
            )
        return {"success": True, "shape_id": int(shape_id), "applied": applied}
