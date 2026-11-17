"""TextMixin — rich text: runs, paragraphs, fonts, bullets.

Mixed into PowerPointController. Relies on base for:
    self._require_shape(), self._parse_color()

PowerPoint text model:
  Shape.TextFrame.TextRange
    ├─ .Characters(start, length)   ← character-range slice (run)
    ├─ .Paragraphs(n)               ← per-paragraph access
    ├─ .Text                        ← whole-text get/set
    ├─ .InsertAfter(text) / .InsertBefore(text)
    ├─ .Font.{Name, Size, Bold, Italic, Underline, Color.RGB}
    └─ .ParagraphFormat.{Alignment, Bullet, IndentLevel, …}

All methods take (slide_index, shape_id, …) so they work on placeholders
AND custom shapes/text boxes uniformly.
"""

from __future__ import annotations

from typing import Any

import pywintypes

from .constants import (
    BULLET_TYPE_MAP,
    PARAGRAPH_ALIGN_MAP,
    msoFalse,
    msoTrue,
    ppBulletNone,
)
from .errors import PowerPointError


class TextMixin:
    def _text_range(self, shape):
        """Return a shape's TextRange or raise if no text frame."""
        try:
            if not shape.HasTextFrame:
                raise PowerPointError("Shape has no text frame.")
        except pywintypes.com_error as e:
            raise PowerPointError(f"Shape inaccessible: {e}")
        return shape.TextFrame.TextRange

    def _resolve_range(
        self,
        text_range,
        start: int | None,
        length: int | None,
        paragraph: int | None,
    ):
        """Resolve a sub-range. Priority: paragraph > (start,length) > full range."""
        if paragraph is not None:
            try:
                return text_range.Paragraphs(int(paragraph), 1)
            except pywintypes.com_error as e:
                raise PowerPointError(f"Paragraph {paragraph} not found: {e}")
        if start is not None or length is not None:
            s = 1 if start is None else int(start)
            n = -1 if length is None else int(length)
            try:
                return text_range.Characters(s, n)
            except pywintypes.com_error as e:
                raise PowerPointError(f"Characters({s}, {n}) failed: {e}")
        return text_range

    def text_set(self, slide_index: int, shape_id: int, text: str) -> dict:
        """Overwrite the entire text of any text-bearing shape."""
        _, shape = self._require_shape(slide_index, shape_id)
        tr = self._text_range(shape)
        try:
            tr.Text = text
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set text failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "text_length": len(text),
        }

    def text_append(
        self,
        slide_index: int,
        shape_id: int,
        text: str,
        new_paragraph: bool = False,
    ) -> dict:
        """Append text to a shape's existing text. ``new_paragraph`` prepends
        a newline so the new text starts on its own paragraph."""
        _, shape = self._require_shape(slide_index, shape_id)
        tr = self._text_range(shape)
        payload = ("\r" + text) if new_paragraph else text
        try:
            tr.InsertAfter(payload)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Append failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "appended_length": len(text),
            "new_paragraph": bool(new_paragraph),
        }

    def text_clear(self, slide_index: int, shape_id: int) -> dict:
        """Clear all text on a shape's text frame."""
        return self.text_set(slide_index, shape_id, "")

    def text_set_font(
        self,
        slide_index: int,
        shape_id: int,
        font: str | None = None,
        size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color=None,
        start: int | None = None,
        length: int | None = None,
        paragraph: int | None = None,
    ) -> dict:
        """Set font attributes on a range. Any unset param is left unchanged.

        Range selection (mutually exclusive, highest-priority first):
          - paragraph : 1-based paragraph index
          - start + length : character range (1-based start, length=-1 for "to end")
          - none of the above : whole text frame
        """
        _, shape = self._require_shape(slide_index, shape_id)
        tr = self._text_range(shape)
        rng = self._resolve_range(tr, start, length, paragraph)
        try:
            f = rng.Font
            applied: dict[str, Any] = {}
            if font is not None:
                f.Name = str(font)
                applied["font"] = font
            if size is not None:
                f.Size = float(size)
                applied["size"] = size
            if bold is not None:
                f.Bold = msoTrue if bold else msoFalse
                applied["bold"] = bold
            if italic is not None:
                f.Italic = msoTrue if italic else msoFalse
                applied["italic"] = italic
            if underline is not None:
                f.Underline = msoTrue if underline else msoFalse
                applied["underline"] = underline
            if color is not None:
                rgb = self._parse_color(color)
                f.Color.RGB = rgb
                applied["rgb"] = rgb
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set font failed: {e}")
        if not applied:
            raise PowerPointError(
                "text_set_font called with no attributes. Pass at least one of "
                "font/size/bold/italic/underline/color."
            )
        return {
            "success": True,
            "shape_id": int(shape_id),
            "applied": applied,
            "range": {"start": start, "length": length, "paragraph": paragraph},
        }

    def text_set_paragraph(
        self,
        slide_index: int,
        shape_id: int,
        alignment: str | None = None,
        bullet: str | None = None,
        indent_level: int | None = None,
        line_spacing: float | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        paragraph: int | None = None,
    ) -> dict:
        """Set paragraph attributes. If ``paragraph`` is None, applies to all paragraphs.

        - alignment: 'left' | 'center' | 'right' | 'justify' | 'distribute'
        - bullet: 'none' | 'bullet' | 'numbered'
        - indent_level: 1-5
        - line_spacing: multiplier (e.g. 1.5)
        - space_before / space_after: points
        """
        _, shape = self._require_shape(slide_index, shape_id)
        tr = self._text_range(shape)
        if paragraph is not None:
            try:
                rng = tr.Paragraphs(int(paragraph), 1)
            except pywintypes.com_error as e:
                raise PowerPointError(f"Paragraph {paragraph} not found: {e}")
        else:
            rng = tr
        applied: dict[str, Any] = {}
        try:
            pf = rng.ParagraphFormat
            if alignment is not None:
                key = alignment.lower().strip()
                if key not in PARAGRAPH_ALIGN_MAP:
                    raise PowerPointError(
                        f"Unknown alignment '{alignment}'. "
                        f"Choose from {sorted(PARAGRAPH_ALIGN_MAP)}"
                    )
                pf.Alignment = PARAGRAPH_ALIGN_MAP[key]
                applied["alignment"] = key
            if bullet is not None:
                key = bullet.lower().strip()
                if key not in BULLET_TYPE_MAP:
                    raise PowerPointError(
                        f"Unknown bullet '{bullet}'. Choose from {sorted(BULLET_TYPE_MAP)}"
                    )
                val = BULLET_TYPE_MAP[key]
                if val == ppBulletNone:
                    pf.Bullet.Visible = msoFalse
                else:
                    pf.Bullet.Visible = msoTrue
                    pf.Bullet.Type = val
                applied["bullet"] = key
            if indent_level is not None:
                lvl = int(indent_level)
                if lvl < 1 or lvl > 5:
                    raise PowerPointError(f"indent_level must be 1-5, got {indent_level}")
                # ParagraphFormat.IndentLevel is read-only on most PowerPoint
                # builds; the writable property lives on TextRange itself.
                rng.IndentLevel = lvl
                applied["indent_level"] = lvl
            if line_spacing is not None:
                pf.SpaceWithin = float(line_spacing)
                applied["line_spacing"] = float(line_spacing)
            if space_before is not None:
                pf.SpaceBefore = float(space_before)
                applied["space_before"] = float(space_before)
            if space_after is not None:
                pf.SpaceAfter = float(space_after)
                applied["space_after"] = float(space_after)
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set paragraph failed: {e}")
        if not applied:
            raise PowerPointError(
                "text_set_paragraph called with no attributes. Pass at least one "
                "of alignment/bullet/indent_level/line_spacing/space_before/space_after."
            )
        return {
            "success": True,
            "shape_id": int(shape_id),
            "paragraph": paragraph,
            "applied": applied,
        }

    def text_get_runs(self, slide_index: int, shape_id: int) -> dict:
        """Read per-paragraph text and effective formatting from a shape."""
        _, shape = self._require_shape(slide_index, shape_id)
        tr = self._text_range(shape)
        try:
            n = tr.Paragraphs().Count
        except pywintypes.com_error as e:
            raise PowerPointError(f"Inspect failed: {e}")
        paragraphs: list[dict] = []
        for i in range(1, n + 1):
            try:
                p = tr.Paragraphs(i, 1)
                pf = p.ParagraphFormat
                font = p.Font
                try:
                    indent = int(p.IndentLevel)
                except (pywintypes.com_error, AttributeError):
                    indent = None
                paragraphs.append(
                    {
                        "index": i,
                        "text": p.Text or "",
                        "alignment": int(pf.Alignment) if pf.Alignment else None,
                        "indent_level": indent,
                        "bullet_visible": bool(pf.Bullet.Visible == -1),
                        "font": font.Name,
                        "size": float(font.Size) if font.Size else None,
                        "bold": bool(font.Bold == -1),
                        "italic": bool(font.Italic == -1),
                        "underline": bool(font.Underline == -1),
                        "rgb": int(font.Color.RGB) if font.Color.RGB else None,
                    }
                )
            except pywintypes.com_error:
                continue
        return {"shape_id": int(shape_id), "paragraph_count": n, "paragraphs": paragraphs}
