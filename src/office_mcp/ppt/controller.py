"""
PowerPointController — programmatic PowerPoint control via pywin32 COM.

Base class hosting app/window, file/session, slide, snapshot, connector,
and hyperlink ops, plus cross-cutting helpers (_parse_color,
_find_shape_by_id, _require_shape, _validate_index, _shape_summary) used
by the mixins. Domain-specific work that talks to a different subsystem
lives in sibling modules:

    shapes.py   — ShapeMixin   (autoshape, fill, line, rotate, z-order, group, align)
    text.py     — TextMixin    (rich text — runs, fonts, paragraphs, bullets)
    pictures.py — PictureMixin (insert / replace / crop / border)
    tables.py   — TableMixin   (table CRUD, cells, sizing, merge)
    charts.py   — ChartMixin   (Excel-backed charts — separate file because
                                chart data lives in an embedded workbook)

Reference: ykuwai/ppt-mcp. Same Office Application object model, synchronous
single-class wrapper instead of an STA worker thread, matching the
ExcelController pattern in the sibling `excel` package.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client

from .charts import ChartMixin
from .constants import (
    CONNECTOR_TYPE_MAP,
    LAYOUT_NAME_MAP,
    PLACEHOLDER_NAME_MAP,
    PLACEHOLDER_TYPE_NAMES,
    SAVE_FORMAT_MAP,
    SHAPE_TYPE_NAMES,
    SLIDE_SIZE_PRESETS,
    VALID_EXTENSIONS,
    WINDOW_STATE_NAMES,
    WINDOW_STATE_VALUES,
    msoConnectorStraight,
    msoFalse,
    msoPlaceholder,
    msoTrue,
    ppLayoutBlank,
    ppPlaceholderBody,
    ppSaveAsOpenXMLPresentation,
)
from .errors import PowerPointError
from .pictures import PictureMixin
from .shapes import ShapeMixin
from .tables import TableMixin
from .text import TextMixin

# ActionSettings constants for hyperlinks
_PP_MOUSE_CLICK = 1
_PP_ACTION_NONE = 0
_PP_ACTION_HYPERLINK = 7


# ─────────────────────────────────────────────────────────────
# Snapshot housekeeping
# ─────────────────────────────────────────────────────────────
# How long a snapshot PNG is kept before background cleanup deletes it.
# TODO(prod): once the default snapshot dir moves under the system temp,
# this knob can be removed and OS-level cleanup tooling can take over.
SNAPSHOT_MAX_AGE_SECONDS: int = 24 * 60 * 60  # 24 hours


def _spawn_snapshot_cleanup(base_dir: "Path") -> None:
    """Fire-and-forget removal of snapshot PNGs older than 24h.

    Runs on a daemon thread so it never blocks the calling thread (the
    COM-pinned worker that performs the actual ``Slide.Export``). The
    agent's async loop is therefore never blocked by cleanup. Errors are
    swallowed — a stale file is harmless and cleanup must never surface
    as a tool failure.
    """
    import threading
    import time

    def _worker(target_dir: "Path") -> None:
        try:
            if not target_dir.exists():
                return
            cutoff = time.time() - SNAPSHOT_MAX_AGE_SECONDS
            for entry in target_dir.iterdir():
                try:
                    if not entry.is_file():
                        continue
                    if entry.suffix.lower() != ".png":
                        continue
                    if entry.stat().st_mtime < cutoff:
                        entry.unlink(missing_ok=True)
                except OSError:
                    continue
        except OSError:
            return

    threading.Thread(
        target=_worker,
        args=(base_dir,),
        daemon=True,
        name="ppt-snapshot-cleanup",
    ).start()


class PowerPointController(ShapeMixin, TextMixin, PictureMixin, TableMixin, ChartMixin):
    """High-level PowerPoint automation controller.

    Manages a single PowerPoint Application COM object and an active
    presentation. All ops route through `self._app` / `self._pres`.
    Domain methods (shape_*, text_*, picture_*, table_*) come from the
    mixins listed in the base list.
    """

    PROG_ID = "PowerPoint.Application"

    def __init__(self, visible: bool = True, app: Any = None):
        pythoncom.CoInitialize()
        if app is not None:
            self._app = app
        else:
            try:
                # Attach to a running PowerPoint if one exists.
                self._app = win32com.client.GetActiveObject(self.PROG_ID)
            except pywintypes.com_error:
                try:
                    self._app = win32com.client.Dispatch(self.PROG_ID)
                except pywintypes.com_error as e:
                    raise PowerPointError(
                        "Failed to launch or attach to PowerPoint. Verify PowerPoint "
                        f"is installed and the COM ProgID '{self.PROG_ID}' is registered. "
                        f"Underlying error: {e}"
                    )
        try:
            self._app.Visible = msoTrue if visible else msoFalse
        except pywintypes.com_error:
            # PowerPoint disallows Visible=False on some versions; non-fatal.
            pass
        self._pres = None  # active presentation handle

    # ─────────────────────────────────────────────
    # INTERNAL
    # ─────────────────────────────────────────────

    @property
    def app(self):
        return self._app

    @property
    def pres(self):
        if self._pres is None:
            raise PowerPointError(
                "No presentation is open. Call open(), create(), or connect() first."
            )
        return self._pres

    def _find_pres_by_name(self, name: str):
        """Locate an open presentation by Name or FullName (case-insensitive)."""
        target = name.lower()
        for i in range(1, self._app.Presentations.Count + 1):
            p = self._app.Presentations(i)
            if p.Name.lower() == target or p.FullName.lower() == target:
                return p
        return None

    # ─────────────────────────────────────────────
    # Cross-cutting helpers (used by mixins)
    # ─────────────────────────────────────────────

    def _validate_index(self, slide_index: int) -> int:
        count = self.pres.Slides.Count
        if slide_index < 1 or slide_index > count:
            raise PowerPointError(f"Slide index {slide_index} out of range (1-{count})")
        return count

    def _find_shape_by_id(self, slide, shape_id: int):
        """Look up a shape on a slide by its COM Id. Returns shape or None."""
        target = int(shape_id)
        for i in range(1, slide.Shapes.Count + 1):
            sh = slide.Shapes(i)
            try:
                if int(sh.Id) == target:
                    return sh
            except pywintypes.com_error:
                continue
        return None

    def _require_shape(self, slide_index: int, shape_id: int):
        """Locate a shape by id or raise. Returns (slide, shape)."""
        self._validate_index(slide_index)
        slide = self.pres.Slides(slide_index)
        shape = self._find_shape_by_id(slide, shape_id)
        if shape is None:
            raise PowerPointError(f"Shape id={shape_id} not found on slide {slide_index}.")
        return slide, shape

    @staticmethod
    def _parse_color(color) -> int:
        """Accept '#RRGGBB' / 'RRGGBB' / [r,g,b] / (r,g,b) / int and return
        PowerPoint's BGR int (RGB() macro: r | g<<8 | b<<16)."""
        if isinstance(color, int):
            return color
        if isinstance(color, (list, tuple)) and len(color) == 3:
            r, g, b = (int(c) & 0xFF for c in color)
            return r | (g << 8) | (b << 16)
        if isinstance(color, str):
            s = color.strip().lstrip("#")
            if len(s) == 6:
                try:
                    r = int(s[0:2], 16)
                    g = int(s[2:4], 16)
                    b = int(s[4:6], 16)
                    return r | (g << 8) | (b << 16)
                except ValueError:
                    pass
        raise PowerPointError(
            f"Invalid color {color!r}. Use '#RRGGBB', 'RRGGBB', [r,g,b], or BGR int."
        )

    # ─────────────────────────────────────────────
    # FILE / SESSION
    # ─────────────────────────────────────────────

    def list_open_presentations(self) -> list[dict]:
        """List every presentation currently open in the running PowerPoint."""
        results: list[dict] = []
        for i in range(1, self._app.Presentations.Count + 1):
            p = self._app.Presentations(i)
            results.append(
                {
                    "index": i,
                    "name": p.Name,
                    "full_name": p.FullName,
                    "path": p.Path,
                    "slides_count": p.Slides.Count,
                    "read_only": int(p.ReadOnly) == -1,
                    "saved": bool(p.Saved),
                }
            )
        return results

    def connect(self, name_or_path: str) -> dict:
        """Attach to an already-open presentation by Name or full path."""
        pres = self._find_pres_by_name(name_or_path)
        if pres is None:
            # Try absolute-path resolution
            try:
                resolved = str(Path(name_or_path).resolve())
                pres = self._find_pres_by_name(resolved)
            except (OSError, ValueError):
                pass
        if pres is None:
            # Build a helpful error listing currently-open presentations.
            open_list = [
                self._app.Presentations(i).Name for i in range(1, self._app.Presentations.Count + 1)
            ]
            raise PowerPointError(
                f"Presentation '{name_or_path}' is not open in PowerPoint. "
                f"Currently open: {open_list or '(none)'}. "
                f"Use open() to load from disk, or create() to make a new one."
            )
        self._pres = pres
        return {
            "success": True,
            "name": pres.Name,
            "full_name": pres.FullName,
            "slides_count": pres.Slides.Count,
        }

    def test_file(self, file_path: str) -> dict:
        """Validate a file before opening."""
        p = Path(file_path).resolve()
        ext = p.suffix.lower()
        exists = p.exists()
        return {
            "file_path": str(p),
            "exists": exists,
            "extension": ext,
            "is_valid": exists and ext in VALID_EXTENSIONS,
            "size": p.stat().st_size if exists else 0,
        }

    def open(self, file_path: str, read_only: bool = False, with_window: bool = True) -> dict:
        """Open a presentation. If it's already open, attach to it instead."""
        p = Path(file_path).resolve()
        if not p.exists():
            raise PowerPointError(f"File not found: {p}")
        if p.suffix.lower() not in VALID_EXTENSIONS:
            raise PowerPointError(f"Unsupported extension: {p.suffix}")

        existing = self._find_pres_by_name(str(p))
        if existing is None:
            existing = self._find_pres_by_name(p.name)
        if existing is not None:
            self._pres = existing
            return {
                "success": True,
                "file_path": str(p),
                "slides_count": existing.Slides.Count,
                "already_open": True,
            }

        self._pres = self._app.Presentations.Open(
            FileName=str(p),
            ReadOnly=msoTrue if read_only else msoFalse,
            Untitled=msoFalse,
            WithWindow=msoTrue if with_window else msoFalse,
        )
        return {
            "success": True,
            "file_path": str(p),
            "slides_count": self._pres.Slides.Count,
            "already_open": False,
        }

    def create(
        self,
        file_path: str | None = None,
        preset: str | None = None,
        slide_width: float | None = None,
        slide_height: float | None = None,
        template_path: str | None = None,
    ) -> dict:
        """Create a new presentation.

        If `template_path` is given, open the template Untitled. Otherwise
        Add() a blank presentation and optionally apply a size preset.
        If `file_path` is given, the new presentation is saved there.
        """
        if template_path:
            tpath = Path(template_path).resolve()
            if not tpath.exists():
                raise PowerPointError(f"Template not found: {tpath}")
            self._pres = self._app.Presentations.Open(str(tpath), 0, msoTrue, msoTrue)
        else:
            self._pres = self._app.Presentations.Add()
            if preset:
                if preset not in SLIDE_SIZE_PRESETS:
                    raise PowerPointError(
                        f"Unknown preset '{preset}'. Choose from {list(SLIDE_SIZE_PRESETS)}"
                    )
                w, h = SLIDE_SIZE_PRESETS[preset]
                self._pres.PageSetup.SlideWidth = w
                self._pres.PageSetup.SlideHeight = h
            elif slide_width is not None and slide_height is not None:
                self._pres.PageSetup.SlideWidth = slide_width
                self._pres.PageSetup.SlideHeight = slide_height

        if file_path:
            p = Path(file_path).resolve()
            self._pres.SaveAs(FileName=str(p), FileFormat=ppSaveAsOpenXMLPresentation)

        return {
            "success": True,
            "name": self._pres.Name,
            "full_name": self._pres.FullName,
            "slides_count": self._pres.Slides.Count,
            "slide_width": self._pres.PageSetup.SlideWidth,
            "slide_height": self._pres.PageSetup.SlideHeight,
        }

    def save(self, file_path: str | None = None, format: str | None = None) -> dict:
        """Save the current presentation. With `file_path`, performs SaveAs.

        - `save()`              → in-place Save() (must have been saved at least once)
        - `save(path)`          → SaveAs(path) — format inferred from extension
        - `save(path, format)`  → SaveAs(path, FileFormat=<map[format]>)
        - `save(format=...)`    → ERROR — format requires a file_path
        """
        if format and not file_path:
            raise PowerPointError(
                "save(format=...) requires a file_path. Pass file_path to perform SaveAs."
            )
        if file_path:
            kwargs: dict[str, Any] = {"FileName": str(Path(file_path).resolve())}
            if format:
                fmt_key = format.lower().strip()
                if fmt_key not in SAVE_FORMAT_MAP:
                    raise PowerPointError(
                        f"Unknown format '{format}'. Choose from {list(SAVE_FORMAT_MAP)}"
                    )
                kwargs["FileFormat"] = SAVE_FORMAT_MAP[fmt_key]
            self.pres.SaveAs(**kwargs)
        else:
            self.pres.Save()
        return {
            "success": True,
            "name": self.pres.Name,
            "full_name": self.pres.FullName,
            "saved": bool(self.pres.Saved),
        }

    def close(self, save: bool = False) -> dict:
        """Close the active presentation."""
        name = self.pres.Name
        if save:
            self.pres.Save()
        else:
            self.pres.Saved = True  # suppress save prompt
        self.pres.Close()
        self._pres = None
        return {"success": True, "closed": name}

    def quit(self) -> dict:
        """Close active presentation (if any) and quit PowerPoint."""
        if self._pres is not None:
            try:
                self._pres.Saved = True
                self._pres.Close()
            except pywintypes.com_error:
                pass
            self._pres = None
        try:
            self._app.Quit()
        except pywintypes.com_error:
            pass
        return {"success": True}

    # ─────────────────────────────────────────────
    # APP / WINDOW
    # ─────────────────────────────────────────────

    def app_info(self) -> dict:
        """Application + active-presentation snapshot."""
        info: dict[str, Any] = {
            "name": self._app.Name,
            "version": self._app.Version,
            "visible": bool(self._app.Visible),
            "window_state": WINDOW_STATE_NAMES.get(self._app.WindowState, "unknown"),
            "presentations_count": self._app.Presentations.Count,
            "active_presentation": None,
        }
        if self._app.Presentations.Count > 0:
            try:
                info["active_presentation"] = self._app.ActivePresentation.Name
            except pywintypes.com_error:
                pass
        return info

    def set_window_state(self, state: str) -> dict:
        """Set PowerPoint window state: 'normal' / 'minimized' / 'maximized'."""
        key = state.lower().strip()
        if key not in WINDOW_STATE_VALUES:
            raise PowerPointError(
                f"Invalid window_state '{state}'. Use {list(WINDOW_STATE_VALUES)}"
            )
        self._app.WindowState = WINDOW_STATE_VALUES[key]
        return {"success": True, "window_state": key}

    # ─────────────────────────────────────────────
    # SLIDES
    # ─────────────────────────────────────────────

    def _custom_layouts_by_design(self) -> dict[str, list[str]]:
        """Collect every CustomLayout.Name on the active presentation,
        grouped by design name. Used for layout error messages and `list_layouts`."""
        result: dict[str, list[str]] = {}
        try:
            count = self.pres.Designs.Count
        except pywintypes.com_error:
            return result
        for d in range(1, count + 1):
            design = self.pres.Designs(d)
            try:
                master = design.SlideMaster
                names = [
                    master.CustomLayouts(i).Name for i in range(1, master.CustomLayouts.Count + 1)
                ]
            except pywintypes.com_error:
                names = []
            result[design.Name] = names
        return result

    def _resolve_layout(self, layout: int | str | None):
        """Return (layout_int, custom_layout_obj). CustomLayout takes priority.

        Accepts:
          - int  : raw PpSlideLayout constant
          - str  : friendly name (see LAYOUT_NAME_MAP) or CustomLayout.Name
          - None : default blank
        """
        if layout is None:
            return ppLayoutBlank, None
        if isinstance(layout, int):
            return layout, None

        # string: friendly map first, then walk all designs
        key = layout.lower().strip().replace(" ", "_")
        if key in LAYOUT_NAME_MAP:
            return LAYOUT_NAME_MAP[key], None

        target = layout.strip()
        for d in range(1, self.pres.Designs.Count + 1):
            master = self.pres.Designs(d).SlideMaster
            for i in range(1, master.CustomLayouts.Count + 1):
                cl = master.CustomLayouts(i)
                if cl.Name == target:
                    return None, cl

        raise PowerPointError(
            f"Layout '{layout}' not found.\n"
            f"  Friendly names: {sorted(LAYOUT_NAME_MAP)}\n"
            f"  Custom layouts by design: {self._custom_layouts_by_design()}"
        )

    def list_layouts(self) -> dict:
        """Inspect every layout available to the active presentation."""
        return {
            "friendly_names": sorted(LAYOUT_NAME_MAP),
            "custom_layouts_by_design": self._custom_layouts_by_design(),
        }

    def slide_list(self) -> list[dict]:
        """List all slides in the active presentation."""
        results: list[dict] = []
        for i in range(1, self.pres.Slides.Count + 1):
            s = self.pres.Slides(i)
            layout_name = ""
            try:
                layout_name = s.CustomLayout.Name
            except pywintypes.com_error:
                pass
            has_notes = False
            try:
                txt = s.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
                has_notes = bool(txt and txt.strip())
            except pywintypes.com_error:
                pass
            results.append(
                {
                    "index": s.SlideIndex,
                    "slide_id": s.SlideID,
                    "name": s.Name,
                    "layout": s.Layout,
                    "layout_name": layout_name,
                    "hidden": bool(s.SlideShowTransition.Hidden),
                    "shapes_count": s.Shapes.Count,
                    "has_notes": has_notes,
                }
            )
        return results

    def slide_add(
        self,
        position: int | None = None,
        layout: int | str | None = None,
    ) -> dict:
        """Add a new slide at `position` (1-based) with `layout`.

        `layout` may be a PpSlideLayout int, a friendly name, or a
        CustomLayout name; if omitted, defaults to blank.
        """
        count = self.pres.Slides.Count
        if position is None:
            position = count + 1
        if position < 1 or position > count + 1:
            raise PowerPointError(f"Position {position} out of range (1-{count + 1})")

        layout_int, custom_layout = self._resolve_layout(layout)
        if custom_layout is not None:
            slide = self.pres.Slides.AddSlide(Index=position, pCustomLayout=custom_layout)
        else:
            slide = self.pres.Slides.Add(Index=position, Layout=layout_int)

        return {
            "success": True,
            "slide_index": slide.SlideIndex,
            "slide_id": slide.SlideID,
            "layout": slide.Layout,
        }

    def slide_delete(self, slide_index: int) -> dict:
        """Delete a slide by 1-based index."""
        self._validate_index(slide_index)
        self.pres.Slides(slide_index).Delete()
        return {
            "success": True,
            "deleted_index": slide_index,
            "remaining_count": self.pres.Slides.Count,
        }

    def slide_duplicate(self, slide_index: int) -> dict:
        """Duplicate a slide; the duplicate is inserted immediately after."""
        self._validate_index(slide_index)
        dup_range = self.pres.Slides(slide_index).Duplicate()
        new_slide = dup_range(1)
        return {
            "success": True,
            "source_index": slide_index,
            "new_slide_index": new_slide.SlideIndex,
            "new_slide_id": new_slide.SlideID,
        }

    def slide_move(self, slide_index: int, new_position: int) -> dict:
        """Move a slide from `slide_index` to `new_position` (both 1-based)."""
        count = self._validate_index(slide_index)
        if new_position < 1 or new_position > count:
            raise PowerPointError(f"New position {new_position} out of range (1-{count})")
        # Use positional arg — some PowerPoint COM versions reject the
        # named `toPos` parameter even though it's documented.
        self.pres.Slides(slide_index).MoveTo(new_position)
        return {"success": True, "moved_from": slide_index, "moved_to": new_position}

    def slide_goto(self, slide_index: int) -> dict:
        """Navigate the active window to a specific slide."""
        self._validate_index(slide_index)
        try:
            self._app.ActiveWindow.View.GotoSlide(slide_index)
        except pywintypes.com_error as e:
            raise PowerPointError(f"GotoSlide failed: {e}")
        return {"success": True, "slide_index": slide_index}

    def slide_set_layout(self, slide_index: int, layout: int | str) -> dict:
        """Change the layout of an existing slide."""
        self._validate_index(slide_index)
        layout_int, custom_layout = self._resolve_layout(layout)
        slide = self.pres.Slides(slide_index)
        if custom_layout is not None:
            slide.CustomLayout = custom_layout
        else:
            slide.Layout = layout_int
        return {"success": True, "slide_index": slide_index, "layout": slide.Layout}

    def slide_set_hidden(self, slide_index: int, hidden: bool) -> dict:
        """Hide or unhide a slide during slideshow."""
        self._validate_index(slide_index)
        self.pres.Slides(slide_index).SlideShowTransition.Hidden = msoTrue if hidden else msoFalse
        return {"success": True, "slide_index": slide_index, "hidden": hidden}

    def _notes_body_placeholder(self, slide):
        """Locate the notes body placeholder on a slide's NotesPage.

        Slide-number / header / footer placeholders may exist alongside the
        body, so we can't assume index 2. Search by placeholder type first,
        then by the largest text-bearing placeholder as a fallback.
        """
        try:
            placeholders = slide.NotesPage.Shapes.Placeholders
        except pywintypes.com_error:
            return None
        body = None
        best_area = 0.0
        for i in range(1, placeholders.Count + 1):
            ph = placeholders(i)
            try:
                ph_type = ph.PlaceholderFormat.Type
            except pywintypes.com_error:
                ph_type = None
            if ph_type == ppPlaceholderBody:
                return ph
            try:
                if ph.HasTextFrame:
                    area = float(ph.Width) * float(ph.Height)
                    if area > best_area:
                        best_area = area
                        body = ph
            except pywintypes.com_error:
                continue
        return body

    def slide_get_notes(self, slide_index: int) -> dict:
        """Read the speaker notes for a slide."""
        self._validate_index(slide_index)
        ph = self._notes_body_placeholder(self.pres.Slides(slide_index))
        text = ""
        if ph is not None:
            try:
                text = ph.TextFrame.TextRange.Text or ""
            except pywintypes.com_error:
                text = ""
        return {"slide_index": slide_index, "notes": text}

    def slide_set_notes(self, slide_index: int, notes: str) -> dict:
        """Set the speaker notes for a slide (overwrites existing)."""
        self._validate_index(slide_index)
        ph = self._notes_body_placeholder(self.pres.Slides(slide_index))
        if ph is None:
            raise PowerPointError(f"Slide {slide_index} has no notes body placeholder.")
        try:
            ph.TextFrame.TextRange.Text = notes
        except pywintypes.com_error as e:
            raise PowerPointError(f"Failed to set notes: {e}")
        return {"success": True, "slide_index": slide_index, "notes_length": len(notes)}

    # ─────────────────────────────────────────────
    # SHAPE / TEXT INTROSPECTION
    # ─────────────────────────────────────────────

    def _shape_summary(self, shape) -> dict:
        """Compact, JSON-safe summary of a single shape."""
        info: dict[str, Any] = {
            "id": int(shape.Id),
            "name": shape.Name,
            "shape_type": int(shape.Type),
            "shape_type_name": SHAPE_TYPE_NAMES.get(int(shape.Type), "unknown"),
            "left": float(shape.Left),
            "top": float(shape.Top),
            "width": float(shape.Width),
            "height": float(shape.Height),
        }
        # Placeholder details
        if int(shape.Type) == msoPlaceholder:
            try:
                ph_type = int(shape.PlaceholderFormat.Type)
                info["placeholder_type"] = ph_type
                info["placeholder_type_name"] = PLACEHOLDER_TYPE_NAMES.get(ph_type, "unknown")
            except pywintypes.com_error:
                pass
        # Text extract (if any)
        try:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                info["text"] = shape.TextFrame.TextRange.Text or ""
        except pywintypes.com_error:
            pass
        return info

    def slide_get_info(self, slide_index: int) -> dict:
        """Detailed slide info: layout, transitions, shapes (with text + bbox)."""
        self._validate_index(slide_index)
        s = self.pres.Slides(slide_index)
        layout_name = ""
        try:
            layout_name = s.CustomLayout.Name
        except pywintypes.com_error:
            pass
        shapes: list[dict] = []
        for i in range(1, s.Shapes.Count + 1):
            try:
                shapes.append(self._shape_summary(s.Shapes(i)))
            except pywintypes.com_error:
                continue
        return {
            "index": s.SlideIndex,
            "slide_id": s.SlideID,
            "name": s.Name,
            "layout": s.Layout,
            "layout_name": layout_name,
            "hidden": bool(s.SlideShowTransition.Hidden),
            "slide_width": float(self.pres.PageSetup.SlideWidth),
            "slide_height": float(self.pres.PageSetup.SlideHeight),
            "shapes_count": s.Shapes.Count,
            "shapes": shapes,
        }

    def slide_get_text(self, slide_index: int) -> dict:
        """Extract all visible text from a slide (one entry per text-bearing shape)."""
        self._validate_index(slide_index)
        s = self.pres.Slides(slide_index)
        items: list[dict] = []
        for i in range(1, s.Shapes.Count + 1):
            shape = s.Shapes(i)
            try:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    items.append(
                        {
                            "shape_id": int(shape.Id),
                            "shape_name": shape.Name,
                            "text": shape.TextFrame.TextRange.Text or "",
                        }
                    )
            except pywintypes.com_error:
                continue
        return {"slide_index": slide_index, "items": items}

    def _find_placeholder(self, slide, kind: int | str):
        """Locate a placeholder shape on a slide by PpPlaceholderType.

        Accepts an int (raw enum) or a friendly key from PLACEHOLDER_NAME_MAP.
        Friendly names may map to multiple types — e.g. 'title' matches both
        ppPlaceholderTitle (1, used by 'Title and Content' layout) and
        ppPlaceholderCenterTitle (3, used by 'Title Slide' layout). Returns
        the first matching shape or None.
        """
        if isinstance(kind, str):
            key = kind.lower().strip().replace(" ", "_")
            if key not in PLACEHOLDER_NAME_MAP:
                raise PowerPointError(
                    f"Unknown placeholder '{kind}'. Choose from {sorted(PLACEHOLDER_NAME_MAP)}"
                )
            target_types: tuple[int, ...] = PLACEHOLDER_NAME_MAP[key]
        else:
            target_types = (int(kind),)

        placeholders = slide.Shapes.Placeholders
        for i in range(1, placeholders.Count + 1):
            ph = placeholders(i)
            try:
                if int(ph.PlaceholderFormat.Type) in target_types:
                    return ph
            except pywintypes.com_error:
                continue
        return None

    def placeholder_set_text(self, slide_index: int, placeholder: int | str, text: str) -> dict:
        """Fill a placeholder (title/body/subtitle/etc.) with text on a slide."""
        self._validate_index(slide_index)
        slide = self.pres.Slides(slide_index)
        ph = self._find_placeholder(slide, placeholder)
        if ph is None:
            available = []
            for i in range(1, slide.Shapes.Placeholders.Count + 1):
                p = slide.Shapes.Placeholders(i)
                try:
                    t = int(p.PlaceholderFormat.Type)
                    available.append(PLACEHOLDER_TYPE_NAMES.get(t, f"type={t}"))
                except pywintypes.com_error:
                    pass
            raise PowerPointError(
                f"Placeholder '{placeholder}' not found on slide {slide_index}. "
                f"Available: {available or '(none)'}."
            )
        try:
            ph.TextFrame.TextRange.Text = text
        except pywintypes.com_error as e:
            raise PowerPointError(f"Failed to set placeholder text: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "placeholder": placeholder,
            "shape_id": int(ph.Id),
            "text_length": len(text),
        }

    # ─────────────────────────────────────────────
    # SNAPSHOT (slide → PNG, optional bbox overlay)
    # ─────────────────────────────────────────────

    def slide_snapshot(
        self,
        slide_index: int,
        file_path: str | None = None,
        with_bboxes: bool = True,
        width: int | None = None,
        height: int | None = None,
    ) -> dict:
        """Export a single slide as PNG, optionally annotating shapes with bounding boxes.

        Args:
            slide_index : 1-based slide to capture.
            file_path   : Destination PNG path. If omitted, written to a
                          managed dir under the system temp dir
                          (office-mcp/ppt-snapshots/).
            with_bboxes : When True, overlays each shape's rectangle and ID label.
            width/height: Optional explicit pixel size for Slide.Export. Default
                          uses PowerPoint's native resolution.

        Returns: {success, slide_index, path, image_size, objects_annotated}
        """
        import tempfile
        from datetime import datetime

        self._validate_index(slide_index)
        slide = self.pres.Slides(slide_index)

        # Resolve output path
        # Default snapshots go under the system temp dir so OS cleanup tooling
        # (Storage Sense / Disk Cleanup) can reclaim them.
        if file_path:
            out_path = Path(file_path).resolve()
            out_path.parent.mkdir(parents=True, exist_ok=True)
            base_dir = out_path.parent
        else:
            base_dir = Path(tempfile.gettempdir()) / "office-mcp" / "ppt-snapshots"
            base_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
            out_path = base_dir / f"slide-{slide_index}-{stamp}.png"

        # Fire-and-forget background cleanup of stale snapshots (>24h old).
        # Runs on a daemon thread so it never blocks the snapshot return —
        # the agent loop sees the snapshot result immediately while old PNGs
        # are pruned in parallel. Cleanup failures are silent on purpose; a
        # stale file is harmless and we never want cleanup to surface as a
        # tool error.
        _spawn_snapshot_cleanup(base_dir)

        # Export to a temp PNG first (so we can re-annotate without
        # touching the user's file_path on failure).
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
            tmp_export = tf.name

        try:
            if width and height:
                slide.Export(tmp_export, "PNG", int(width), int(height))
            else:
                slide.Export(tmp_export, "PNG")

            if not with_bboxes:
                # Just move the file into place.
                Path(tmp_export).replace(out_path)
                img_size = None
                try:
                    from PIL import Image

                    with Image.open(out_path) as im:
                        img_size = f"{im.size[0]}x{im.size[1]}"
                except Exception:
                    pass
                return {
                    "success": True,
                    "slide_index": slide_index,
                    "path": str(out_path),
                    "image_size": img_size,
                    "objects_annotated": 0,
                }

            # Annotate
            from PIL import Image, ImageDraw, ImageFont

            with Image.open(tmp_export) as im:
                annotated = im.convert("RGBA").copy()

            draw = ImageDraw.Draw(annotated)
            img_w, img_h = annotated.size
            slide_w = float(self.pres.PageSetup.SlideWidth) or 720.0
            slide_h = float(self.pres.PageSetup.SlideHeight) or 540.0
            scale_x = img_w / slide_w
            scale_y = img_h / slide_h

            try:
                font = ImageFont.truetype("arial.ttf", max(12, int(img_w / 110)))
            except Exception:
                font = ImageFont.load_default()

            box_color = (0, 200, 0, 255)
            bg_color = (255, 235, 50, 230)
            text_color = (0, 0, 0, 255)

            count = 0
            for i in range(1, slide.Shapes.Count + 1):
                try:
                    sh = slide.Shapes(i)
                    x = int(sh.Left * scale_x)
                    y = int(sh.Top * scale_y)
                    w = int(sh.Width * scale_x)
                    h = int(sh.Height * scale_y)
                    draw.rectangle([x, y, x + w, y + h], outline=box_color, width=2)
                    label = f"ID:{int(sh.Id)}"
                    try:
                        tb = draw.textbbox((0, 0), label, font=font)
                        tw, th = tb[2] - tb[0], tb[3] - tb[1]
                    except AttributeError:
                        tw, th = draw.textsize(label, font=font)
                    lx = x
                    ly = y - th - 4 if y - th - 4 >= 0 else y + h + 4
                    draw.rectangle([lx, ly, lx + tw + 4, ly + th + 2], fill=bg_color)
                    draw.text((lx + 2, ly + 1), label, fill=text_color, font=font)
                    count += 1
                except pywintypes.com_error:
                    continue

            annotated.convert("RGB").save(str(out_path), "PNG")
            return {
                "success": True,
                "slide_index": slide_index,
                "path": str(out_path),
                "image_size": f"{img_w}x{img_h}",
                "objects_annotated": count,
            }
        finally:
            try:
                os.unlink(tmp_export)
            except OSError:
                pass

    # ─────────────────────────────────────────────────────────────
    # Connectors — line shapes that bind to other shapes and reroute
    # on move. Small ops; lives here rather than its own mixin.
    # ─────────────────────────────────────────────────────────────

    @staticmethod
    def _resolve_connector_type(connector_type: int | str | None) -> int:
        if connector_type is None:
            return msoConnectorStraight
        if isinstance(connector_type, int):
            return connector_type
        key = connector_type.lower().strip()
        if key not in CONNECTOR_TYPE_MAP:
            raise PowerPointError(
                f"Unknown connector_type '{connector_type}'. "
                f"Choose from {sorted(CONNECTOR_TYPE_MAP)}"
            )
        return CONNECTOR_TYPE_MAP[key]

    def connector_add(
        self,
        slide_index: int,
        connector_type: int | str,
        begin_x: float,
        begin_y: float,
        end_x: float,
        end_y: float,
    ) -> dict:
        """Add a free-floating connector at coordinates. Use connector_connect
        afterwards to bind endpoints to shapes."""
        self._validate_index(slide_index)
        ct = self._resolve_connector_type(connector_type)
        slide = self.pres.Slides(slide_index)
        try:
            shape = slide.Shapes.AddConnector(
                ct,
                float(begin_x),
                float(begin_y),
                float(end_x),
                float(end_y),
            )
        except pywintypes.com_error as e:
            raise PowerPointError(f"AddConnector failed: {e}")
        return {
            "success": True,
            "slide_index": slide_index,
            "shape_id": int(shape.Id),
            "name": shape.Name,
            "connector_type": ct,
            "begin": (float(begin_x), float(begin_y)),
            "end": (float(end_x), float(end_y)),
        }

    def connector_connect(
        self,
        slide_index: int,
        connector_id: int,
        begin_shape_id: int | None = None,
        end_shape_id: int | None = None,
        begin_site: int = 1,
        end_site: int = 3,
    ) -> dict:
        """Bind one or both endpoints of a connector to shapes.

        ``begin_site`` / ``end_site`` are 1-based connection-site indices
        (rectangles expose 4: top/right/bottom/left → 1/2/3/4). Defaults
        wire top→bottom for the common vertical-flow case.
        """
        _, conn = self._require_shape(slide_index, connector_id)
        try:
            cf = conn.ConnectorFormat
        except pywintypes.com_error as e:
            raise PowerPointError(f"Shape id={connector_id} is not a connector: {e}")
        applied: dict = {}
        try:
            if begin_shape_id is not None:
                _, begin_shape = self._require_shape(slide_index, begin_shape_id)
                cf.BeginConnect(begin_shape, int(begin_site))
                applied["begin"] = {
                    "shape_id": int(begin_shape_id),
                    "site": int(begin_site),
                }
            if end_shape_id is not None:
                _, end_shape = self._require_shape(slide_index, end_shape_id)
                cf.EndConnect(end_shape, int(end_site))
                applied["end"] = {
                    "shape_id": int(end_shape_id),
                    "site": int(end_site),
                }
            conn.RerouteConnections()
        except pywintypes.com_error as e:
            raise PowerPointError(f"Connect failed: {e}")
        if not applied:
            raise PowerPointError(
                "connector_connect called with no endpoints. "
                "Pass begin_shape_id and/or end_shape_id."
            )
        return {"success": True, "shape_id": int(connector_id), "applied": applied}

    def connector_disconnect(
        self,
        slide_index: int,
        connector_id: int,
        end: str = "both",
    ) -> dict:
        """Release one or both endpoints. ``end`` ∈ {begin, end, both}."""
        _, conn = self._require_shape(slide_index, connector_id)
        try:
            cf = conn.ConnectorFormat
        except pywintypes.com_error as e:
            raise PowerPointError(f"Shape id={connector_id} is not a connector: {e}")
        key = end.lower().strip()
        if key not in ("begin", "end", "both"):
            raise PowerPointError(f"end must be one of begin/end/both, got '{end}'")
        try:
            if key in ("begin", "both"):
                cf.BeginDisconnect()
            if key in ("end", "both"):
                cf.EndDisconnect()
        except pywintypes.com_error as e:
            raise PowerPointError(f"Disconnect failed: {e}")
        return {"success": True, "shape_id": int(connector_id), "disconnected": key}

    def connector_reroute(self, slide_index: int, connector_id: int) -> dict:
        """Force a connector to recompute its path."""
        _, conn = self._require_shape(slide_index, connector_id)
        try:
            conn.RerouteConnections()
        except pywintypes.com_error as e:
            raise PowerPointError(f"Reroute failed: {e}")
        return {"success": True, "shape_id": int(connector_id)}

    # ─────────────────────────────────────────────────────────────
    # Hyperlinks — URL/email/slide jumps on shapes or paragraphs.
    # Address forms accepted:
    #   http(s)://...           → standard URL
    #   mailto:user@host        → email
    #   slide://N               → internal slide jump
    #   plain path / file:///   → external file
    # ─────────────────────────────────────────────────────────────

    @staticmethod
    def _split_hyperlink_address(address: str) -> tuple[str, str | None]:
        addr = address.strip()
        if addr.lower().startswith("slide://"):
            try:
                idx = int(addr[len("slide://") :])
            except ValueError:
                raise PowerPointError(
                    f"slide:// hyperlink needs an integer slide index, got '{addr}'"
                )
            return "", str(idx)
        return addr, None

    def _hyperlink_target(self, slide_index: int, shape_id: int, paragraph: int | None):
        """Return the COM target — either the shape or a specific paragraph
        TextRange — that owns ``ActionSettings``."""
        _, shape = self._require_shape(slide_index, shape_id)
        if paragraph is None:
            return shape
        try:
            paragraphs = shape.TextFrame.TextRange.Paragraphs()
            count = int(paragraphs.Count)
            if paragraph < 1 or paragraph > count:
                raise PowerPointError(f"paragraph {paragraph} out of range (1-{count})")
            return shape.TextFrame.TextRange.Paragraphs(int(paragraph))
        except pywintypes.com_error as e:
            raise PowerPointError(f"Shape has no text frame: {e}")

    def hyperlink_set(
        self,
        slide_index: int,
        shape_id: int,
        address: str,
        screen_tip: str | None = None,
        paragraph: int | None = None,
    ) -> dict:
        """Attach a hyperlink to a shape (default) or one paragraph of its text."""
        target = self._hyperlink_target(slide_index, shape_id, paragraph)
        addr, sub = self._split_hyperlink_address(address)
        try:
            action = target.ActionSettings(_PP_MOUSE_CLICK)
            action.Action = _PP_ACTION_HYPERLINK
            hl = action.Hyperlink
            hl.Address = addr
            if sub is not None:
                hl.SubAddress = sub
            if screen_tip is not None:
                hl.ScreenTip = screen_tip
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set hyperlink failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "paragraph": paragraph,
            "address": addr,
            "sub_address": sub,
            "screen_tip": screen_tip,
        }

    def hyperlink_remove(
        self,
        slide_index: int,
        shape_id: int,
        paragraph: int | None = None,
    ) -> dict:
        """Remove the hyperlink from a shape (or one paragraph)."""
        target = self._hyperlink_target(slide_index, shape_id, paragraph)
        try:
            action = target.ActionSettings(_PP_MOUSE_CLICK)
            action.Action = _PP_ACTION_NONE
            action.Hyperlink.Address = ""
            action.Hyperlink.SubAddress = ""
        except pywintypes.com_error as e:
            raise PowerPointError(f"Remove hyperlink failed: {e}")
        return {
            "success": True,
            "shape_id": int(shape_id),
            "paragraph": paragraph,
        }

    def hyperlink_get(
        self,
        slide_index: int,
        shape_id: int,
        paragraph: int | None = None,
    ) -> dict:
        """Read the current hyperlink target."""
        target = self._hyperlink_target(slide_index, shape_id, paragraph)
        try:
            action = target.ActionSettings(_PP_MOUSE_CLICK)
            hl = action.Hyperlink
            addr = hl.Address or ""
            sub = hl.SubAddress or ""
            tip = hl.ScreenTip or ""
            has_link = int(action.Action) == _PP_ACTION_HYPERLINK
        except pywintypes.com_error as e:
            raise PowerPointError(f"Read hyperlink failed: {e}")
        return {
            "shape_id": int(shape_id),
            "paragraph": paragraph,
            "has_hyperlink": bool(has_link),
            "address": addr or None,
            "sub_address": sub or None,
            "screen_tip": tip or None,
        }

    # ─────────────────────────────────────────────────────────
    # Slide background + footer — presentation polish at the slide level.
    # ─────────────────────────────────────────────────────────

    def slide_set_background(
        self,
        slide_index: int,
        color=None,
        picture_path: str | None = None,
        follow_master: bool | None = None,
    ) -> dict:
        """Set the background of one slide.

        Modes:
          - ``color="#1A3F6B"`` → solid fill behind everything on the slide.
            Useful for section dividers and KPI-tile backdrops.
          - ``picture_path=<path>`` → full-bleed picture background.
          - ``follow_master=True`` → drop any per-slide override and revert
            to the master/theme background.

        Each mode flips ``slide.FollowMasterBackground=False`` so the slide
        owns its own fill; that's the only way ``slide.Background.Fill`` is
        honored on a per-slide basis.

        FIXME(theme-master): per-slide background overrides fragment the
        deck's theming — if you change the master later, slides with manual
        backgrounds won't follow it. The professional alternative is to
        define a *custom slide layout* on the master (e.g. "Section Divider
        — Blue") and apply it via ``slide_set_layout``. That requires a
        master/theme op surface we haven't built yet; until then, this
        function is the working compromise. Prefer it for one-off section
        dividers, not for deck-wide theming.
        """
        self._validate_index(slide_index)
        slide = self.pres.Slides(slide_index)
        applied: dict = {}
        try:
            if follow_master is True:
                slide.FollowMasterBackground = msoTrue
                applied["follow_master"] = True
                # Nothing else to do; master takes over.
                return {
                    "success": True,
                    "slide_index": slide_index,
                    "applied": applied,
                }

            if picture_path is not None:
                p = Path(picture_path).resolve()
                if not p.exists():
                    raise PowerPointError(f"Background picture not found: {p}")
                slide.FollowMasterBackground = msoFalse
                slide.Background.Fill.UserPicture(str(p))
                applied["picture_path"] = str(p)

            if color is not None:
                rgb = self._parse_color(color)
                slide.FollowMasterBackground = msoFalse
                slide.Background.Fill.Solid()
                slide.Background.Fill.ForeColor.RGB = rgb
                applied["rgb"] = rgb
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set background failed: {e}")

        if not applied:
            raise PowerPointError(
                "slide_set_background called with no attributes. Pass color, "
                "picture_path, or follow_master=True."
            )
        return {
            "success": True,
            "slide_index": slide_index,
            "applied": applied,
        }

    def slide_set_footer(
        self,
        footer: str | None = None,
        slide_number: bool | None = None,
        date: bool | None = None,
        date_format: str | None = None,
        slide_index: int | None = None,
    ) -> dict:
        """Configure the standard slide footer (HeadersFooters). Without
        ``slide_index`` the change applies deck-wide via the master.

        - ``footer="Confidential — 2026 Board Meeting"`` sets the footer
          text. Pass empty string to hide.
        - ``slide_number=True`` turns on automatic slide numbers.
        - ``date=True`` enables the date placeholder. ``date_format`` is
          a strftime-style string forwarded to PowerPoint (ignored if the
          underlying COM build doesn't accept Format).
        """
        if slide_index is not None:
            self._validate_index(slide_index)
            hf = self.pres.Slides(slide_index).HeadersFooters
        else:
            hf = self.pres.SlideMaster.HeadersFooters

        applied: dict = {}
        try:
            if footer is not None:
                if footer:
                    hf.Footer.Visible = msoTrue
                    hf.Footer.Text = footer
                else:
                    hf.Footer.Visible = msoFalse
                applied["footer"] = footer
            if slide_number is not None:
                hf.SlideNumber.Visible = msoTrue if slide_number else msoFalse
                applied["slide_number"] = bool(slide_number)
            if date is not None:
                hf.DateAndTime.Visible = msoTrue if date else msoFalse
                applied["date"] = bool(date)
                if date and date_format is not None:
                    try:
                        # 2 = ppUpdateOptionFixed-ish; use UseFormat=False+Text for fixed.
                        hf.DateAndTime.UseFormat = msoFalse
                        hf.DateAndTime.Text = date_format
                        applied["date_format"] = date_format
                    except pywintypes.com_error:
                        pass
        except pywintypes.com_error as e:
            raise PowerPointError(f"Set footer failed: {e}")

        if not applied:
            raise PowerPointError(
                "slide_set_footer called with no attributes. Pass footer, "
                "slide_number, and/or date."
            )
        return {
            "success": True,
            "scope": "slide" if slide_index is not None else "deck",
            "slide_index": slide_index,
            "applied": applied,
        }
