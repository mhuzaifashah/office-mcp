"""Render a sheet or range as PNG pages, the way it would print.

Excel exports the range to a temporary PDF, which PyMuPDF rasterises. This
never touches the clipboard and adds nothing to the workbook. For the export
the sheet's page setup briefly prints gridlines and row/column headings (so
the agent can read cell addresses off the picture) and fits the range to one
page wide; the previous settings and the workbook's saved state are restored
afterwards.
"""

from __future__ import annotations

import os
import shutil
import tempfile
from typing import Any

from .dispatch import _sheet_name
from .session import controller_for

__all__ = ["render"]

# PageSetup properties the export changes, restored afterwards in this order
# (Zoom last: setting a number turns fit-to-pages off).
_PAGE_SETUP_FIELDS = ("PrintGridlines", "PrintHeadings", "FitToPagesWide", "FitToPagesTall", "Zoom")


def _import_fitz() -> Any:
    try:
        import pymupdf  # type: ignore[import-not-found]

        return pymupdf
    except ImportError:  # pragma: no cover - older PyMuPDF
        import fitz  # type: ignore[import-not-found]

        return fitz


def _rasterize(pdf_path: str, dpi: int, max_dim: int | None) -> list[bytes]:
    fitz = _import_fitz()
    pages: list[bytes] = []
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            zoom = dpi / 72.0
            png = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom)).tobytes("png")
            pages.append(_trim(png, max_dim))
    return pages


def _trim(png: bytes, max_dim: int | None, pad: int = 12) -> bytes:
    """Crop the page's blank margin, then cap the long edge at `max_dim` pixels.

    A small range on a letter-size page is mostly white; cropping before the
    resize keeps the cells legible within the same image budget.
    """
    from io import BytesIO  # noqa: PLC0415

    from PIL import Image, ImageChops  # noqa: PLC0415

    with Image.open(BytesIO(png)) as img:
        rgb = img.convert("RGB")
        box = ImageChops.difference(rgb, Image.new("RGB", rgb.size, "white")).getbbox()
        if box is not None:
            left, top, right, bottom = box
            rgb = rgb.crop(
                (
                    max(0, left - pad),
                    max(0, top - pad),
                    min(rgb.width, right + pad),
                    min(rgb.height, bottom + pad),
                )
            )
        if max_dim and max(rgb.size) > max_dim:
            scale = max_dim / max(rgb.size)
            size = (max(1, round(rgb.width * scale)), max(1, round(rgb.height * scale)))
            rgb = rgb.resize(size, Image.LANCZOS)
        out = BytesIO()
        rgb.save(out, format="PNG", optimize=True)
        return out.getvalue()


def _set_print_communication(app_api: Any, on: bool) -> None:
    # Batches page-setup changes into one printer round trip (Excel 2010+).
    try:
        app_api.PrintCommunication = on
    except Exception:  # noqa: BLE001
        pass


def render(
    book: Any,
    *,
    sheet: str | None = None,
    address: str | None = None,
    dpi: int = 150,
    max_dim: int | None = 1200,
    headings: bool = True,
) -> list[tuple[int, bytes]]:
    """`[(page_number, png_bytes), ...]` for the range (default: the used range)."""
    ctrl = controller_for(book)
    ws = book.sheets[_sheet_name(ctrl, sheet)]
    rng = ws.range(address).api if address else ws.api.UsedRange
    setup = ws.api.PageSetup
    app_api = book.app.api
    was_saved = bool(book.api.Saved)
    before: dict[str, Any] = {}
    tmp = tempfile.mkdtemp(prefix="office-mcp-xl-")
    pdf_path = os.path.join(tmp, "range.pdf")
    try:
        _set_print_communication(app_api, False)
        try:
            for name in _PAGE_SETUP_FIELDS:
                before[name] = getattr(setup, name)
            setup.PrintGridlines = True
            setup.PrintHeadings = bool(headings)
            setup.Zoom = False
            setup.FitToPagesWide = 1
            setup.FitToPagesTall = False
        finally:
            _set_print_communication(app_api, True)
        rng.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=pdf_path,
            Quality=0,
            IncludeDocProperties=False,
            IgnorePrintAreas=True,
            OpenAfterPublish=False,
        )
        pngs = _rasterize(pdf_path, dpi, max_dim)
    finally:
        if before:
            _set_print_communication(app_api, False)
            try:
                for name in _PAGE_SETUP_FIELDS:
                    if name == "Zoom" and before[name] is False:
                        continue  # fit-to-pages was on; its own fields restore it
                    try:
                        setattr(setup, name, before[name])
                    except Exception:  # noqa: BLE001 - keep restoring the rest
                        pass
            finally:
                _set_print_communication(app_api, True)
        if was_saved:
            try:
                book.api.Saved = True
            except Exception:  # noqa: BLE001
                pass
        shutil.rmtree(tmp, ignore_errors=True)
    return list(enumerate(pngs, start=1))
