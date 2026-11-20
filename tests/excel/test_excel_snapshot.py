"""Snapshot pages are cropped to their content and capped in size."""

from __future__ import annotations

from io import BytesIO

from PIL import Image, ImageDraw

from office_mcp.excel.snapshot import _trim


def _page(size=(1275, 1650), box=(100, 100, 400, 250)) -> bytes:
    img = Image.new("RGB", size, "white")
    if box is not None:
        ImageDraw.Draw(img).rectangle(box, fill="black")
    out = BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


def _size(png: bytes) -> tuple[int, int]:
    return Image.open(BytesIO(png)).size


def test_blank_margins_are_cropped():
    assert _size(_trim(_page(), None)) == (301 + 24, 151 + 24)


def test_the_long_edge_is_capped_after_cropping():
    w, h = _size(_trim(_page(box=(0, 0, 1274, 400)), 600))
    assert max(w, h) == 600


def test_a_blank_page_is_kept_whole_but_capped():
    assert max(_size(_trim(_page(box=None), 500))) == 500
