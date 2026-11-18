"""snapshot — page resolution, PDF-export args, file naming, and errors.

The PDF rasteriser (PyMuPDF) is the one piece that needs a heavy optional
dependency and a real PDF on disk, so it's isolated behind
`_snapshot._rasterize_pdf` and faked here. That lets these tests cover all of
wordcom's own logic — which pages get exported, how they're numbered, where
files land — against the mock Word, with no PyMuPDF install. One test at the
bottom exercises the real rasteriser when PyMuPDF *is* available.
"""

from __future__ import annotations

import pytest

from office_mcp.word import wordcom
from office_mcp.word.wordcom.constants import WdExportFormat, WdExportItem, WdExportRange
from office_mcp.word.wordcom.exceptions import OpError, SnapshotError
from office_mcp.word.wordcom.render import _snapshot


@pytest.fixture
def fake_pages(monkeypatch):
    """Make `_rasterize_pdf` return a controllable list of fake PNGs.

    Returns a setter: call `fake_pages(n)` to have the next render yield `n`
    pages (`b"PNG0"`, `b"PNG1"`, …) and to capture the args it was called with.
    """
    state: dict = {"pngs": [b"PNG0"], "calls": []}

    def fake_rasterize(pdf_path, dpi, max_dim=None):
        state["calls"].append({"pdf_path": pdf_path, "dpi": dpi, "max_dim": max_dim})
        return list(state["pngs"])

    monkeypatch.setattr(_snapshot, "_rasterize_pdf", fake_rasterize)

    def configure(n):
        state["pngs"] = [f"PNG{i}".encode() for i in range(n)]
        return state

    configure.state = state  # type: ignore[attr-defined]
    return configure


def _export_kwargs(fake_word):
    fake_word.ActiveDocument.ExportAsFixedFormat.assert_called_once()
    return fake_word.ActiveDocument.ExportAsFixedFormat.call_args.kwargs


# ---------------------------------------------------------------------------
# Library: page selection -> export args + page numbering
# ---------------------------------------------------------------------------


def test_snapshot_all_pages_exports_whole_document(fake_word, fake_pages):
    state = fake_pages(2)
    with wordcom.attach() as word:
        shots = word.documents.active.snapshot()
    kwargs = _export_kwargs(fake_word)
    assert kwargs["ExportFormat"] == int(WdExportFormat.PDF)
    assert kwargs["Range"] == int(WdExportRange.ALL_DOCUMENT)
    assert "From" not in kwargs and "To" not in kwargs
    # Whole-document render numbers pages from 1.
    assert [s.page for s in shots] == [1, 2]
    assert [s.png for s in shots] == state["pngs"]
    assert all(s.path is None for s in shots)


def test_snapshot_default_markup_is_content_only(fake_word, fake_pages):
    fake_pages(1)
    with wordcom.attach() as word:
        word.documents.active.snapshot()
    assert _export_kwargs(fake_word)["Item"] == int(WdExportItem.DOCUMENT_CONTENT)


def test_snapshot_markup_all_exports_with_markup(fake_word, fake_pages):
    fake_pages(1)
    with wordcom.attach() as word:
        word.documents.active.snapshot(markup="all")
    assert _export_kwargs(fake_word)["Item"] == int(WdExportItem.WITH_MARKUP)


def test_anchor_snapshot_markup_all(fake_word, fake_pages):
    fake_pages(1)
    with wordcom.attach() as word:
        doc = word.documents.active
        doc.snapshot_anchor(doc.heading("Introduction"), markup="all")
    assert _export_kwargs(fake_word)["Item"] == int(WdExportItem.WITH_MARKUP)


def test_snapshot_bad_markup_raises(fake_word, fake_pages):
    fake_pages(1)
    with wordcom.attach() as word:
        with pytest.raises(OpError):
            word.documents.active.snapshot(markup="loud")


def test_snapshot_single_page(fake_word, fake_pages):
    fake_pages(1)
    with wordcom.attach() as word:
        shots = word.documents.active.snapshot(pages=3)
    kwargs = _export_kwargs(fake_word)
    assert kwargs["Range"] == int(WdExportRange.FROM_TO)
    assert (kwargs["From"], kwargs["To"]) == (3, 3)
    assert [s.page for s in shots] == [3]


def test_snapshot_page_span_numbers_from_start(fake_word, fake_pages):
    fake_pages(3)
    with wordcom.attach() as word:
        shots = word.documents.active.snapshot(pages=(2, 4))
    kwargs = _export_kwargs(fake_word)
    assert (kwargs["From"], kwargs["To"]) == (2, 4)
    # The exported PDF holds only pages 2-4, so they map back to 2, 3, 4.
    assert [s.page for s in shots] == [2, 3, 4]


def test_snapshot_forwards_dpi(fake_word, fake_pages):
    state = fake_pages(1)
    with wordcom.attach() as word:
        word.documents.active.snapshot(pages=1, dpi=222)
    assert state["calls"][0]["dpi"] == 222


def test_snapshot_default_dpi_is_150(fake_word, fake_pages):
    state = fake_pages(1)
    with wordcom.attach() as word:
        word.documents.active.snapshot(pages=1)
    assert state["calls"][0]["dpi"] == 150


def test_snapshot_default_max_dim_is_none(fake_word, fake_pages):
    state = fake_pages(1)
    with wordcom.attach() as word:
        word.documents.active.snapshot(pages=1)
    assert state["calls"][0]["max_dim"] is None


def test_snapshot_forwards_max_dim(fake_word, fake_pages):
    state = fake_pages(1)
    with wordcom.attach() as word:
        word.documents.active.snapshot(pages=1, max_dim=1000)
    assert state["calls"][0]["max_dim"] == 1000


def test_anchor_snapshot_forwards_max_dim(fake_word, fake_pages):
    state = fake_pages(1)
    with wordcom.attach() as word:
        doc = word.documents.active
        doc.heading("Introduction").snapshot(max_dim=800)
    assert state["calls"][0]["max_dim"] == 800


@pytest.mark.parametrize("bad", [0, -5])
def test_snapshot_rejects_bad_max_dim(fake_word, fake_pages, bad):
    fake_pages(1)
    with wordcom.attach() as word:
        with pytest.raises(OpError):
            word.documents.active.snapshot(pages=1, max_dim=bad)


@pytest.mark.parametrize("bad", [0, -1, (0, 2), (3, 2), (1, 2, 3)])
def test_snapshot_rejects_bad_pages(fake_word, fake_pages, bad):
    fake_pages(1)
    with wordcom.attach() as word:
        with pytest.raises(ValueError):
            word.documents.active.snapshot(pages=bad)


# ---------------------------------------------------------------------------
# Library: writing files via out=
# ---------------------------------------------------------------------------


def test_snapshot_out_single_page_writes_exact_path(fake_word, fake_pages, tmp_path):
    state = fake_pages(1)
    out = tmp_path / "section.png"
    with wordcom.attach() as word:
        shots = word.documents.active.snapshot(out, pages=3)
    assert out.read_bytes() == state["pngs"][0]
    assert shots[0].path == out


def test_snapshot_out_multi_page_suffixes_page_number(fake_word, fake_pages, tmp_path):
    state = fake_pages(2)
    out = tmp_path / "doc.png"
    with wordcom.attach() as word:
        shots = word.documents.active.snapshot(out, pages=(4, 5))
    assert (tmp_path / "doc-p4.png").read_bytes() == state["pngs"][0]
    assert (tmp_path / "doc-p5.png").read_bytes() == state["pngs"][1]
    assert [s.path.name for s in shots] == ["doc-p4.png", "doc-p5.png"]
    # The original bare path is not written when there's more than one page.
    assert not out.exists()


# ---------------------------------------------------------------------------
# Library: anchor snapshots
# ---------------------------------------------------------------------------


def test_anchor_snapshot_uses_anchor_page_span(fake_word, fake_pages):
    fake_pages(1)
    with wordcom.attach() as word:
        doc = word.documents.active
        shots = doc.bookmarks["Address"].snapshot()
    kwargs = _export_kwargs(fake_word)
    # The fake document is a single page, so the anchor resolves to page 1.
    assert kwargs["Range"] == int(WdExportRange.FROM_TO)
    assert (kwargs["From"], kwargs["To"]) == (1, 1)
    assert [s.page for s in shots] == [1]


def test_heading_snapshot_renders_its_section(fake_word, fake_pages):
    """A heading anchor expands to its section span (start of heading -> body end)."""
    fake_pages(1)
    with wordcom.attach() as word:
        doc = word.documents.active
        shots = doc.heading("Introduction").snapshot()
    kwargs = _export_kwargs(fake_word)
    assert (kwargs["From"], kwargs["To"]) == (1, 1)
    assert len(shots) == 1


def test_anchor_snapshot_writes_file(fake_word, fake_pages, tmp_path):
    state = fake_pages(1)
    out = tmp_path / "anchor.png"
    with wordcom.attach() as word:
        doc = word.documents.active
        shots = doc.bookmarks["Address"].snapshot(out)
    assert out.read_bytes() == state["pngs"][0]
    assert shots[0].path == out


# ---------------------------------------------------------------------------
# Library: missing PyMuPDF -> SnapshotError
# ---------------------------------------------------------------------------


def test_snapshot_without_pymupdf_raises(fake_word, monkeypatch):
    def boom():
        raise SnapshotError("PyMuPDF missing")

    # Don't fake _rasterize_pdf here — let the real one run and hit the import.
    monkeypatch.setattr(_snapshot, "_import_fitz", boom)
    with wordcom.attach() as word:
        with pytest.raises(SnapshotError):
            word.documents.active.snapshot(pages=1)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# Real rasteriser — only when PyMuPDF is actually installed
# ---------------------------------------------------------------------------


def test_rasterize_pdf_real(tmp_path):
    fitz = pytest.importorskip("pymupdf")
    pdf_path = tmp_path / "two.pdf"
    doc = fitz.open()
    doc.new_page()
    doc.new_page()
    doc.save(str(pdf_path))
    doc.close()

    pngs = _snapshot._rasterize_pdf(str(pdf_path), dpi=72)
    assert len(pngs) == 2
    assert all(p.startswith(b"\x89PNG\r\n\x1a\n") for p in pngs)


def _png_dims(png: bytes) -> tuple[int, int]:
    """(width, height) from a PNG's IHDR chunk (big-endian, at byte offset 16)."""
    return int.from_bytes(png[16:20], "big"), int.from_bytes(png[20:24], "big")


def test_rasterize_pdf_max_dim_caps_long_edge(tmp_path):
    fitz = pytest.importorskip("pymupdf")
    pdf_path = tmp_path / "letter.pdf"
    doc = fitz.open()
    doc.new_page(width=612, height=792)  # US Letter (long edge 792 pt)
    doc.save(str(pdf_path))
    doc.close()

    # dpi=300 alone would render the long edge at 792/72*300 = 3300 px; the cap
    # pulls it down to ~300 px (zoom = 300/792), and only ever lowers resolution.
    pngs = _snapshot._rasterize_pdf(str(pdf_path), dpi=300, max_dim=300)
    w, h = _png_dims(pngs[0])
    assert max(w, h) <= 300
    assert max(w, h) >= 298  # rounding only


def test_rasterize_pdf_max_dim_does_not_upscale(tmp_path):
    fitz = pytest.importorskip("pymupdf")
    pdf_path = tmp_path / "letter.pdf"
    doc = fitz.open()
    doc.new_page(width=612, height=792)
    doc.save(str(pdf_path))
    doc.close()

    # A generous cap above the dpi-driven size leaves the dpi resolution intact.
    capped = _snapshot._rasterize_pdf(str(pdf_path), dpi=72, max_dim=5000)
    plain = _snapshot._rasterize_pdf(str(pdf_path), dpi=72)
    assert _png_dims(capped[0]) == _png_dims(plain[0])
