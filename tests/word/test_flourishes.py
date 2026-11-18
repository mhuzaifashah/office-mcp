"""Publishing flourishes — `doc.set_watermark` / `remove_watermark` and
`anchor.insert_text_box`.

Floating shapes are COM-heavy and version-sensitive; these tests pin the wiring
against the `fake_word` MagicMock (the right shapes get added to the right story,
named the watermark convention, and removed again). Visual correctness is a
smoke / live-Word concern.
"""

from __future__ import annotations

import pytest

from office_mcp.word import wordcom
from office_mcp.word.wordcom.exceptions import OpError


def _primary_header_shapes(fake_word):
    return fake_word.ActiveDocument.Sections(1).Headers(1).Shapes


# --- watermark ------------------------------------------------------------------


def test_set_watermark_adds_wordart_to_header(fake_word):
    with wordcom.attach() as word:
        n = word.documents.active.set_watermark("DRAFT")
    assert n == 1
    shapes = _primary_header_shapes(fake_word)
    assert shapes.Count == 1
    shape = shapes(1)
    assert shape.Text == "DRAFT"
    assert shape.Name.startswith("PowerPlusWaterMarkObject")


def test_set_watermark_replaces_prior(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        doc.set_watermark("DRAFT")
        doc.set_watermark("FINAL")
    shapes = _primary_header_shapes(fake_word)
    # Calling twice doesn't stack — the first watermark was cleared.
    assert shapes.Count == 1
    assert shapes(1).Text == "FINAL"


def test_remove_watermark_clears_and_counts(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        doc.set_watermark("CONFIDENTIAL")
        removed = doc.remove_watermark()
    assert removed == 1
    assert _primary_header_shapes(fake_word).Count == 0


def test_remove_watermark_none_present_is_zero(fake_word):
    with wordcom.attach() as word:
        assert word.documents.active.remove_watermark() == 0


def test_set_watermark_bad_layout_raises(fake_word):
    with wordcom.attach() as word:
        with pytest.raises(OpError):
            word.documents.active.set_watermark("X", layout="sideways")


# --- text box -------------------------------------------------------------------


def test_insert_text_box_adds_floating_shape(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_text_box("Pull quote", width="3in")
    shapes = fake_word.ActiveDocument.Shapes
    assert shapes.Count == 1
    # The text landed in the box's text frame.
    assert shapes(1).TextFrame.TextRange.Text == "Pull quote"


def test_insert_text_box_unknown_wrap_raises(fake_word):
    with wordcom.attach() as word:
        with pytest.raises(ValueError):
            word.documents.active.bookmarks["Address"].insert_text_box("x", wrap="inline")
