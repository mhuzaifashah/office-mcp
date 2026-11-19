"""Fast paragraph lookups (`_para_index`) and the table text split used by find."""

from __future__ import annotations

from office_mcp.word.wordcom.model import _para_index
from office_mcp.word.wordcom.model._document import _cell_segments


def test_paragraph_by_index(fake_word):
    doc = fake_word.ActiveDocument
    first = _para_index.paragraph(doc, 1)
    assert first is doc.Paragraphs(1)
    assert _para_index.paragraph(doc, 0) is None
    assert _para_index.paragraph(doc, _para_index.count(doc) + 1) is None


def test_index_at_maps_a_position_to_its_paragraph(fake_word):
    doc = fake_word.ActiveDocument
    for index, para in enumerate(doc.Paragraphs, start=1):
        start, end = int(para.Range.Start), int(para.Range.End)
        if end > start:
            assert _para_index.index_at(doc, start) == index
            assert _para_index.index_at(doc, end - 1) == index


def test_paragraphs_between_yields_the_slice_with_indices(fake_word):
    doc = fake_word.ActiveDocument
    total = _para_index.count(doc)
    rows = list(_para_index.paragraphs_between(doc, 2, total))
    assert [i for i, _ in rows] == list(range(2, total + 1))
    assert [p for _, p in rows] == [doc.Paragraphs(i) for i in range(2, total + 1)]


def test_headings_fall_back_to_a_walk_without_find(fake_word):
    doc = fake_word.ActiveDocument
    expected = [
        (i, int(p.OutlineLevel))
        for i, p in enumerate(doc.Paragraphs, start=1)
        if int(p.OutlineLevel) < 10
    ]
    assert expected, "the fake document should have headings"
    assert [(i, lvl) for i, lvl, _ in _para_index.headings(doc)] == expected
    first = _para_index.headings(doc, first=True)
    assert [(i, lvl) for i, lvl, _ in first] == expected[:1]


def test_texts_splits_the_document_text_when_it_lines_up(fake_word):
    doc = fake_word.ActiveDocument
    doc.Content.Text = "".join(str(p.Range.Text) for p in doc.Paragraphs)
    assert _para_index.texts(doc) == [str(p.Range.Text).rstrip("\r") for p in doc.Paragraphs]


def test_texts_returns_none_when_the_split_cannot_be_trusted(fake_word):
    doc = fake_word.ActiveDocument
    doc.Content.Text = "".join(str(p.Range.Text) for p in doc.Paragraphs).replace("Risks", "X")
    assert _para_index.texts(doc) is None


def test_cell_segments_split_cells_and_skip_row_marks():
    # Two rows of two cells; each cell ends with \r\x07 (one position), and
    # each row with one more \r\x07.
    text = "A1\r\x07B1\r\x07\r\x07\r\x07B2\r\x07\r\x07"
    assert _cell_segments(100, text) == [(100, "A1"), (103, "B1"), (108, "B2")]


def test_cell_segments_count_astral_characters_as_two_positions():
    text = "\U0001f600x\r\x07y\r\x07"
    assert _cell_segments(0, text) == [(0, "\U0001f600x"), (4, "y")]


def test_cell_segments_keep_paragraphs_inside_a_cell():
    assert _cell_segments(0, "one\rtwo\r\x07") == [(0, "one\rtwo")]
