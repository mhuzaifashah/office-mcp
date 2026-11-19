"""Smoke tests: the fast paragraph lookups agree with a full walk in real Word.

Skipped by default; run with `pytest -m smoke` while Word is running. The fake
document can't model Find or table text, so these checks need real Word.
"""

from __future__ import annotations

import contextlib

import pytest

from office_mcp.word.wordcom.model import _para_index
from office_mcp.word.wordcom.model._document import _cell_segments, _cell_spans_by_paragraph

pytestmark = pytest.mark.smoke

_MARKDOWN = (
    "# One\n\nIntro text.\n\n## One A\n\nBody **bold** and *italic*.\n\n"
    "## One B\n\nMore.\n\n# Two\n\n### Two deep\n\nEnd."
)


@pytest.fixture
def doc(real_word):
    real_word.com.Documents.Add()
    d = real_word.documents.active
    with d.edit("smoke: para index"):
        d.anchor_by_id("start").insert_markdown(_MARKDOWN)
        # A 2x2 table after "Body" (insert_markdown makes no tables).
        com = d.com
        body = next(p for p in com.Paragraphs if str(p.Range.Text).startswith("Body"))
        at = com.Range(int(body.Range.End), int(body.Range.End))
        at.InsertParagraphAfter()
        table = com.Tables.Add(com.Range(int(body.Range.End), int(body.Range.End)), 2, 2)
        for (r, c), text in {(1, 1): "a", (1, 2): "b", (2, 1): "c", (2, 2): "\U0001f600d"}.items():
            table.Cell(r, c).Range.Text = text
    try:
        yield d
    finally:
        with contextlib.suppress(Exception):
            d.com.Close(SaveChanges=0)


def _walk(doc_com):
    return [
        (i, int(p.OutlineLevel))
        for i, p in enumerate(doc_com.Paragraphs, start=1)
        if int(p.OutlineLevel) < 10
    ]


def test_find_headings_match_the_walk(doc):
    com = doc.com
    expected = _walk(com)
    assert [(i, lvl) for i, lvl, _ in _para_index.headings(com)] == expected
    for max_level in (1, 2):
        got = [(i, lvl) for i, lvl, _ in _para_index.headings(com, max_level=max_level)]
        assert got == [h for h in expected if h[1] <= max_level]


def test_next_heading_after_each_heading(doc):
    com = doc.com
    expected = _walk(com)
    for index, level in expected:
        after = int(_para_index.paragraph(com, index).Range.End)
        nxt = _para_index.headings(com, start=after, max_level=level, first=True)
        want = next(((i, lv) for i, lv in expected if i > index and lv <= level), None)
        assert [(i, lv) for i, lv, _ in nxt] == ([want] if want else [])


def test_index_at_and_slices_match_the_walk(doc):
    com = doc.com
    spans = [(int(p.Range.Start), int(p.Range.End)) for p in com.Paragraphs]
    for index, (start, end) in enumerate(spans, start=1):
        assert _para_index.index_at(com, start) == index
        assert _para_index.index_at(com, end - 1) == index
    rows = list(_para_index.paragraphs_between(com, 2, len(spans)))
    assert [(int(p.Range.Start), int(p.Range.End)) for _, p in rows] == spans[1:]


def test_texts_match_each_paragraph(doc):
    com = doc.com
    texts = _para_index.texts(com)
    assert texts is not None
    assert texts == [str(p.Range.Text).rstrip("\r\x07") for p in com.Paragraphs]


def test_table_text_split_matches_the_cell_walk(doc):
    com = doc.com
    table = com.Tables(1).Range
    start, end = int(table.Start), int(table.End)
    fast = _cell_segments(start, str(com.Range(start, end).Text))
    slow = [
        (a, str(com.Range(a, b).Text).rstrip("\r\n\x07"))
        for a, b, _ in _cell_spans_by_paragraph(com, start, end)
    ]
    assert fast == [s for s in slow if s[1]]


def test_find_inside_a_table_cell(doc):
    hits = doc.find("d")
    cell = [h for h in hits if h["text"] == "d"]
    assert cell and str(doc.com.Range(cell[0]["start"], cell[0]["end"]).Text) == "d"
