"""replace_section_body must not delete outside the section body.

Covers two COM behaviours the generic fake doesn't model on its own:
- `Range.Delete()` on a zero-width range deletes the following character.
- A range reaching `Content.End` includes the undeletable final paragraph mark.
"""

from __future__ import annotations

from office_mcp.word import wordcom
from office_mcp.word.wordcom import _com as _com_module
from tests.word.conftest import _make_application, _make_document


def _attach_doc(monkeypatch, paragraphs, content):
    doc_com = _make_document(paragraphs=paragraphs, content=content)
    deleted: list[tuple[int, int]] = []
    make_range = doc_com.Range.side_effect

    def recording_range(start=None, end=None):
        rng = make_range(start, end)
        if start is not None and end is not None:
            rng.Delete.side_effect = lambda s=int(start), e=int(end): deleted.append((s, e))
        return rng

    doc_com.Range.side_effect = recording_range
    app = _make_application([doc_com])
    monkeypatch.setattr(_com_module, "get_active_word", lambda: app)
    return doc_com, deleted


def test_empty_section_deletes_nothing(monkeypatch):
    # "A" heading directly followed by sibling "B": section body is zero-width.
    _, deleted = _attach_doc(
        monkeypatch,
        paragraphs=[
            {"level": 1, "text": "A", "start": 0, "end": 2},
            {"level": 1, "text": "B", "start": 2, "end": 4},
        ],
        content="A\rB\r",
    )
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("rewrite"):
            doc.heading("A").replace_section_body(["new"])
    assert deleted == []


def test_last_section_delete_stops_before_final_mark(monkeypatch):
    doc_com, deleted = _attach_doc(
        monkeypatch,
        paragraphs=[
            {"level": 1, "text": "Head", "start": 0, "end": 5},
            {"level": 10, "text": "Body", "start": 5, "end": 10},
        ],
        content="Head\rBody\r",
    )
    content_end = int(doc_com.Content.End)
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("rewrite"):
            doc.heading("Head").replace_section_body(["new"])
    assert deleted, "section body should be deleted"
    for _start, end in deleted:
        assert end <= content_end - 1
