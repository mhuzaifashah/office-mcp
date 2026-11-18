"""Proofing — `doc.proofing()`, the `proofing` CLI command, MCP read.

Round-trips against `fake_word`, seeded with one spelling error ("teh"), one
grammar error ("is are"), and a handful of readability statistics.
"""

from __future__ import annotations

import pytest

from office_mcp.word import wordcom


def _make_doc(**kwargs):
    # Build a bespoke document with whatever proofing seeds a test needs.
    from tests.word.conftest import _make_application, _make_document

    return _make_application([_make_document(**kwargs)])


# --- the reader ----------------------------------------------------------------


def test_proofing_shape(fake_word):
    with wordcom.attach() as word:
        data = word.documents.active.proofing()
    assert data["spelling"]["count"] == 1
    assert data["spelling"]["errors"][0]["text"] == "teh"
    assert data["spelling"]["errors"][0]["anchor_id"] == "range:14-17"
    assert data["grammar"]["count"] == 1
    assert data["grammar"]["errors"][0]["text"] == "is are"


def test_readability_slugged(fake_word):
    with wordcom.attach() as word:
        read = word.documents.active.proofing()["readability"]
    assert read["flesch_reading_ease"] == 65.5
    assert read["flesch_kincaid_grade_level"] == 7.2
    assert read["passive_sentences"] == 12.0
    assert read["words_per_sentence"] == 15.3


def test_proofing_unavailable_is_graceful(monkeypatch):
    # A checker that raises (proofing off / protected doc) -> count None, [].
    from tests.word.conftest import _FakeProofErrors, _FakeReadability

    app = _make_doc()
    doc = app.ActiveDocument
    doc.SpellingErrors = _FakeProofErrors([], raises=True)
    doc.GrammaticalErrors = _FakeProofErrors([], raises=True)
    doc.ReadabilityStatistics = _FakeReadability({}, raises=True)
    from office_mcp.word.wordcom import _com

    monkeypatch.setattr(_com, "get_active_word", lambda: app)
    monkeypatch.setattr(_com, "launch_word", lambda visible=True: app)
    with wordcom.attach() as word:
        data = word.documents.active.proofing()
    assert data["spelling"]["count"] is None
    assert data["spelling"]["errors"] == []
    assert data["readability"] == {}


# --- CLI -----------------------------------------------------------------------


# --- MCP -----------------------------------------------------------------------


def test_mcp_proofing(fake_word):
    pytest.importorskip("mcp")
    from office_mcp.word.wordcom._dispatch import _read_impl
    from office_mcp.word.wordcom._worker import InlineWorker

    data = _read_impl(InlineWorker(), "proofing", {})
    assert data["spelling"]["count"] == 1
    assert data["readability"]["flesch_kincaid_grade_level"] == 7.2
