"""Citations & bibliography — `Anchor.insert_citation` / `insert_bibliography`,
`Document.add_bibliography`, the `insert_citation` / `insert_bibliography` ops, CLI.

Both insert a Word field via the EMPTY raw-code path, so — like `test_fields.py` —
assertions read `<anchor>.Range.Duplicate.Fields.Add` (the field code is the 3rd
positional arg; EMPTY=-1 is the 2nd).
"""

from __future__ import annotations

import pytest

from office_mcp.word import wordcom
from office_mcp.word.wordcom._ops import run_batch
from office_mcp.word.wordcom.constants import WdFieldType
from office_mcp.word.wordcom.exceptions import OpError


def _field_add(fake_word):
    """The `Fields.Add` mock for a field inserted at bookmark:Address."""
    return fake_word.ActiveDocument.Bookmarks("Address").Range.Duplicate.Fields.Add


# --- insert_citation -----------------------------------------------------------


def test_insert_citation_basic_code(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("cite"):
            cit = doc.bookmarks["Address"].insert_citation("Smith2020")
        assert isinstance(cit, wordcom.Citation)
    call = _field_add(fake_word).call_args
    assert call.args[1] == int(WdFieldType.EMPTY)
    assert call.args[2] == "CITATION Smith2020 \\l 1033"


def test_insert_citation_all_switches(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("cite"):
            doc.bookmarks["Address"].insert_citation(
                "Smith2020",
                pages="15",
                prefix="see ",
                suffix=", at 12",
                volume="5",
                suppress_author=True,
                suppress_year=True,
                suppress_title=True,
                locale=2057,
            )
    code = _field_add(fake_word).call_args.args[2]
    assert code == 'CITATION Smith2020 \\l 2057 \\p "15" \\v 5 \\f "see " \\s ", at 12" \\n \\y \\t'


def test_insert_citation_empty_tag_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].insert_citation("   ")


def test_insert_citation_bad_where_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].insert_citation("Smith2020", where="sideways")


# --- insert_bibliography -------------------------------------------------------


def test_insert_bibliography_code(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("bib"):
            bib = doc.bookmarks["Address"].insert_bibliography()
        assert isinstance(bib, wordcom.Bibliography)
    assert _field_add(fake_word).call_args.args[2] == "BIBLIOGRAPHY"


def test_add_bibliography_returns_bibliography(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("bib"):
            bib = doc.add_bibliography()
        assert isinstance(bib, wordcom.Bibliography)


def test_bibliography_update_calls_com(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("bib"):
            bib = doc.bookmarks["Address"].insert_bibliography()
        bib.update()
    assert bib.com.Update.call_count == 1


# --- exec ops ------------------------------------------------------------------


def test_exec_insert_citation(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {
                    "op": "insert_citation",
                    "anchor_id": "bookmark:Address",
                    "tag": "Smith2020",
                    "pages": "15",
                }
            ],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert result["outputs"][0]["citation"] == "Smith2020"
    assert '\\p "15"' in _field_add(fake_word).call_args.args[2]


def test_exec_insert_bibliography(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "insert_bibliography", "anchor_id": "end"}], label="test"
        )
    assert exc is None
    assert result["outputs"][0]["bibliography"] is True


# --- CLI -----------------------------------------------------------------------
