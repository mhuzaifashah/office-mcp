"""Content control post-creation setters — `ContentControl.set_properties` /
`set_items`, the `set_cc_properties` / `set_cc_items` ops, and the CLI.

`fake_word` seeds one control (title "Signatory", tag "sig", a rich_text by
default — Type 0). Tests that need a dropdown build their own document seeded
with a combo_box/dropdown control carrying `kind` + `items`.
"""

from __future__ import annotations

import pytest

from office_mcp.word import wordcom
from office_mcp.word.wordcom.constants import WdContentControlType
from office_mcp.word.wordcom.exceptions import OpError
from office_mcp.word.wordcom.runtime._ops import run_batch


def _make_app(**kwargs):
    from tests.word.conftest import _make_application, _make_document

    return _make_application([_make_document(**kwargs)])


def _dropdown_app(monkeypatch, items=("High", "Low")):
    """An app whose only control is a dropdown 'Priority' with `items` seeded."""
    app = _make_app(
        content_controls=[
            {
                "title": "Priority",
                "tag": "prio",
                "kind": int(WdContentControlType.DROPDOWN_LIST),
                "items": list(items),
                "start": 0,
                "end": 5,
            }
        ],
        content="Hello\r",
        paragraphs=[{"level": 10, "text": "Hello", "start": 0, "end": 5}],
    )
    from office_mcp.word.wordcom.runtime import _com

    monkeypatch.setattr(_com, "get_active_word", lambda: app)
    monkeypatch.setattr(_com, "launch_word", lambda visible=True: app)
    return app


# --- set_properties ------------------------------------------------------------


def test_set_properties_writes_and_chains(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        cc = doc.content_controls["Signatory"]
        with doc.edit("cc"):
            ret = cc.set_properties(title="Signer", lock_contents=True)
        assert ret is cc
        live = cc._cc()
        assert live.Title == "Signer"
        assert live.LockContents is True
        # rename keeps the anchor id honest
        assert cc.anchor_id == "cc:Signer"


def test_set_properties_title_none_clears(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        cc = doc.content_controls["Signatory"]
        with doc.edit("cc"):
            cc.set_properties(title=None)
        assert cc._cc().Title == ""


def test_set_properties_omitted_left_untouched(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        cc = doc.content_controls["Signatory"]
        with doc.edit("cc"):
            cc.set_properties(tag="newtag")
        live = cc._cc()
        assert live.Tag == "newtag"
        assert live.Title == "Signatory"  # untouched


# --- set_items -----------------------------------------------------------------


def test_set_items_replaces_list(monkeypatch):
    _dropdown_app(monkeypatch, items=("High", "Low"))
    with wordcom.attach() as word:
        doc = word.documents.active
        cc = doc.content_controls["Priority"]
        with doc.edit("cc"):
            cc.set_items(["A", {"text": "Bee", "value": "B"}, "C"])
        entries = list(cc._cc().DropdownListEntries)
    assert [(e.Text, e.Value) for e in entries] == [("A", "A"), ("Bee", "B"), ("C", "C")]


def test_set_items_on_non_list_kind_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        cc = doc.content_controls["Signatory"]  # rich_text (Type 0)
        with pytest.raises(OpError):
            cc.set_items(["a"])


# --- exec ops ------------------------------------------------------------------


def test_op_set_cc_properties(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_cc_properties", "anchor_id": "cc:Signatory", "tag": "t2"}],
            label="test",
        )
    assert exc is None and result["ok"] is True


def test_op_set_cc_properties_no_fields_fails(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "set_cc_properties", "anchor_id": "cc:Signatory"}], label="test"
        )
    assert isinstance(exc, OpError) and result["ok"] is False


def test_op_set_cc_items(monkeypatch):
    _dropdown_app(monkeypatch)
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_cc_items", "anchor_id": "cc:Priority", "items": ["X", "Y"]}],
            label="test",
        )
        entries = list(doc.content_controls["Priority"]._cc().DropdownListEntries)
    assert exc is None and result["ok"] is True
    assert [e.Text for e in entries] == ["X", "Y"]


def test_op_set_cc_on_wrong_anchor_fails(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "set_cc_properties", "anchor_id": "para:1", "title": "x"}], label="test"
        )
    assert isinstance(exc, OpError) and result["ok"] is False


# --- CLI -----------------------------------------------------------------------


# --- MCP -----------------------------------------------------------------------


def test_mcp_build_set_cc_properties(fake_word):
    pytest.importorskip("mcp")
    from office_mcp.word.wordcom.runtime._dispatch import write_command_to_op

    op = write_command_to_op("set_cc_properties", {"anchor_id": "cc:Signatory", "title": "Z"})
    assert op == {"op": "set_cc_properties", "anchor_id": "cc:Signatory", "title": "Z"}


def test_mcp_build_set_cc_properties_no_fields_raises(fake_word):
    pytest.importorskip("mcp")
    from office_mcp.word.wordcom.runtime._dispatch import write_command_to_op

    with pytest.raises(OpError):
        write_command_to_op("set_cc_properties", {"anchor_id": "cc:Signatory"})


def test_mcp_build_set_cc_items(fake_word):
    pytest.importorskip("mcp")
    from office_mcp.word.wordcom.runtime._dispatch import write_command_to_op

    op = write_command_to_op("set_cc_items", {"anchor_id": "cc:Priority", "items": ["a", "b"]})
    assert op == {"op": "set_cc_items", "anchor_id": "cc:Priority", "items": ["a", "b"]}
