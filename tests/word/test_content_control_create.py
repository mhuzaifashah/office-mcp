"""Content control creation — `Anchor.insert_content_control`,
`ContentControlCollection.add`, the `create_content_control` op, and the CLI.

`doc.ContentControls.Add` is a recording MagicMock (see `_FakeContentControls`),
so the `WdContentControlType` wordcom passes is assertable without real Word.
"""

from __future__ import annotations

import pytest

from office_mcp.word import wordcom
from office_mcp.word.wordcom.constants import WdContentControlType
from office_mcp.word.wordcom.exceptions import OpError
from office_mcp.word.wordcom.runtime._ops import run_batch


def _cc_add(fake_word):
    return fake_word.ActiveDocument.ContentControls.Add


# --- the Anchor method ---------------------------------------------------------


def test_insert_content_control_maps_kind_and_returns_named_anchor(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("cc"):
            cc = doc.range(0, 5).insert_content_control("dropdown", title="Priority")
        assert isinstance(cc, wordcom.ContentControl)
        assert cc.anchor_id == "cc:Priority"
    # Type is the first positional arg to ContentControls.Add.
    assert _cc_add(fake_word).call_args.args[0] == int(WdContentControlType.DROPDOWN_LIST)


def test_insert_content_control_default_kind_is_rich_text(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("cc"):
            doc.range(0, 5).insert_content_control()
    assert _cc_add(fake_word).call_args.args[0] == int(WdContentControlType.RICH_TEXT)


def test_insert_content_control_unnamed_returns_usable_wrapper(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("cc"):
            cc = doc.range(0, 0).insert_content_control("date", where="after")
        # No title/tag => empty name, but the wrapper caches the live control and
        # still resolves (text read goes through the cached COM object).
        assert cc.name == ""
        assert cc.text == ""


def test_insert_content_control_items_populate_dropdown(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("cc"):
            doc.range(0, 0).insert_content_control(
                "combo_box",
                title="Pick",
                items=["High", {"text": "Medium", "value": "MED"}, "Low"],
            )
        created = _cc_add(fake_word).side_effect.__self__._items[-1]
    entries = created.DropdownListEntries.Add
    assert entries.call_count == 3
    assert entries.call_args_list[0].args == ("High", "High")
    assert entries.call_args_list[1].args == ("Medium", "MED")


def test_insert_content_control_lock_flags(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("cc"):
            doc.range(0, 0).insert_content_control(
                "text", title="Locked", lock_contents=True, lock_control=True
            )
        created = _cc_add(fake_word).side_effect.__self__._items[-1]
    assert created.LockContents is True
    assert created.LockContentControl is True


def test_insert_content_control_items_on_non_list_kind_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.range(0, 0).insert_content_control("rich_text", items=["a"])


def test_insert_content_control_bad_kind_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.range(0, 0).insert_content_control("banana")


def test_insert_content_control_bad_where_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.range(0, 0).insert_content_control("text", where="sideways")


def test_collection_add_delegates(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("cc"):
            cc = doc.content_controls.add("range:0-5", "checkbox", tag="agree")
        assert cc.anchor_id == "cc:agree"
    assert _cc_add(fake_word).call_args.args[0] == int(WdContentControlType.CHECKBOX)


# --- exec op -------------------------------------------------------------------


def test_exec_create_content_control_reports_anchor_id(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {
                    "op": "create_content_control",
                    "anchor_id": "range:0-5",
                    "kind": "dropdown",
                    "title": "Priority",
                    "items": ["A", "B"],
                }
            ],
            label="test",
        )
    assert exc is None and result["ok"] is True
    out = result["outputs"][0]
    assert out["content_control"] == "Priority"
    assert out["anchor_id"] == "cc:Priority"


def test_exec_create_content_control_unnamed_reports_none(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "create_content_control", "anchor_id": "range:0-0", "kind": "date"}],
            label="test",
        )
    assert exc is None
    out = result["outputs"][0]
    assert out["content_control"] is None
    assert out["anchor_id"] is None


# --- CLI -----------------------------------------------------------------------
