"""Table autofit — `Table.autofit`, the `autofit_table` op, `table autofit` CLI, MCP.

Round-trips against `fake_word`, whose `_FakeTable` records `AllowAutoFit` and
the last `AutoFitBehavior` argument.
"""

from __future__ import annotations

import pytest

from office_mcp.word import wordcom
from office_mcp.word.wordcom.constants import WdAutoFitBehavior
from office_mcp.word.wordcom.exceptions import OpError
from office_mcp.word.wordcom.runtime._ops import run_batch


def _table_com(fake_word):
    return fake_word.ActiveDocument.Tables(1)


# --- the method ----------------------------------------------------------------


def test_autofit_content(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("t"):
            doc.tables[1].autofit("content")
    t = _table_com(fake_word)
    assert t.AllowAutoFit is True
    assert t.autofit_behavior == int(WdAutoFitBehavior.CONTENT)


def test_autofit_window(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("t"):
            doc.tables[1].autofit("window")
    t = _table_com(fake_word)
    assert t.AllowAutoFit is True
    assert t.autofit_behavior == int(WdAutoFitBehavior.WINDOW)


def test_autofit_fixed_disables_autofit(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("t"):
            doc.tables[1].autofit("fixed")
    t = _table_com(fake_word)
    assert t.AllowAutoFit is False
    assert t.autofit_behavior == int(WdAutoFitBehavior.FIXED)


def test_autofit_default_is_content(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("t"):
            doc.tables[1].autofit()
    assert _table_com(fake_word).autofit_behavior == int(WdAutoFitBehavior.CONTENT)


def test_autofit_bad_mode(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError, match="autofit mode"):
            doc.tables[1].autofit("stretch")


# --- exec op -------------------------------------------------------------------


def test_op_autofit_table(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "autofit_table", "table": 1, "mode": "window"}], label="t"
        )
        assert exc is None and result["ok"]
    assert _table_com(fake_word).autofit_behavior == int(WdAutoFitBehavior.WINDOW)


def test_op_autofit_defaults_content(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "autofit_table", "table": 1}], label="t")
        assert exc is None and result["ok"]
    assert _table_com(fake_word).autofit_behavior == int(WdAutoFitBehavior.CONTENT)


# --- CLI -----------------------------------------------------------------------


# --- MCP -----------------------------------------------------------------------


def test_mcp_table_autofit(fake_word):
    pytest.importorskip("mcp")
    from office_mcp.word.wordcom.runtime._dispatch import _write_impl
    from office_mcp.word.wordcom.runtime._worker import InlineWorker

    r = _write_impl(InlineWorker(), "table", {"action": "autofit", "table": 1, "mode": "fixed"})
    assert r["ok"]
    assert _table_com(fake_word).autofit_behavior == int(WdAutoFitBehavior.FIXED)
