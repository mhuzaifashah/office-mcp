"""A failed batch is undone as one step and the failure is reported (W12/W13)."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest
import wordcom
from wordcom._ops import run_batch
from wordcom.exceptions import OpError

GOOD = {"op": "insert_break", "anchor_id": "bookmark:Address", "kind": "page"}
BAD = {"op": "insert_break", "kind": "page"}  # no anchor_id


def _run(ops):
    with wordcom.attach() as word:
        return run_batch(word.documents.active, ops, label="t")


def test_failure_after_applied_ops_undoes_whole_batch(fake_word):
    calls = MagicMock()
    calls.attach_mock(fake_word.UndoRecord.EndCustomRecord, "end")
    calls.attach_mock(fake_word.ActiveDocument.Undo, "undo")
    fake_word.ActiveDocument.Undo.return_value = True

    result, exc = _run([GOOD, BAD])

    assert exc is not None
    assert result["ok"] is False
    assert result["ops_run"] == 1
    assert result["failure"]["index"] == 1
    assert result["rolled_back"] is True
    # Exactly one Undo, issued after the custom record closed, so it reverts
    # the whole batch rather than a single action inside it.
    fake_word.ActiveDocument.Undo.assert_called_once_with(1)
    names = [c[0] for c in calls.mock_calls]
    assert names.index("end") < names.index("undo")


def test_failure_on_first_op_does_not_undo_user_history(fake_word):
    result, exc = _run([BAD, GOOD])
    assert exc is not None
    assert result["ops_run"] == 0
    assert result["rolled_back"] is False
    fake_word.ActiveDocument.Undo.assert_not_called()


def test_success_does_not_undo(fake_word):
    result, exc = _run([GOOD])
    assert exc is None
    assert result["ok"] is True
    assert "rolled_back" not in result
    fake_word.ActiveDocument.Undo.assert_not_called()


def test_no_undo_record_means_no_partial_undo(fake_word):
    # Without UndoRecord the batch isn't one undo entry; a single Undo would
    # revert only the last action, so none is attempted.
    fake_word.UndoRecord.StartCustomRecord.side_effect = RuntimeError("no UndoRecord")
    result, exc = _run([GOOD, BAD])
    assert exc is not None
    assert result["rolled_back"] is False
    fake_word.ActiveDocument.Undo.assert_not_called()


def test_undo_refused_reports_not_rolled_back(fake_word):
    fake_word.ActiveDocument.Undo.return_value = False
    result, _ = _run([GOOD, BAD])
    assert result["rolled_back"] is False


def test_controller_exec_attaches_failure_to_error(monkeypatch):
    from office_mcp.word import WordController
    from office_mcp.word import controller as controller_mod

    failure = {"index": 2, "op": BAD, "error": "boom", "type": "OpError"}
    err = OpError("boom")
    monkeypatch.setattr(
        controller_mod,
        "_exec_impl",
        lambda *a, **k: ({"ok": False, "ops_run": 2, "failure": failure, "rolled_back": True}, err),
    )
    ctl = object.__new__(WordController)
    ctl._worker = None
    ctl._policy = None

    with pytest.raises(OpError) as info:
        ctl.exec([GOOD, GOOD, BAD])

    assert info.value.failure == failure
    assert info.value.ops_run == 2
    assert info.value.rolled_back is True


# W20: a ValueError/TypeError from bad input fails the batch cleanly.


@pytest.mark.parametrize(
    "bad",
    [
        {"op": "set_cell", "table": 1, "row": "abc", "col": 1, "text": "x"},  # TypeError
        {"op": "apply_list", "anchor_id": "bookmark:Address", "type": "zigzag"},  # ValueError
    ],
)
def test_value_error_becomes_structured_failure_and_rolls_back(fake_word, bad):
    fake_word.ActiveDocument.Undo.return_value = True
    result, exc = _run([GOOD, bad])
    assert isinstance(exc, OpError)
    assert result["ok"] is False
    assert result["failure"]["index"] == 1
    assert result["failure"]["type"] == "OpError"
    assert isinstance(exc.__cause__, (ValueError, TypeError))
    assert result["rolled_back"] is True
    fake_word.ActiveDocument.Undo.assert_called_once_with(1)
