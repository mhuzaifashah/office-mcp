"""Document variables — `doc.variables`, the `variables` CLI group, MCP.

Round-trips against `fake_word`, seeded with one variable (ClientName=Acme).
"""

from __future__ import annotations

import pytest

from office_mcp.word import wordcom
from office_mcp.word.wordcom.exceptions import AnchorNotFoundError
from office_mcp.word.wordcom.runtime._ops import run_batch

# --- the reader ----------------------------------------------------------------


def test_list_and_get(fake_word):
    with wordcom.attach() as word:
        variables = word.documents.active.variables
        assert len(variables) == 1
        assert variables.list() == {"ClientName": "Acme"}
        assert variables.get("ClientName") == "Acme"
        assert "ClientName" in variables
        with pytest.raises(AnchorNotFoundError):
            variables.get("Missing")


# --- writes --------------------------------------------------------------------


def test_set_new_and_update(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("t"):
            doc.variables.set("Region", "EU")  # new
            doc.variables.set("ClientName", "Globex")  # update
        assert doc.variables.list() == {"ClientName": "Globex", "Region": "EU"}


def test_set_coerces_to_string(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("t"):
            doc.variables.set("Count", 7)
        assert doc.variables.get("Count") == "7"


def test_delete_and_missing(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("t"):
            doc.variables.delete("ClientName")
        assert doc.variables.list() == {}
        with pytest.raises(AnchorNotFoundError):
            doc.variables.delete("ClientName")


# --- exec ops ------------------------------------------------------------------


def test_op_set_and_delete_variable(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {"op": "set_variable", "name": "Region", "value": "EU"},
                {"op": "delete_variable", "name": "ClientName"},
            ],
            label="t",
        )
        assert exc is None and result["ok"]
        assert doc.variables.list() == {"Region": "EU"}


# --- CLI -----------------------------------------------------------------------


# --- MCP -----------------------------------------------------------------------


def test_mcp_variables(fake_word):
    pytest.importorskip("mcp")
    from office_mcp.word.wordcom.runtime._dispatch import _read_impl, _write_impl
    from office_mcp.word.wordcom.runtime._worker import InlineWorker

    w = InlineWorker()
    assert _read_impl(w, "variables", {}) == {"ClientName": "Acme"}
    assert _write_impl(w, "set_variable", {"name": "R", "value": "EU"})["ok"]
    assert _write_impl(w, "delete_variable", {"name": "ClientName"})["ok"]
