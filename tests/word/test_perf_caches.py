"""Caches that speed up formatting batches on long documents.

- `_para_index.batch_cache`: paragraph lookups inside one exec batch step from a
  paragraph already resolved instead of walking from the top of the document.
- `_com.install_dispatch_type_cache`: pywin32's late-bound wrappers share one
  type description per COM interface instead of rebuilding it per object.
"""

from __future__ import annotations

import types
from unittest.mock import MagicMock

from office_mcp.word import wordcom
from office_mcp.word.wordcom.model import _para_index
from office_mcp.word.wordcom.runtime import _com
from office_mcp.word.wordcom.runtime._ops import run_batch

# --- batch paragraph cache -------------------------------------------------------


def _reset_steps(doc):
    for para in doc.Paragraphs:
        para.Next.reset_mock()
        para.Previous.reset_mock()


def test_batch_cache_steps_from_a_cached_paragraph(fake_word):
    doc = fake_word.ActiveDocument
    total = _para_index.count(doc)
    with _para_index.batch_cache(doc):
        last = _para_index.paragraph(doc, total)
        before_last = _para_index.paragraph(doc, total - 1)
    assert last is doc.Paragraphs(total)
    assert before_last is doc.Paragraphs(total - 1)
    last.Previous.assert_called_once_with(1)


def test_batch_cache_returns_a_cached_paragraph_without_stepping(fake_word):
    doc = fake_word.ActiveDocument
    with _para_index.batch_cache(doc):
        first = _para_index.paragraph(doc, 2)
        _reset_steps(doc)
        again = _para_index.paragraph(doc, 2)
    assert again is first
    assert not any(p.Next.called or p.Previous.called for p in doc.Paragraphs)


def test_forget_cached_drops_the_lookups(fake_word):
    doc = fake_word.ActiveDocument
    total = _para_index.count(doc)
    with _para_index.batch_cache(doc):
        _para_index.paragraph(doc, total)
        _para_index.forget_cached()
        _reset_steps(doc)
        para = _para_index.paragraph(doc, total - 1)
    assert para is doc.Paragraphs(total - 1)
    assert not doc.Paragraphs(total).Previous.called


def test_no_caching_outside_a_batch(fake_word):
    doc = fake_word.ActiveDocument
    total = _para_index.count(doc)
    _para_index.paragraph(doc, total)
    _reset_steps(doc)
    assert _para_index.paragraph(doc, total - 1) is doc.Paragraphs(total - 1)
    assert not doc.Paragraphs(total).Previous.called


def test_batch_cache_ignores_other_documents(fake_word):
    doc = fake_word.ActiveDocument
    other = MagicMock()
    with _para_index.batch_cache(other):
        _para_index.paragraph(doc, 2)
        _reset_steps(doc)
        assert _para_index.paragraph(doc, 1) is doc.Paragraphs(1)
    assert not doc.Paragraphs(2).Previous.called


def test_stepping_past_the_end_falls_back_to_a_direct_lookup(fake_word):
    doc = fake_word.ActiveDocument
    total = _para_index.count(doc)
    with _para_index.batch_cache(doc):
        _para_index.paragraph(doc, total)
        assert _para_index.paragraph(doc, total + 1) is None


def test_run_batch_forgets_lookups_after_ops_that_may_add_paragraphs(fake_word, monkeypatch):
    forgets = []
    monkeypatch.setattr(_para_index, "forget_cached", lambda: forgets.append(1))
    ops = [
        {"op": "format_paragraph", "anchor_id": "para:2", "alignment": "center"},
        {"op": "apply_style", "anchor_id": "para:2", "name": "Normal"},
        {"op": "insert_break", "anchor_id": "bookmark:Address", "kind": "page"},
    ]
    with wordcom.attach() as word:
        result, exc = run_batch(word.documents.active, ops, label="t")
    assert exc is None, result
    assert len(forgets) == 1  # only after insert_break


# --- shared pywin32 type descriptions --------------------------------------------


def _fake_dynamic():
    class LazyDispatchItem:
        pass

    class CDispatch:
        def __init__(self, IDispatch, olerepr, userName=None, lazydata=None):
            self.__dict__["_olerepr_"] = olerepr
            self.__dict__["_builtMethods_"] = {}

    made = []

    def make_ole_repr(IDispatch, typeinfo, typecomp):
        item = LazyDispatchItem()
        made.append(item)
        return item

    module = types.SimpleNamespace(
        build=types.SimpleNamespace(LazyDispatchItem=LazyDispatchItem),
        CDispatch=CDispatch,
        MakeOleRepr=make_ole_repr,
    )
    return module, made


def _typeinfo(iid):
    info = MagicMock()
    info.GetTypeAttr.return_value = (iid, 0, 0)
    return info


def test_type_cache_shares_one_description_per_interface():
    dynamic, made = _fake_dynamic()
    assert _com.install_dispatch_type_cache(dynamic) is True
    a1 = dynamic.MakeOleRepr(None, _typeinfo("IID-A"), object())
    a2 = dynamic.MakeOleRepr(None, _typeinfo("IID-A"), object())
    b = dynamic.MakeOleRepr(None, _typeinfo("IID-B"), object())
    assert a1 is a2
    assert b is not a1
    assert len(made) == 2


def test_type_cache_shares_compiled_methods_between_objects():
    dynamic, _ = _fake_dynamic()
    _com.install_dispatch_type_cache(dynamic)
    olerepr = dynamic.MakeOleRepr(None, _typeinfo("IID-A"), object())
    one = dynamic.CDispatch(object(), olerepr)
    two = dynamic.CDispatch(object(), olerepr)
    one._builtMethods_["Item"] = "fn"
    assert two._builtMethods_ == {"Item": "fn"}


def test_type_cache_leaves_objects_without_type_info_alone():
    dynamic, made = _fake_dynamic()
    _com.install_dispatch_type_cache(dynamic)
    first = dynamic.MakeOleRepr(None, None, None)
    second = dynamic.MakeOleRepr(None, _typeinfo("IID-A"), None)  # no typecomp
    third = dynamic.MakeOleRepr(None, _typeinfo("IID-A"), None)
    assert len({id(first), id(second), id(third)}) == 3
    plain = dynamic.CDispatch(object(), first)
    assert plain._builtMethods_ == {}


def test_type_cache_installs_once():
    dynamic, made = _fake_dynamic()
    _com.install_dispatch_type_cache(dynamic)
    patched = dynamic.MakeOleRepr
    assert _com.install_dispatch_type_cache(dynamic) is True
    assert dynamic.MakeOleRepr is patched


def test_type_cache_skips_an_unexpected_pywin32():
    assert _com.install_dispatch_type_cache(types.SimpleNamespace()) is False
