"""Finding Word in the ROT, and not waiting on it forever (W16)."""

from __future__ import annotations

import sys
import threading
import types
from unittest.mock import MagicMock

import pytest

from office_mcp.word.wordcom.exceptions import WordBusyError
from office_mcp.word.wordcom.runtime import _com
from office_mcp.word.wordcom.runtime._worker import ComWorker, call_timeout_from_env

WORD_APP_MONIKER = "!{000209FF-0000-0000-C000-000000000046}"


class _Moniker:
    def __init__(self, name, obj):
        self.name, self.obj = name, obj

    def GetDisplayName(self, ctx, other):
        return self.name


def _app(name, docs):
    app = MagicMock(name=name)
    app.Name = name
    app.Documents.Count = docs
    return app


def _install_rot(monkeypatch, monikers):
    """Fake pythoncom + win32com.client over a ROT holding `monikers`."""
    bound = []

    class Enum:
        def __init__(self):
            self.items = list(monikers)

        def Next(self, n):
            return [self.items.pop(0)] if self.items else []

    class Rot:
        def EnumRunning(self):
            return Enum()

        def GetObject(self, moniker):
            bound.append(moniker.name)
            unknown = MagicMock()
            unknown.QueryInterface.return_value = moniker.obj
            return unknown

    pythoncom = types.ModuleType("pythoncom")
    pythoncom.GetRunningObjectTable = lambda reserved: Rot()
    pythoncom.CreateBindCtx = lambda reserved: object()
    pythoncom.IID_IDispatch = "IID_IDispatch"
    client = types.ModuleType("win32com.client")
    client.Dispatch = lambda obj: obj
    win32com = types.ModuleType("win32com")
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    return bound


def test_other_apps_entries_are_never_bound(monkeypatch):
    visio = _app("Microsoft Visio", 3)  # has Documents + ActiveDocument too
    word = _app("Microsoft Word", 1)
    bound = _install_rot(
        monkeypatch,
        [
            _Moniker("!{00021A20-0000-0000-C000-000000000046}", visio),
            _Moniker(r"C:\data\book.xlsx", MagicMock()),
            _Moniker(WORD_APP_MONIKER, word),
        ],
    )
    assert _com._find_word_with_docs() is word
    assert bound == [WORD_APP_MONIKER]


def test_non_word_app_behind_word_clsid_is_rejected(monkeypatch):
    _install_rot(monkeypatch, [_Moniker(WORD_APP_MONIKER, _app("Microsoft Visio", 2))])
    assert _com._find_word_with_docs() is None


def test_document_moniker_resolves_to_its_word_app(monkeypatch):
    word = _app("Microsoft Word", 1)
    doc = MagicMock()
    doc.Application = word
    _install_rot(
        monkeypatch,
        [_Moniker(WORD_APP_MONIKER, _app("Microsoft Word", 0)), _Moniker(r"C:\r.docm", doc)],
    )
    assert _com._find_word_with_docs() is word


def test_call_that_hangs_raises_word_busy():
    worker = ComWorker(timeout=0.2)
    release = threading.Event()
    try:
        with pytest.raises(WordBusyError) as info:
            worker.run_on_word(release.wait)
        assert info.value.retryable is True
    finally:
        release.set()
        worker.shutdown()


def test_queued_call_is_dropped_after_timeout():
    worker = ComWorker(timeout=0.2)
    release, started = threading.Event(), threading.Event()
    ran = []

    def block():
        started.set()
        release.wait()

    blocker = threading.Thread(target=lambda: _swallow(worker, block))
    blocker.start()
    assert started.wait(5)  # the worker is now busy with `block`
    try:
        with pytest.raises(WordBusyError):
            worker.run_on_word(lambda: ran.append(1))
    finally:
        release.set()
        blocker.join()
    assert worker.run_on_word(lambda: "ok") == "ok"
    worker.shutdown()
    assert ran == []  # gave up before it started, so it never ran later


def _swallow(worker, fn):
    try:
        worker.run_on_word(fn)
    except WordBusyError:
        pass


@pytest.mark.parametrize(
    ("raw", "expected"),
    [("", 300.0), ("45", 45.0), ("0", None), ("-1", None), ("junk", 300.0)],
)
def test_timeout_from_env(monkeypatch, raw, expected):
    monkeypatch.setenv("WORDCOM_CALL_TIMEOUT", raw)
    assert call_timeout_from_env() == expected
