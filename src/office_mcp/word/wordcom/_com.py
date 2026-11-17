"""Thin COM helpers — the mockable seam between wordcom and pywin32.

Tests substitute fakes for `get_active_word` / `launch_word` via monkeypatch.
Everything else in wordcom only sees duck-typed dispatch objects.
"""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any

from .exceptions import WordNotRunningError, from_com_error

_WORD_PROG_ID = "Word.Application"


@contextmanager
def com_apartment() -> Iterator[None]:
    """STA apartment lifecycle. Nests safely via pythoncom's reference counting."""
    try:
        import pythoncom  # type: ignore[import-not-found]
    except ImportError:
        # Non-Windows or pywin32 missing: yield without initialising. The first
        # real COM call will fail with a clearer error.
        yield
        return

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def get_active_word() -> Any:
    """Return the Application COM object for an already-running Word, or raise.

    `GetActiveObject` is tried first. Under Office 365 / OneDrive (and some
    multi-instance setups) it can return an *empty* Application proxy whose
    `Documents.Count` is 0 even though a document is open in another instance —
    the open doc lives in the Running Object Table (ROT) as a file moniker. In
    that case (or if `GetActiveObject` fails outright) we scan the ROT for an
    instance that actually has documents. Ported from word-mcp-live's
    `_find_word_with_docs` (ykarapazar/word-mcp-live).

    LOCAL-PATCH: this ROT fallback is the only change to vendored `_com.py`.
    """
    try:
        from win32com.client import GetActiveObject  # type: ignore[import-not-found]
    except ImportError as e:
        raise WordNotRunningError(
            "pywin32 is not installed; wordcom requires Windows + pywin32"
        ) from e

    try:
        app = GetActiveObject(_WORD_PROG_ID)
    except Exception:  # pywintypes.com_error or similar — try the ROT next
        app = None

    if app is not None:
        try:
            if app.Documents.Count > 0:
                return app
        except Exception:  # noqa: BLE001 — couldn't introspect; fall through to ROT
            pass

    with_docs = _find_word_with_docs()
    if with_docs is not None:
        return with_docs

    if app is not None:
        # An (empty) instance exists; hand it back so the caller can open a file.
        return app
    raise WordNotRunningError("no running Microsoft Word instance found")


def _find_word_with_docs() -> Any | None:
    """Scan the Running Object Table for a Word.Application that has open docs.

    Returns the Application COM object, or None if none is found. Best-effort:
    any per-entry failure is swallowed so one bad moniker can't abort the scan.
    """
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        return None

    try:
        # COM handles are duck-typed; keep them `Any` like the rest of wordcom
        # so the type checker doesn't trip on pywin32's partial ROT stubs.
        rot: Any = pythoncom.GetRunningObjectTable(0)
        enum: Any = rot.EnumRunning()
    except Exception:  # noqa: BLE001
        return None

    file_monikers: list[Any] = []
    # Pass 1: a direct Word.Application ROT entry that already has documents.
    while True:
        try:
            batch = enum.Next(1)
        except Exception:  # noqa: BLE001
            break
        if not batch:
            break
        moniker = batch[0]
        try:
            ctx = pythoncom.CreateBindCtx(0)  # type: ignore[call-arg]  # stub omits reserved arg
            name = moniker.GetDisplayName(ctx, None)
            obj = rot.GetObject(moniker)
            dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
            com_obj = win32com.client.Dispatch(dispatch)
            if hasattr(com_obj, "Documents") and hasattr(com_obj, "ActiveDocument"):
                if com_obj.Documents.Count > 0:
                    return com_obj
            if name and name.lower().endswith((".docx", ".doc")):
                file_monikers.append(moniker)
        except Exception:  # noqa: BLE001 — still remember file monikers for pass 2
            try:
                ctx = pythoncom.CreateBindCtx(0)  # type: ignore[call-arg]  # stub omits reserved arg
                name = moniker.GetDisplayName(ctx, None)
                if name and name.lower().endswith((".docx", ".doc")):
                    file_monikers.append(moniker)
            except Exception:  # noqa: BLE001
                pass

    # Pass 2: a document file moniker → Document COM → its parent Application.
    for moniker in file_monikers:
        try:
            obj = rot.GetObject(moniker)
            dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
            doc = win32com.client.Dispatch(dispatch)
            app = doc.Application
            if app.Documents.Count > 0:
                return app
        except Exception:  # noqa: BLE001
            continue
    return None


def launch_word(visible: bool = True) -> Any:
    """Launch a new Word instance and return its Application COM object."""
    try:
        from win32com.client import Dispatch  # type: ignore[import-not-found]
    except ImportError as e:
        raise WordNotRunningError(
            "pywin32 is not installed; wordcom requires Windows + pywin32"
        ) from e

    app = Dispatch(_WORD_PROG_ID)
    try:
        app.Visible = bool(visible)
    except Exception:
        # Some COM stubs may not let us flip Visible immediately; not fatal.
        pass
    return app


@contextmanager
def preserve_saved(doc_com: Any) -> Iterator[None]:
    """Snapshot and restore a document's dirty bit around a content-neutral op.

    Word's `Repaginate()` (and similar layout-only calls) flip `Document.Saved`
    to `False` even though nothing the user cares about changed. A *read* that
    repaginates — `stats`, `location` — would then leave a spurious unsaved-
    changes star on a document the user had just saved. Wrap the repaginate in
    this and the flag is put back exactly as it was. Best-effort: if reading or
    restoring `Saved` raises, the read still proceeds (a stale star beats a
    crash on a pure read).
    """
    try:
        was_saved = bool(doc_com.Saved)
    except Exception:  # noqa: BLE001 — never let dirty-bit bookkeeping break a read
        was_saved = None
    try:
        yield
    finally:
        if was_saved:
            try:
                doc_com.Saved = True
            except Exception:  # noqa: BLE001
                pass


@contextmanager
def translate_com_errors() -> Iterator[None]:
    """Translate pywintypes.com_error into wordcom's typed exceptions."""
    try:
        import pywintypes  # type: ignore[import-not-found]

        com_error_type: type = pywintypes.com_error
    except ImportError:
        com_error_type = ()  # type: ignore[assignment]

    try:
        yield
    except Exception as exc:  # noqa: BLE001
        if com_error_type and isinstance(exc, com_error_type):  # type: ignore[arg-type]
            raise from_com_error(exc) from exc
        raise
