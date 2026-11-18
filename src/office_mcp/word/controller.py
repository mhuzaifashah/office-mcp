"""WordController — thin facade over the `wordcom` engine package.

This is the public entry point of the `word` package. It owns a single
`ComWorker` (one CoInitialize'd STA thread) so every Word/COM call is serialised
onto that thread — required because hosts (e.g. an MCP server) call in from a thread pool
and COM objects are apartment-threaded.

It deliberately does **not** reimplement the operation vocabulary. The four
dispatch entry points reuse the vendored op-dispatch impls verbatim
(`wordcom/_dispatch.py`, relocated out of upstream's `mcp/server.py` with the
FastMCP shell stripped):

    read(command, **params)   -> wordcom.runtime._dispatch._read_impl     (introspection)
    write(command, **params)  -> wordcom.runtime._dispatch._write_impl    (one op, atomic undo)
    exec(ops, ...)            -> wordcom.runtime._dispatch._exec_impl      (batch, one undo)
    snapshot(...)             -> wordcom.runtime._dispatch._snapshot_impl  (page/anchor -> PNG)

The op param shapes are exactly those documented by upstream `wordcom` (and its
`_skill/` SKILL.md guides). Server/tool layers are built on top of this class.
"""

from __future__ import annotations

import os
from typing import Any

from .wordcom.exceptions import WordNotRunningError
from .wordcom.runtime import _com
from .wordcom.runtime._dispatch import (
    _exec_impl,
    _read_impl,
    _snapshot_impl,
    _write_impl,
)
from .wordcom.runtime._path_policy import PathPolicy, reject_nonlocal_path
from .wordcom.runtime._worker import ComWorker, call_timeout_from_env

__all__ = ["WordController"]


class WordController:
    """Drive a live Microsoft Word instance via a single dedicated COM thread.

    Args:
        save_dirs: Directories writes (`save`/`save_as`/`export_pdf`) are allowed
            to target. **Empty = saving disabled** (default-deny). When ``None``,
            the policy is read from the ``WORDCOM_SAVE_DIRS`` env var.
        image_dirs: Optional allowlist further restricting image-source paths
            (non-local UNC/URL forms are always screened regardless). When
            ``None``, read from ``WORDCOM_IMAGE_DIRS``.

    Each Word call waits at most ``WORDCOM_CALL_TIMEOUT`` seconds (default 300;
    ``0`` = no limit) and then fails with a retryable `WordBusyError`.

    The controller never opens or closes Word on its own beyond `ensure_running`;
    the user owns Word's lifecycle.
    """

    def __init__(
        self,
        *,
        save_dirs: list[str] | None = None,
        image_dirs: list[str] | None = None,
    ) -> None:
        self._worker = ComWorker(timeout=call_timeout_from_env())
        if save_dirs is None and image_dirs is None:
            self._policy = PathPolicy.from_env()
        else:
            self._policy = PathPolicy(save_dirs=save_dirs or (), image_dirs=image_dirs or ())

    # -- lifecycle ---------------------------------------------------------

    def ensure_running(self, *, visible: bool = True) -> bool:
        """Attach to a running Word, launching one on the STA thread if none is up."""

        def job() -> bool:
            try:
                _com.get_active_word()
            except WordNotRunningError:
                _com.launch_word(visible=visible)
            return True

        return self._worker.run_on_word(job)

    def open(self, path: str, *, visible: bool = True) -> dict[str, Any]:
        """Open a document in Word, launching Word if none is running.

        If the file is already open in the running instance, activates that
        document instead of opening a duplicate. Mirrors ``ExcelController.open``
        (the vendored ``wordcom`` engine has no open op of its own — it assumes
        the user already has the document open — so this lives on the facade).

        Args:
            path: Path to a ``.docx``/``.doc`` (or other Word-openable) file.
            visible: Show the Word window (default True).

        Returns ``{"success", "name", "path", "already_open"}``.
        """
        # Before any filesystem probe: exists() on a UNC path authenticates to
        # the remote host.
        reject_nonlocal_path(path, what="document")
        resolved = os.path.abspath(os.path.expanduser(path))
        if not os.path.exists(resolved):
            raise FileNotFoundError(f"File not found: {resolved}")
        if not os.path.isfile(resolved):
            raise IsADirectoryError(f"Not a file: {resolved}")

        from .wordcom import connect  # noqa: PLC0415

        def job() -> dict[str, Any]:
            with connect(launch_if_missing=True, visible=visible) as word:
                app = word.com
                target = os.path.normcase(resolved)
                for doc in app.Documents:
                    try:
                        if os.path.normcase(str(doc.FullName)) == target:
                            doc.Activate()
                            return {
                                "success": True,
                                "name": str(doc.Name),
                                "path": str(doc.FullName),
                                "already_open": True,
                            }
                    except Exception:  # noqa: BLE001 — skip a doc we can't introspect
                        continue
                doc = app.Documents.Open(resolved)
                try:
                    app.Visible = bool(visible)
                except Exception:  # noqa: BLE001 — some COM stubs defer Visible
                    pass
                return {
                    "success": True,
                    "name": str(doc.Name),
                    "path": str(doc.FullName),
                    "already_open": False,
                }

        return self._worker.run_on_word(job)

    def close(self) -> None:
        """Stop the COM worker thread (does not close Word). The controller is unusable after."""
        self._worker.shutdown()

    def __enter__(self) -> WordController:
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    # -- dispatch surface (mirrors the 4 upstream MCP tools) ---------------

    def read(self, command: str, /, **params: Any) -> Any:
        """Read-only introspection (`outline`, `status`, `find`, `read_text`, …)."""
        return _read_impl(self._worker, command, params)

    def write(self, command: str, /, **params: Any) -> dict[str, Any]:
        """Apply a single op as one atomic undo (`insert`, `replace`, `add_table`, …).

        `save`/`save_as`/`export_pdf` are gated by the controller's `PathPolicy`.
        """
        return _write_impl(self._worker, command, params, policy=self._policy)

    def exec(
        self,
        ops: list[dict[str, Any]],
        *,
        doc: str | None = None,
        label: str | None = None,
        tracked: bool = False,
    ) -> dict[str, Any]:
        """Apply a batch of ops as a single undo record.

        On the first failing op the whole batch is rolled back and the error is
        raised with the batch result attached: `exc.failure` (index, op, error,
        type), `exc.ops_run` and `exc.rolled_back`.
        """
        result, exc = _exec_impl(
            self._worker, ops, doc=doc, label=label, tracked=tracked, policy=self._policy
        )
        if exc is not None:
            exc.failure = result.get("failure")  # type: ignore[attr-defined]
            exc.ops_run = result.get("ops_run", 0)  # type: ignore[attr-defined]
            exc.rolled_back = result.get("rolled_back", False)  # type: ignore[attr-defined]
            raise exc
        return result

    def snapshot(
        self,
        *,
        doc: str | None = None,
        pages: str | None = None,
        anchor: str | None = None,
        dpi: int = 150,
        markup: str = "none",
        max_dim: int | None = None,
    ) -> list[tuple[int, bytes]]:
        """Render pages (or an anchor's region) to PNG bytes for a vision model.

        Returns a list of ``(page_number, png_bytes)``. Needs the `snapshot` extra
        (PyMuPDF). `pages` is ``"4"`` or ``"2-5"``; `anchor` is an anchor id like
        ``"heading:3"``.
        """
        return _snapshot_impl(
            self._worker,
            doc=doc,
            pages=pages,
            anchor=anchor,
            dpi=dpi,
            markup=markup,
            max_dim=max_dim,
        )

    # -- convenience -------------------------------------------------------

    def status(self) -> Any:
        """Per-open-document name/path/saved/active — no document needed."""
        return self.read("status")

    def outline(self, *, doc: str | None = None) -> Any:
        """Heading structure (`heading:N` anchor ids)."""
        return self.read("outline", doc=doc) if doc else self.read("outline")
