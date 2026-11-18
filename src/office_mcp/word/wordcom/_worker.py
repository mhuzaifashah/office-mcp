"""A single dedicated COM worker thread.

COM objects are apartment-threaded: the thread that calls `CoInitialize` is the
only one allowed to talk to them. Callers (e.g. a server's thread pool) dispatch
work across many threads, so there is no stable COM-initialised thread to reach
Word from. `ComWorker` provides one: a long-lived daemon thread that
`CoInitialize`s once and then runs submitted callables one at a time. Every Word
operation funnels through `run_on_word(...)`, which blocks the caller until the
worker returns the result — so concurrent calls are also serialised, matching
Word's single-threaded reality.

`InlineWorker` is the test seam: it runs callables on the calling thread (the
fake COM in the test suite has no thread affinity), so unit tests skip the
thread machinery entirely.
"""

from __future__ import annotations

import os
import queue
import threading
from collections.abc import Callable
from concurrent.futures import Future
from concurrent.futures import TimeoutError as FutureTimeout
from typing import Any, Protocol, TypeVar

from .exceptions import WordBusyError

T = TypeVar("T")

# Sentinel pushed onto the queue to stop the worker loop.
_SHUTDOWN = object()

# How long a caller waits for one Word call before giving up (seconds).
# Override with WORDCOM_CALL_TIMEOUT; 0 or a negative value waits forever.
DEFAULT_CALL_TIMEOUT = 300.0


def call_timeout_from_env() -> float | None:
    """The per-call timeout from ``WORDCOM_CALL_TIMEOUT`` (None = no limit)."""
    raw = os.environ.get("WORDCOM_CALL_TIMEOUT", "").strip()
    if not raw:
        return DEFAULT_CALL_TIMEOUT
    try:
        value = float(raw)
    except ValueError:
        return DEFAULT_CALL_TIMEOUT
    return value if value > 0 else None


class Worker(Protocol):
    """Anything that can run a Word-touching callable and return its result."""

    def run_on_word(self, fn: Callable[[], T]) -> T: ...


class InlineWorker:
    """Runs callables synchronously on the caller's thread (tests only)."""

    def run_on_word(self, fn: Callable[[], T]) -> T:
        return fn()


class ComWorker:
    """Serialises all Word/COM work onto one CoInitialize'd daemon thread."""

    def __init__(self, *, timeout: float | None = DEFAULT_CALL_TIMEOUT) -> None:
        # `timeout` bounds how long `run_on_word` waits; None waits forever.
        self._timeout = timeout
        self._jobs: queue.Queue[Any] = queue.Queue()
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()

    def _ensure_started(self) -> None:
        with self._lock:
            if self._thread is not None and self._thread.is_alive():
                return
            self._thread = threading.Thread(target=self._loop, name="wordcom-com", daemon=True)
            self._thread.start()

    def _loop(self) -> None:
        # One outer CoInitialize keeps the STA apartment alive for the thread's
        # whole life; each job's own `attach()` nests safely (refcounted).
        # Guarded so the module stays importable where pythoncom is absent.
        try:
            import pythoncom

            pythoncom.CoInitialize()
            initialised = True
        except Exception:
            initialised = False
        try:
            while True:
                job = self._jobs.get()
                if job is _SHUTDOWN:
                    return
                fn, future = job
                if future.set_running_or_notify_cancel():
                    try:
                        future.set_result(fn())
                    except BaseException as exc:  # noqa: BLE001 — marshalled to caller
                        future.set_exception(exc)
        finally:
            if initialised:
                try:
                    import pythoncom

                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def run_on_word(self, fn: Callable[[], T]) -> T:
        """Run `fn` on the worker thread and block until it returns (or raises).

        Raises `WordBusyError` (retryable) if Word doesn't answer within the
        worker's timeout, e.g. because a dialog is blocking it. The call keeps
        running on the worker; later calls queue behind it until it finishes.
        """
        self._ensure_started()
        future: Future[T] = Future()
        self._jobs.put((fn, future))
        try:
            return future.result(timeout=self._timeout)
        except FutureTimeout:
            # Not started yet: drop it so it never runs after we've given up.
            future.cancel()
            raise WordBusyError(
                f"Word did not respond within {self._timeout:g}s "
                "(a dialog may be open, or it is still busy); retry shortly"
            ) from None

    def shutdown(self) -> None:
        """Stop the worker thread (best-effort; the thread is a daemon anyway)."""
        with self._lock:
            thread = self._thread
            self._thread = None
        if thread is not None and thread.is_alive():
            self._jobs.put(_SHUTDOWN)
            thread.join(timeout=2.0)
