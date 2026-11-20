"""One COM thread per Office app.

COM objects are apartment-threaded: only the thread that called `CoInitialize`
may use them. The MCP server calls tools from a thread pool, so each app gets a
long-lived thread that initialises COM once and runs submitted calls one at a
time. Calls are serialised, matching how an Office app handles automation.

Word keeps its own copy of this (`office_mcp.word.wordcom.runtime._worker`),
which raises Word's error types. Other apps use this one.
"""

from __future__ import annotations

import os
import queue
import threading
from collections.abc import Callable
from concurrent.futures import Future
from concurrent.futures import TimeoutError as FutureTimeout
from typing import Any, Protocol, TypeVar

T = TypeVar("T")

_SHUTDOWN = object()

DEFAULT_CALL_TIMEOUT = 300.0


def call_timeout_from_env(var: str) -> float | None:
    """Seconds one call may take, from env var `var` (default 300; 0 = no limit)."""
    raw = os.environ.get(var, "").strip()
    if not raw:
        return DEFAULT_CALL_TIMEOUT
    try:
        value = float(raw)
    except ValueError:
        return DEFAULT_CALL_TIMEOUT
    return value if value > 0 else None


class Runner(Protocol):
    """Anything that runs an app-touching callable and returns its result."""

    def run(self, fn: Callable[[], T]) -> T: ...


class InlineRunner:
    """Runs callables on the caller's thread (tests with fake COM objects)."""

    def run(self, fn: Callable[[], T]) -> T:
        return fn()

    def shutdown(self, timeout: float = 2.0) -> None:
        return None


class ComThread:
    """Serialises one app's COM calls onto a single CoInitialize'd daemon thread.

    Args:
        name: Thread name, also used in messages ("excel").
        busy_error: Builds the exception raised when a call outlives `timeout`
            (the app is showing a dialog or is still busy). It gets the message.
        closed_error: Builds the exception raised for calls after `shutdown`.
        timeout: Seconds a caller waits for one call; None waits forever.
    """

    def __init__(
        self,
        name: str,
        *,
        busy_error: Callable[[str], Exception],
        closed_error: Callable[[str], Exception],
        timeout: float | None = DEFAULT_CALL_TIMEOUT,
    ) -> None:
        self._name = name
        self._busy_error = busy_error
        self._closed_error = closed_error
        self._timeout = timeout
        self._jobs: queue.Queue[Any] = queue.Queue()
        self._thread: threading.Thread | None = None
        self._closed = False
        # Guards thread start, `_closed`, and every put onto `_jobs`, so no job
        # lands behind the shutdown sentinel and no second thread starts.
        self._lock = threading.Lock()

    def _submit(self, job: Any) -> None:
        with self._lock:
            if self._closed:
                raise self._closed_error(f"the {self._name} COM thread has been shut down")
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(
                    target=self._loop, name=f"{self._name}-com", daemon=True
                )
                self._thread.start()
            self._jobs.put(job)

    def _loop(self) -> None:
        try:
            import pythoncom  # type: ignore[import-not-found]

            pythoncom.CoInitialize()
            initialised = True
        except Exception:  # noqa: BLE001 - importable where pywin32 is absent
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
                    except BaseException as exc:  # noqa: BLE001 - marshalled to the caller
                        future.set_exception(exc)
        finally:
            if initialised:
                try:
                    import pythoncom  # type: ignore[import-not-found]

                    pythoncom.CoUninitialize()
                except Exception:  # noqa: BLE001
                    pass

    def run(self, fn: Callable[[], T]) -> T:
        """Run `fn` on the COM thread and wait for its result (or exception).

        Raises `busy_error` when the app doesn't answer within the timeout. The
        call keeps running on the thread; later calls queue behind it.
        """
        future: Future[T] = Future()
        self._submit((fn, future))
        try:
            return future.result(timeout=self._timeout)
        except FutureTimeout:
            future.cancel()  # not started yet: never run it after giving up
            raise self._busy_error(
                f"{self._name.capitalize()} did not respond within {self._timeout:g}s "
                "(a dialog may be open, or it is still busy); retry shortly"
            ) from None

    def shutdown(self, timeout: float = 2.0) -> None:
        """Stop the thread. Queued calls still run; new calls are refused. Idempotent."""
        with self._lock:
            if self._closed:
                return
            self._closed = True
            thread = self._thread
            if thread is not None and thread.is_alive():
                self._jobs.put(_SHUTDOWN)
        if thread is not None:
            thread.join(timeout=timeout)
