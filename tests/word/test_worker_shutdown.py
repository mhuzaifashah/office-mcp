"""ComWorker shutdown never leaves two COM threads running (W17)."""

from __future__ import annotations

import threading

import pytest
from wordcom._worker import ComWorker
from wordcom.exceptions import WordcomError


def _com_threads():
    return [t for t in threading.enumerate() if t.name == "wordcom-com" and t.is_alive()]


def test_calls_after_shutdown_are_refused_while_old_thread_is_busy():
    before = len(_com_threads())
    worker = ComWorker(timeout=None)
    release, started = threading.Event(), threading.Event()

    def long_call():
        started.set()
        release.wait()

    caller = threading.Thread(target=lambda: worker.run_on_word(long_call), daemon=True)
    caller.start()
    try:
        assert started.wait(5)
        worker.shutdown(timeout=0.05)  # join times out: the thread is still busy
        with pytest.raises(WordcomError, match="shut down"):
            worker.run_on_word(lambda: "x")
        assert len(_com_threads()) == before + 1  # no second thread was started
    finally:
        release.set()
        caller.join(5)
    worker._thread.join(5)
    assert len(_com_threads()) == before  # the old thread exited on the sentinel


def test_jobs_queued_before_shutdown_still_run():
    worker = ComWorker(timeout=None)
    assert worker.run_on_word(lambda: 1) == 1
    worker.shutdown()
    assert not worker._thread.is_alive()


def test_shutdown_is_idempotent_and_safe_before_start():
    worker = ComWorker()
    worker.shutdown()
    worker.shutdown()
    with pytest.raises(WordcomError):
        worker.run_on_word(lambda: 1)
