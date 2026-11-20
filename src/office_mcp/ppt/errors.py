"""Exceptions for the PowerPoint controller and tools, each with a stable code.

`PowerPointError` lives in its own module so the controller's mixins can
import it without a cycle back to controller.py.
"""

from __future__ import annotations

from typing import Any


class PowerPointError(Exception):
    """Raised when a PowerPoint operation fails."""


class PowerPointNotRunningError(PowerPointError):
    """PowerPoint isn't running."""


class PresentationNotFoundError(PowerPointError):
    """No open presentation matches the name, or none is open."""


class PptOpError(PowerPointError):
    """A malformed command or op, or a batch that failed part-way.

    `details` carries `failure` (index, op, error), `ops_run` and `applied`.
    """

    def __init__(self, message: str, **details: Any) -> None:
        super().__init__(message)
        self.details = details


class PowerPointBusyError(PowerPointError):
    """PowerPoint didn't answer in time (a dialog is open, or it is busy)."""


_CODES: tuple[tuple[type[PowerPointError], str, bool], ...] = (
    (PowerPointNotRunningError, "powerpoint_not_running", False),
    (PresentationNotFoundError, "presentation_not_found", False),
    (PowerPointBusyError, "powerpoint_busy", True),
    (PptOpError, "op_error", False),
)


def classify(exc: PowerPointError) -> tuple[str, bool]:
    """`(code, retryable)` for a PowerPoint error."""
    for cls, code, retryable in _CODES:
        if isinstance(exc, cls):
            return code, retryable
    return "powerpoint_error", False
