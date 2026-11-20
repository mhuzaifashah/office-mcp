"""Errors the Excel tools report, each with a stable code the agent can act on."""

from __future__ import annotations

from typing import Any

from .controller import ExcelError

__all__ = [
    "ExcelBusyError",
    "ExcelError",
    "ExcelNotRunningError",
    "ExcelOpError",
    "NotFoundError",
    "WorkbookNotFoundError",
    "classify",
]


class ExcelNotRunningError(ExcelError):
    """No Excel instance is running."""


class WorkbookNotFoundError(ExcelError):
    """No open workbook matches the name, or none is open."""


class NotFoundError(ExcelError):
    """A sheet, table, chart, pivot table or name doesn't exist.

    `kind` says which, so the hint can name the read that lists them.
    """

    def __init__(self, kind: str, name: str) -> None:
        super().__init__(f"{kind} {name!r} not found")
        self.kind = kind
        self.name = name


class ExcelOpError(ExcelError):
    """A malformed command or op, or a batch that failed part-way.

    A failed batch sets `failure` (index, op, error), `ops_run`, and what was
    restored (`rolled_back`, `not_rolled_back`).
    """

    def __init__(self, message: str, **details: Any) -> None:
        super().__init__(message)
        self.details = details


class ExcelBusyError(ExcelError):
    """Excel didn't answer in time (a dialog is open, or it is busy)."""


# Stable error codes, matched by the hints in `tools.py`.
_CODES: tuple[tuple[type[ExcelError], str, bool], ...] = (
    (ExcelNotRunningError, "excel_not_running", False),
    (WorkbookNotFoundError, "workbook_not_found", False),
    (NotFoundError, "not_found", False),
    (ExcelBusyError, "excel_busy", True),
    (ExcelOpError, "op_error", False),
)


def classify(exc: ExcelError) -> tuple[str, bool]:
    """`(code, retryable)` for an Excel error."""
    for cls, code, retryable in _CODES:
        if isinstance(exc, cls):
            return code, retryable
    return "excel_error", False
