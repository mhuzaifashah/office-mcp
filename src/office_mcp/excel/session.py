"""Find the workbook a tool call is about, in the Excel the user has open.

Every call resolves its workbook afresh (by name, or the active one), so the
tools follow the user as they switch workbooks. Runs on the Excel COM thread.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .controller import ExcelController
from .errors import ExcelNotRunningError, WorkbookNotFoundError


def _xw() -> Any:
    import xlwings as xw  # noqa: PLC0415 - heavy, and only needed once Excel is used

    return xw


def running_apps() -> list[Any]:
    """Every running Excel instance xlwings can reach."""
    try:
        return list(_xw().apps)
    except Exception:  # noqa: BLE001 - no Excel, or COM refuses: treat as none running
        return []


def open_books() -> list[Any]:
    """Every open workbook across the running Excel instances."""
    books: list[Any] = []
    for app in running_apps():
        try:
            books.extend(app.books)
        except Exception:  # noqa: BLE001 - an instance busy with a dialog: skip it
            continue
    return books


def find_book(name: str | None = None) -> Any:
    """The workbook called `name` (file name or full path), or the active one."""
    apps = running_apps()
    if not apps:
        raise ExcelNotRunningError("Excel is not running")
    if name:
        wanted = name.strip().lower()
        for book in open_books():
            if str(book.name).lower() == wanted or str(book.fullname).lower() == wanted:
                return book
        # Allow the name without its extension ("Budget" for "Budget.xlsx").
        for book in open_books():
            if Path(str(book.name)).stem.lower() == wanted:
                return book
        raise WorkbookNotFoundError(f"no open workbook is called {name!r}")
    try:
        book = _xw().apps.active.books.active
    except Exception:  # noqa: BLE001
        book = None
    if book is None:
        books = open_books()
        if not books:
            raise WorkbookNotFoundError("Excel has no workbook open")
        book = books[0]
    return book


def controller_for(book: Any) -> ExcelController:
    """An `ExcelController` bound to `book`, without touching Excel's settings.

    `ExcelController(...)` turns off Excel's alerts for good; the tools must not
    change the user's Excel like that, so they bind the controller directly.
    """
    ctrl = ExcelController.__new__(ExcelController)
    ctrl._app = book.app
    ctrl._book = book
    return ctrl
