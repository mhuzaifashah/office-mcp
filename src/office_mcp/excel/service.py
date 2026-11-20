"""The Excel backend the MCP tools call: one COM thread, one save policy.

Every method resolves its workbook on the Excel COM thread (by name, or the
active one), so calls follow the user between workbooks. The service never
quits Excel or closes a workbook; the user owns both.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ..comthread import ComThread, Runner, call_timeout_from_env
from ..paths import PathNotAllowed, SavePolicy, reject_nonlocal
from . import snapshot as _snapshot
from .dispatch import com_message, dispatch_read, run_batch
from .errors import ExcelBusyError, ExcelError, ExcelOpError
from .session import _xw, find_book, open_books, running_apps

__all__ = ["OPEN_EXTENSIONS", "SAVE_EXTENSIONS", "ExcelService"]

OPEN_EXTENSIONS = (".xlsx", ".xlsm", ".xlsb", ".xls", ".csv", ".xltx", ".xltm")
SAVE_EXTENSIONS = {".xlsx": 51, ".xlsm": 52, ".xlsb": 50}  # XlFileFormat


def _translate(exc: Exception) -> Exception:
    """COM errors become `ExcelError` with Excel's own message."""
    if type(exc).__name__ == "com_error":
        return ExcelError(com_message(exc))
    return exc


class ExcelService:
    """Read, edit, render and save the workbooks open in Excel.

    Args:
        save_dirs: Folders `save`, `save_as` and `export_pdf` may write into.
            Empty means writing is off.
        runner: Where COM calls run. Defaults to a dedicated COM thread; tests
            pass an `InlineRunner`.
    """

    def __init__(self, *, save_dirs: tuple[str, ...] = (), runner: Runner | None = None) -> None:
        self._policy = SavePolicy(save_dirs)
        self._runner = runner or ComThread(
            "excel",
            busy_error=ExcelBusyError,
            closed_error=ExcelError,
            timeout=call_timeout_from_env("OFFICE_MCP_CALL_TIMEOUT"),
        )

    def _run(self, fn: Any) -> Any:
        def job() -> Any:
            try:
                return fn()
            except Exception as exc:  # noqa: BLE001 - translated, then re-raised
                translated = _translate(exc)
                if translated is exc:
                    raise
                raise translated from exc

        return self._runner.run(job)

    def close(self) -> None:
        shutdown = getattr(self._runner, "shutdown", None)
        if shutdown is not None:
            shutdown()

    # -- tools -------------------------------------------------------------------

    def read(self, command: str, *, workbook: str | None = None, **params: Any) -> Any:
        return self._run(lambda: dispatch_read(find_book(workbook), command, params))

    def exec(self, ops: list[dict[str, Any]], *, workbook: str | None = None) -> dict[str, Any]:
        return self._run(lambda: run_batch(find_book(workbook), ops))

    def snapshot(
        self,
        *,
        workbook: str | None = None,
        sheet: str | None = None,
        address: str | None = None,
        dpi: int = 150,
        max_dim: int | None = 1200,
        headings: bool = True,
    ) -> list[tuple[int, bytes]]:
        return self._run(
            lambda: _snapshot.render(
                find_book(workbook),
                sheet=sheet,
                address=address,
                dpi=dpi,
                max_dim=max_dim,
                headings=headings,
            )
        )

    def write(
        self,
        command: str,
        *,
        workbook: str | None = None,
        path: str | None = None,
        overwrite: bool = False,
        sheet: str | None = None,
    ) -> dict[str, Any]:
        if command == "open":
            return self._open(path)
        if command == "new":
            return self._run(self._new)
        if command == "save":
            return self._run(lambda: self._save(find_book(workbook)))
        if command == "save_as":
            target = self._target(path, command, tuple(SAVE_EXTENSIONS), overwrite)
            return self._run(lambda: self._save_as(find_book(workbook), target))
        if command == "export_pdf":
            target = self._target(path, command, (".pdf",), overwrite)
            return self._run(lambda: self._export_pdf(find_book(workbook), target, sheet))
        raise ExcelOpError(f"unknown write command {command!r}")

    # -- write commands ----------------------------------------------------------

    def _target(
        self, path: str | None, command: str, extensions: tuple[str, ...], overwrite: bool
    ) -> Path:
        if not path:
            raise ExcelOpError(f"write command {command!r} needs path")
        target = self._policy.resolve(path)
        if target.suffix.lower() not in extensions:
            raise ExcelOpError(f"{command} writes {', '.join(extensions)} files, not {target.name}")
        if target.exists() and not overwrite:
            raise ExcelOpError(f"{target} already exists; pass overwrite=true to replace it")
        return target

    def _open(self, path: str | None) -> dict[str, Any]:
        if not path:
            raise ExcelOpError("write command 'open' needs path")
        reject_nonlocal(path, what="workbook")
        file = Path(path).expanduser().resolve()
        if file.suffix.lower() not in OPEN_EXTENSIONS:
            raise ExcelOpError(f"Excel opens {', '.join(OPEN_EXTENSIONS)} files, not {file.name}")
        if not file.is_file():
            raise FileNotFoundError(f"no such file: {file}")

        def job() -> dict[str, Any]:
            for book in open_books():
                if str(book.fullname).lower() == str(file).lower():
                    book.activate()
                    return _book_info(book, already_open=True)
            book = _app().books.open(str(file))
            return _book_info(book, already_open=False)

        return self._run(job)

    def _new(self) -> dict[str, Any]:
        book = _app().books.add()
        return _book_info(book)

    def _save(self, book: Any) -> dict[str, Any]:
        full = str(book.fullname)
        if not Path(full).is_absolute():
            raise ExcelOpError(
                f"{book.name} has never been saved; use save_as with a path in an allowed folder"
            )
        self._policy.resolve(full)
        book.api.Save()
        return {"ok": True, "command": "save", "path": full}

    def _save_as(self, book: Any, target: Path) -> dict[str, Any]:
        app_api = book.app.api
        alerts = app_api.DisplayAlerts
        app_api.DisplayAlerts = False  # the overwrite was already allowed
        try:
            book.api.SaveAs(Filename=str(target), FileFormat=SAVE_EXTENSIONS[target.suffix.lower()])
        finally:
            app_api.DisplayAlerts = alerts
        return {"ok": True, "command": "save_as", "path": str(target), "workbook": str(book.name)}

    def _export_pdf(self, book: Any, target: Path, sheet: str | None) -> dict[str, Any]:
        from .dispatch import _sheet_name  # noqa: PLC0415
        from .session import controller_for  # noqa: PLC0415

        src = book.sheets[_sheet_name(controller_for(book), sheet)].api if sheet else book.api
        src.ExportAsFixedFormat(Type=0, Filename=str(target), OpenAfterPublish=False)
        return {"ok": True, "command": "export_pdf", "path": str(target)}


def _app() -> Any:
    """The running Excel, or a new visible one when none is running."""
    apps = running_apps()
    if apps:
        return _xw().apps.active or apps[0]
    return _xw().App(visible=True, add_book=False)


def _book_info(book: Any, *, already_open: bool | None = None) -> dict[str, Any]:
    out: dict[str, Any] = {
        "ok": True,
        "workbook": str(book.name),
        "path": str(book.fullname),
        "sheets": [str(s.name) for s in book.sheets],
    }
    if already_open is not None:
        out["already_open"] = already_open
    return out


# Re-exported so the tools layer can classify path refusals without importing paths.
__all__ += ["PathNotAllowed"]
