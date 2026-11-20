"""The Excel tools: `excel_read`, `excel_exec`, `excel_snapshot`, `excel_write`.

Four task-shaped tools, like Word's. The read commands and edit ops behind them
live in `dispatch.py` and are documented in the bundled guide
(`excel_read(command="guide")`, also the `office-mcp://excel/guide` resource).
"""

from __future__ import annotations

import logging
import threading
from collections.abc import Callable
from importlib.resources import files
from typing import Annotated, Any, Literal

from mcp_types import CallToolResult, ContentBlock
from pydantic import Field
from pydantic.json_schema import SkipJsonSchema

from ..paths import PathNotAllowed
from ..server import response
from ..server.config import ServerConfig
from ..server.tool_definition import ResourceDefinition, ToolDefinition
from .errors import ExcelError, ExcelOpError, NotFoundError, classify
from .service import ExcelService

__all__ = ["READ_COMMANDS", "WRITE_COMMANDS", "ExcelToolset", "guide_text"]

log = logging.getLogger(__name__)

GUIDE_URI = "office-mcp://excel/guide"

# Every `excel_read` command, matching `dispatch.READS` plus guide (a test keeps them in sync).
ReadCommand = Literal[
    "guide",
    "status",
    "sheets",
    "used_range",
    "read",
    "find",
    "summary",
    "tables",
    "table_data",
    "charts",
    "pivots",
    "pivot_fields",
    "names",
    "calc_mode",
]
READ_COMMANDS: tuple[str, ...] = ReadCommand.__args__  # type: ignore[attr-defined]

WriteCommand = Literal["open", "new", "save", "save_as", "export_pdf"]
WRITE_COMMANDS: tuple[str, ...] = WriteCommand.__args__  # type: ignore[attr-defined]

_MORE_HINT = (
    "Read a smaller range (sheet and range), or use summary for statistics instead of values."
)

_HINTS: dict[str, str] = {
    "excel_not_running": "Ask the user to start Excel and open their workbook, or open a file "
    'with excel_write(command="open", path=...).',
    "workbook_not_found": 'Call excel_read(command="status") for the names of open workbooks.',
    "excel_busy": "A dialog may be open in Excel, or a cell is being edited. Ask the user to "
    "close it or press Enter, then retry the same call.",
    "path_not_allowed": "Files are only written inside the folders the server was started "
    "with (--save-dir). Ask the user which allowed folder to use.",
    "file_not_found": "Check the path with the user; it must be a local file.",
}
_NOT_FOUND_HINTS: dict[str, str] = {
    "sheet": 'List the sheets with excel_read(command="sheets").',
    "table": 'List the tables with excel_read(command="tables").',
    "chart": 'List the charts with excel_read(command="charts").',
    "pivot table": 'List the pivot tables with excel_read(command="pivots").',
}
_OP_HINT = (
    'Fix the command or op fields. excel_read(command="guide") lists every op with its '
    "required fields."
)
_EXCEL_ERROR_HINT = (
    "Excel refused the change. Check the address, sheet and values, read the cells "
    "involved, then retry."
)

Workbook = Annotated[
    str | SkipJsonSchema[None],
    Field(
        description='Workbook file name from excel_read(command="status"); omit for the active one.'
    ),
]
Sheet = Annotated[
    str | SkipJsonSchema[None],
    Field(description="Sheet name; omit for the active sheet."),
]


def guide_text() -> str:
    return (files("office_mcp.excel") / "guide.md").read_text(encoding="utf-8")


def _error(exc: BaseException) -> CallToolResult:
    if isinstance(exc, ExcelError):
        code, retryable = classify(exc)
        hint = _HINTS.get(code)
        if isinstance(exc, NotFoundError):
            hint = _NOT_FOUND_HINTS.get(exc.kind)
        elif isinstance(exc, ExcelOpError) and hint is None:
            hint = _OP_HINT
        elif hint is None:
            hint = _EXCEL_ERROR_HINT
        details = dict(getattr(exc, "details", None) or {})
        if details.get("failure") and hint in (_OP_HINT, _EXCEL_ERROR_HINT):
            hint = (
                "Fix the failing op (failure.index) and resend the batch. Check rolled_back "
                "and not_rolled_back first: changes listed there are still in the workbook."
            )
        return response.error_result(
            code, str(exc), retryable=retryable, hint=hint, details=details or None
        )
    if isinstance(exc, PathNotAllowed):
        return response.error_result("path_not_allowed", str(exc), hint=_HINTS["path_not_allowed"])
    if isinstance(exc, FileNotFoundError):
        return response.error_result("file_not_found", str(exc), hint=_HINTS["file_not_found"])
    if isinstance(exc, (ValueError, TypeError, KeyError)):
        return response.error_result("error", str(exc), hint=_OP_HINT)
    log.exception("unexpected error in an Excel tool")
    return response.error_result("internal_error", f"{type(exc).__name__}: {exc}")


def _guarded(fn: Callable[[], CallToolResult]) -> CallToolResult:
    try:
        return fn()
    except Exception as exc:  # noqa: BLE001 - every failure becomes a readable tool error
        return _error(exc)


def _without_none(**values: Any) -> dict[str, Any]:
    return {k: v for k, v in values.items() if v is not None}


class ExcelToolset:
    """The Excel app's tools, guide resource and server instructions.

    The `ExcelService` (and its COM thread) starts on the first tool call.
    `close` stops the thread; it never closes Excel or a workbook.
    """

    app = "excel"
    instructions = (
        "Excel tools (excel_*) edit the workbooks open in Excel on this PC. Start with "
        'excel_read(command="status"), then read the cells you need. Before your first edit, '
        'call excel_read(command="guide") once: it lists every excel_exec op with its fields. '
        "Excel can't undo automated edits, so read data before you overwrite it."
    )

    def __init__(
        self,
        config: ServerConfig,
        *,
        service_factory: Callable[[], ExcelService] | None = None,
    ) -> None:
        self._config = config
        self._factory = service_factory or (lambda: ExcelService(save_dirs=tuple(config.save_dirs)))
        self._service: ExcelService | None = None
        self._lock = threading.Lock()

    @property
    def service(self) -> ExcelService:
        with self._lock:
            if self._service is None:
                self._service = self._factory()
            return self._service

    def close(self) -> None:
        with self._lock:
            if self._service is not None:
                self._service.close()
                self._service = None

    def tools(self) -> list[ToolDefinition]:
        return [
            ToolDefinition(
                name="excel_read",
                title="Read Excel workbook",
                description=_READ_DESCRIPTION,
                handler=self._read_handler(),
                read_only=True,
            ),
            ToolDefinition(
                name="excel_exec",
                title="Edit Excel workbook",
                description=_EXEC_DESCRIPTION,
                handler=self._exec_handler(),
                read_only=False,
                destructive=True,
            ),
            ToolDefinition(
                name="excel_snapshot",
                title="Render Excel range",
                description=_SNAPSHOT_DESCRIPTION,
                handler=self._snapshot_handler(),
                read_only=True,
            ),
            ToolDefinition(
                name="excel_write",
                title="Open, save or export an Excel workbook",
                description=_WRITE_DESCRIPTION,
                handler=self._write_handler(),
                read_only=False,
                destructive=True,
            ),
        ]

    def resources(self) -> list[ResourceDefinition]:
        return [
            ResourceDefinition(
                uri=GUIDE_URI,
                name="excel-guide",
                title="Excel tools guide",
                description="Addressing, every excel_read command and every excel_exec op.",
                mime_type="text/markdown",
                read=guide_text,
            )
        ]

    # -- handlers ----------------------------------------------------------------

    def _read_handler(self) -> Callable[..., CallToolResult]:
        max_chars = self._config.max_output_chars

        def excel_read(
            command: Annotated[ReadCommand, Field(description="What to read.")],
            workbook: Workbook = None,
            sheet: Sheet = None,
            range: Annotated[
                str | SkipJsonSchema[None],
                Field(description="A1 address or defined name (read, summary, find)."),
            ] = None,
            formulas: Annotated[
                bool | SkipJsonSchema[None],
                Field(description="Also return the formulas (read)."),
            ] = None,
            text: Annotated[
                str | SkipJsonSchema[None], Field(description="Text to look for (find).")
            ] = None,
            match_case: Annotated[
                bool | SkipJsonSchema[None], Field(description="Case-sensitive (find).")
            ] = None,
            whole_cell: Annotated[
                bool | SkipJsonSchema[None],
                Field(description="Match the whole cell, not part of it (find)."),
            ] = None,
            limit: Annotated[
                int | SkipJsonSchema[None],
                Field(ge=1, le=1000, description="Most matches, default 50 (find)."),
            ] = None,
            table: Annotated[
                str | SkipJsonSchema[None], Field(description="Table name (table_data).")
            ] = None,
            visible_only: Annotated[
                bool | SkipJsonSchema[None],
                Field(description="Only rows a filter leaves visible (table_data)."),
            ] = None,
            pivot: Annotated[
                str | SkipJsonSchema[None], Field(description="Pivot table name (pivot_fields).")
            ] = None,
        ) -> CallToolResult:
            def run() -> CallToolResult:
                if command == "guide":
                    return response.text_result(
                        guide_text(),
                        max_chars=max_chars,
                        more_hint=f"Read the full guide as the {GUIDE_URI} resource.",
                        label="guide",
                    )
                params = _without_none(
                    sheet=sheet,
                    range=range,
                    formulas=formulas,
                    text=text,
                    match_case=match_case,
                    whole_cell=whole_cell,
                    limit=limit,
                    table=table,
                    visible_only=visible_only,
                    pivot=pivot,
                )
                result = self.service.read(command, workbook=workbook, **params)
                return response.json_result(result, max_chars=max_chars, more_hint=_MORE_HINT)

            return _guarded(run)

        return excel_read

    def _exec_handler(self) -> Callable[..., CallToolResult]:
        max_chars = self._config.max_output_chars

        def excel_exec(
            ops: Annotated[
                list[dict[str, Any]],
                Field(
                    min_length=1,
                    description='Ops to apply in order. Each is {"op": "<name>", ...fields}; '
                    'see excel_read(command="guide").',
                ),
            ],
            workbook: Workbook = None,
        ) -> CallToolResult:
            def run() -> CallToolResult:
                result = self.service.exec(ops, workbook=workbook)
                return response.json_result(result, max_chars=max_chars)

            return _guarded(run)

        return excel_exec

    def _snapshot_handler(self) -> Callable[..., CallToolResult]:
        default_max_dim = self._config.snapshot_max_dim
        max_pages = self._config.snapshot_max_pages

        def excel_snapshot(
            sheet: Sheet = None,
            range: Annotated[
                str | SkipJsonSchema[None],
                Field(description="A1 address to render; omit for the sheet's used range."),
            ] = None,
            headings: Annotated[
                bool, Field(description="Show row numbers and column letters.")
            ] = True,
            dpi: Annotated[int, Field(ge=72, le=300, description="Render resolution.")] = 150,
            max_dim: Annotated[
                int | SkipJsonSchema[None],
                Field(
                    ge=100,
                    le=4000,
                    description="Cap on each page's long edge in pixels "
                    "(default: the server's --snapshot-max-dim).",
                ),
            ] = None,
            workbook: Workbook = None,
        ) -> CallToolResult:
            def run() -> CallToolResult:
                shots = self.service.snapshot(
                    workbook=workbook,
                    sheet=sheet,
                    address=range,
                    dpi=dpi,
                    max_dim=max_dim if max_dim is not None else default_max_dim,
                    headings=headings,
                )
                return _snapshot_result(shots, max_pages)

            return _guarded(run)

        return excel_snapshot

    def _write_handler(self) -> Callable[..., CallToolResult]:
        max_chars = self._config.max_output_chars

        def excel_write(
            command: Annotated[WriteCommand, Field(description="What to do.")],
            path: Annotated[
                str | SkipJsonSchema[None],
                Field(
                    description="Local file path: the workbook to open (open), or the target "
                    "file (save_as, export_pdf)."
                ),
            ] = None,
            overwrite: Annotated[
                bool | SkipJsonSchema[None],
                Field(description="Replace an existing file (save_as, export_pdf)."),
            ] = None,
            sheet: Annotated[
                str | SkipJsonSchema[None],
                Field(description="Export only this sheet (export_pdf)."),
            ] = None,
            workbook: Workbook = None,
        ) -> CallToolResult:
            def run() -> CallToolResult:
                result = self.service.write(
                    command,
                    workbook=workbook,
                    path=path,
                    overwrite=bool(overwrite),
                    sheet=sheet,
                )
                return response.json_result(result, max_chars=max_chars)

            return _guarded(run)

        return excel_write


def _snapshot_result(shots: list[tuple[int, bytes]], max_pages: int) -> CallToolResult:
    if not shots:
        return response.error_result(
            "error", "nothing was rendered", hint="Check the sheet and range with used_range."
        )
    shown = shots[:max_pages]
    content: list[ContentBlock] = []
    for page, png in shown:
        content.append(response.text_block(f"Page {page}"))
        content.append(response.image_block(png, "image/png"))
    if len(shots) > len(shown):
        content.append(
            response.text_block(
                f"[truncated] Showing {len(shown)} of {len(shots)} pages. "
                "Render a smaller range to see the rest."
            )
        )
    return CallToolResult(content=content)


_READ_DESCRIPTION = """\
Read the Excel workbooks open on this PC. Never changes them.

Start with status (open workbooks, active sheet, selection), then sheets or used_range.

Commands, with the fields they take:
- Orient: status, sheets, used_range (sheet)
- Cells: read (sheet, range, formulas) returns rows of values, at most 20,000 cells; \
find (text, sheet, range, match_case, whole_cell, limit) every matching cell; \
summary (sheet, range) count/sum/min/max/average/median
- Objects: tables, table_data (table, visible_only), charts, pivots, pivot_fields (pivot), \
names, calc_mode
- guide: addressing and every excel_exec op. Read it once before your first edit.

Large results are truncated with a note saying how to narrow the request."""

_EXEC_DESCRIPTION = """\
Edit the Excel workbook: apply a list of ops in order.

Each op is {"op": "<name>", ...fields}, addressed by sheet (default: active) and range \
(A1 address). Every op is checked before any runs. Excel can't undo automated edits \
(Ctrl+Z won't work), so if an op fails part-way the cell contents earlier ops overwrote are \
written back; formatting and structural changes are not, and the error lists them in \
not_rolled_back.

Common ops: write, fill_formula, clear, format, sort, autofit, insert_rows, add_sheet, \
create_table, create_chart, create_pivot, pivot_add_field, cond_format.
The full op list with required fields is in excel_read(command="guide").

Example: {"ops": [{"op": "write", "range": "A1", "values": [["Item", "Cost"], ["Pens", 4]]}, \
{"op": "format", "range": "A1:B1", "bold": true}]}"""

_SNAPSHOT_DESCRIPTION = """\
Render a range of the Excel workbook as PNG images, the way it would print, with row \
numbers and column letters. Default: the active sheet's used range. Use it to check \
layout, formatting and charts. Never changes the workbook (page setup is restored)."""

_WRITE_DESCRIPTION = """\
Open, create, save or export an Excel workbook.

- open (path): open a local workbook (starts Excel if needed); activates it if already open.
- new: a new blank workbook.
- save: save the workbook to its current file.
- save_as (path, overwrite): save as .xlsx, .xlsm or .xlsb.
- export_pdf (path, sheet, overwrite): a PDF of the workbook or one sheet.

Writing only works inside the folders the server was configured to allow; otherwise it \
fails with path_not_allowed. Edits go through excel_exec, not this tool."""
