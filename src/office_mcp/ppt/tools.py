"""The PowerPoint tools: `ppt_read`, `ppt_exec`, `ppt_snapshot`, `ppt_write`.

Four task-shaped tools, like Word's and Excel's. The ops behind them live in
`dispatch.py`; the guide (`ppt_read(command="guide")`, also the
`office-mcp://ppt/guide` resource) is generated from the same table.
"""

from __future__ import annotations

import logging
import threading
from collections.abc import Callable
from typing import Annotated, Any, Literal

from mcp_types import CallToolResult, ContentBlock
from pydantic import Field
from pydantic.json_schema import SkipJsonSchema

from ..paths import PathNotAllowed
from ..server import response
from ..server.config import ServerConfig
from ..server.tool_definition import ResourceDefinition, ToolDefinition
from .errors import PowerPointError, classify
from .guide import guide_text
from .service import PowerPointService

__all__ = ["READ_COMMANDS", "WRITE_COMMANDS", "PptToolset"]

log = logging.getLogger(__name__)

GUIDE_URI = "office-mcp://ppt/guide"

ReadCommand = Literal[
    "guide", "status", "slides", "slide", "text", "runs", "notes", "layouts", "table", "link"
]
READ_COMMANDS: tuple[str, ...] = ReadCommand.__args__  # type: ignore[attr-defined]

WriteCommand = Literal["open", "new", "save", "save_as", "export_pdf"]
WRITE_COMMANDS: tuple[str, ...] = WriteCommand.__args__  # type: ignore[attr-defined]

_MORE_HINT = "Pass slide to read one slide at a time."

_HINTS: dict[str, str] = {
    "powerpoint_not_running": "Ask the user to start PowerPoint and open their presentation, "
    'or open a file with ppt_write(command="open", path=...).',
    "presentation_not_found": 'Call ppt_read(command="status") for the names of open '
    "presentations.",
    "powerpoint_busy": "A dialog may be open in PowerPoint, or it is busy. Ask the user to "
    "close it, then retry the same call.",
    "path_not_allowed": "Files are only written inside the folders the server was started "
    "with (--save-dir), and pictures only come from local paths (and --image-dir folders, "
    "if set). Ask the user which folder to use.",
    "file_not_found": "Check the path with the user; it must be a local file.",
}
_OP_HINT = (
    'Fix the command or op fields. ppt_read(command="guide") lists every op with its fields; '
    'ppt_read(command="slide", slide=N) lists shape ids.'
)
_BATCH_HINT = (
    "Ops listed in `applied` already changed the presentation. Fix the failing op "
    "(failure.index) and send it with the ops after it, not the whole batch again."
)

Presentation = Annotated[
    str | SkipJsonSchema[None],
    Field(
        description='Presentation file name from ppt_read(command="status"); omit for the '
        "active one."
    ),
]


def _error(exc: BaseException) -> CallToolResult:
    if isinstance(exc, PowerPointError):
        code, retryable = classify(exc)
        details = dict(getattr(exc, "details", None) or {})
        hint = _HINTS.get(code)
        if hint is None:
            hint = _BATCH_HINT if details.get("applied") else _OP_HINT
        return response.error_result(
            code, str(exc), retryable=retryable, hint=hint, details=details or None
        )
    if isinstance(exc, PathNotAllowed):
        details = dict(getattr(exc, "details", None) or {})
        return response.error_result(
            "path_not_allowed", str(exc), hint=_HINTS["path_not_allowed"], details=details or None
        )
    if isinstance(exc, FileNotFoundError):
        return response.error_result("file_not_found", str(exc), hint=_HINTS["file_not_found"])
    if isinstance(exc, (ValueError, TypeError, KeyError)):
        return response.error_result("error", str(exc), hint=_OP_HINT)
    log.exception("unexpected error in a PowerPoint tool")
    return response.error_result("internal_error", f"{type(exc).__name__}: {exc}")


def _guarded(fn: Callable[[], CallToolResult]) -> CallToolResult:
    try:
        return fn()
    except Exception as exc:  # noqa: BLE001 - every failure becomes a readable tool error
        return _error(exc)


def _without_none(**values: Any) -> dict[str, Any]:
    return {k: v for k, v in values.items() if v is not None}


class PptToolset:
    """The PowerPoint app's tools, guide resource and server instructions."""

    app = "ppt"
    instructions = (
        "PowerPoint tools (ppt_*) edit the presentations open in PowerPoint on this PC. Start "
        'with ppt_read(command="status") and ppt_read(command="slides"). Before your first '
        'edit, call ppt_read(command="guide") once: it lists every ppt_exec op with its fields. '
        "Check your work with ppt_snapshot, which labels each shape with its id."
    )

    def __init__(
        self,
        config: ServerConfig,
        *,
        service_factory: Callable[[], PowerPointService] | None = None,
    ) -> None:
        self._config = config
        self._factory = service_factory or (
            lambda: PowerPointService(
                save_dirs=tuple(config.save_dirs), image_dirs=tuple(config.image_dirs)
            )
        )
        self._service: PowerPointService | None = None
        self._lock = threading.Lock()

    @property
    def service(self) -> PowerPointService:
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
                name="ppt_read",
                title="Read PowerPoint presentation",
                description=_READ_DESCRIPTION,
                handler=self._read_handler(),
                read_only=True,
            ),
            ToolDefinition(
                name="ppt_exec",
                title="Edit PowerPoint presentation",
                description=_EXEC_DESCRIPTION,
                handler=self._exec_handler(),
                read_only=False,
                destructive=True,
            ),
            ToolDefinition(
                name="ppt_snapshot",
                title="Render PowerPoint slides",
                description=_SNAPSHOT_DESCRIPTION,
                handler=self._snapshot_handler(),
                read_only=True,
            ),
            ToolDefinition(
                name="ppt_write",
                title="Open, save or export a PowerPoint presentation",
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
                name="ppt-guide",
                title="PowerPoint tools guide",
                description="Addressing, every ppt_read command and every ppt_exec op.",
                mime_type="text/markdown",
                read=guide_text,
            )
        ]

    def _read_handler(self) -> Callable[..., CallToolResult]:
        max_chars = self._config.max_output_chars

        def ppt_read(
            command: Annotated[ReadCommand, Field(description="What to read.")],
            slide: Annotated[
                int | SkipJsonSchema[None],
                Field(ge=1, description="Slide number (slide, text, runs, notes, table, link)."),
            ] = None,
            shape: Annotated[
                int | SkipJsonSchema[None],
                Field(description="Shape id from the slide read (runs, table, link)."),
            ] = None,
            paragraph: Annotated[
                int | SkipJsonSchema[None],
                Field(ge=1, description="Paragraph number in the shape (link)."),
            ] = None,
            presentation: Presentation = None,
        ) -> CallToolResult:
            def run() -> CallToolResult:
                if command == "guide":
                    return response.text_result(
                        guide_text(),
                        max_chars=max_chars,
                        more_hint=f"Read the full guide as the {GUIDE_URI} resource.",
                        label="guide",
                    )
                params = _without_none(slide=slide, shape=shape, paragraph=paragraph)
                result = self.service.read(command, presentation=presentation, **params)
                return response.json_result(result, max_chars=max_chars, more_hint=_MORE_HINT)

            return _guarded(run)

        return ppt_read

    def _exec_handler(self) -> Callable[..., CallToolResult]:
        max_chars = self._config.max_output_chars

        def ppt_exec(
            ops: Annotated[
                list[dict[str, Any]],
                Field(
                    min_length=1,
                    description='Ops to apply in order. Each is {"op": "<name>", ...fields}; '
                    'see ppt_read(command="guide").',
                ),
            ],
            presentation: Presentation = None,
        ) -> CallToolResult:
            def run() -> CallToolResult:
                result = self.service.exec(ops, presentation=presentation)
                return response.json_result(result, max_chars=max_chars)

            return _guarded(run)

        return ppt_exec

    def _snapshot_handler(self) -> Callable[..., CallToolResult]:
        default_max_dim = self._config.snapshot_max_dim
        max_pages = self._config.snapshot_max_pages

        def ppt_snapshot(
            slides: Annotated[
                str | SkipJsonSchema[None],
                Field(description='One slide ("4") or a range ("2-5"). Omit for every slide.'),
            ] = None,
            shape_ids: Annotated[
                bool, Field(description="Outline each shape and label it with its id.")
            ] = True,
            max_dim: Annotated[
                int | SkipJsonSchema[None],
                Field(
                    ge=100,
                    le=4000,
                    description="Long edge in pixels (default: the server's --snapshot-max-dim).",
                ),
            ] = None,
            presentation: Presentation = None,
        ) -> CallToolResult:
            def run() -> CallToolResult:
                shots, total = self.service.snapshot(
                    presentation=presentation,
                    slides=slides,
                    max_dim=max_dim if max_dim is not None else default_max_dim,
                    shape_ids=shape_ids,
                    max_slides=max_pages,
                )
                return _snapshot_result(shots, total)

            return _guarded(run)

        return ppt_snapshot

    def _write_handler(self) -> Callable[..., CallToolResult]:
        max_chars = self._config.max_output_chars

        def ppt_write(
            command: Annotated[WriteCommand, Field(description="What to do.")],
            path: Annotated[
                str | SkipJsonSchema[None],
                Field(
                    description="Local file path: the presentation to open (open), or the "
                    "target file (save_as, export_pdf)."
                ),
            ] = None,
            overwrite: Annotated[
                bool | SkipJsonSchema[None],
                Field(description="Replace an existing file (save_as, export_pdf)."),
            ] = None,
            presentation: Presentation = None,
        ) -> CallToolResult:
            def run() -> CallToolResult:
                result = self.service.write(
                    command, presentation=presentation, path=path, overwrite=bool(overwrite)
                )
                return response.json_result(result, max_chars=max_chars)

            return _guarded(run)

        return ppt_write


def _snapshot_result(shots: list[tuple[int, bytes]], total: int) -> CallToolResult:
    if not shots:
        return response.error_result(
            "error", "nothing was rendered", hint='Check the slide numbers with command="slides".'
        )
    content: list[ContentBlock] = []
    for n, png in shots:
        content.append(response.text_block(f"Slide {n}"))
        content.append(response.image_block(png, "image/png"))
    if total > len(shots):
        nxt = shots[-1][0] + 1
        content.append(
            response.text_block(
                f"[truncated] Showing {len(shots)} of {total} slides. "
                f'Ask for the rest with slides="{nxt}-{shots[0][0] + total - 1}".'
            )
        )
    return CallToolResult(content=content)


_READ_DESCRIPTION = """\
Read the PowerPoint presentations open on this PC. Never changes them.

Start with status (open presentations, slide size, current slide), then slides.

Commands, with the fields they take:
- Orient: status, slides, layouts
- A slide: slide (slide) lists its shapes with ids, positions and text; text (slide, or \
every slide); notes (slide)
- A shape: runs (slide, shape) text with font settings; table (slide, shape); \
link (slide, shape, paragraph)
- guide: addressing and every ppt_exec op. Read it once before your first edit."""

_EXEC_DESCRIPTION = """\
Edit the PowerPoint presentation: apply a list of ops in order.

Each op is {"op": "<name>", ...fields}, addressed by slide number and shape id \
(from ppt_read slide, or the ppt_snapshot labels). Every op is checked before any runs. \
PowerPoint can't roll back automation, so if an op fails part-way the ops before it stay \
applied; the error lists them in `applied`.

Common ops: add_slide, set_placeholder, set_text, format_text, format_paragraph, \
add_text_box, add_shape, fill, move_shape, add_picture, add_table, set_cell, add_chart, \
set_notes, delete_shape.
The full op list with fields is in ppt_read(command="guide").

Example: {"ops": [{"op": "add_slide", "layout": "title_only"}, {"op": "set_placeholder", \
"slide": 2, "placeholder": "title", "text": "Q2 results"}]}"""

_SNAPSHOT_DESCRIPTION = """\
Render slides as PNG images. By default each shape is outlined and labelled with its id, \
so you can find what to edit; shape_ids=false shows the clean slide. Never changes the \
presentation. Images cost tokens: render only the slides you need."""

_WRITE_DESCRIPTION = """\
Open, create, save or export a PowerPoint presentation.

- open (path): open a local presentation (starts PowerPoint if needed).
- new: a new presentation with a title slide.
- save: save to its current file.
- save_as (path, overwrite): save as .pptx or .pptm.
- export_pdf (path, overwrite): a PDF copy.

Writing only works inside the folders the server was configured to allow; otherwise it \
fails with path_not_allowed. Edits go through ppt_exec, not this tool."""
