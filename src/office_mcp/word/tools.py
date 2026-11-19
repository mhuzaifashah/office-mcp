"""The Word tools: `word_read`, `word_exec`, `word_snapshot`, `word_write`.

Four task-shaped tools instead of one per operation. The roughly 120 edit ops
and 50 read commands live behind them, documented in the bundled guide
(`word_read(command="guide")`, also the `office-mcp://word/guide` resource).

Every handler calls one `WordController` method and turns the result (or the
error) into a tool result with `office_mcp.server.response`.
"""

from __future__ import annotations

import logging
import threading
from collections.abc import Callable
from typing import Annotated, Any, Literal

from mcp_types import CallToolResult, ContentBlock
from pydantic import Field
from pydantic.json_schema import SkipJsonSchema

from ..server import response
from ..server.config import ServerConfig
from ..server.tool_definition import ResourceDefinition, ToolDefinition
from .controller import WordController
from .wordcom._guide import skill_body
from .wordcom.exceptions import (
    AmbiguousMatchError,
    ComError,
    OpError,
    WordcomError,
    classify,
)

__all__ = ["READ_COMMANDS", "WRITE_COMMANDS", "WordToolset"]

log = logging.getLogger(__name__)

GUIDE_URI = "office-mcp://word/guide"

# Every `word_read` command, matching `dispatch_read` (a test keeps them in sync).
ReadCommand = Literal[
    "guide",
    "status",
    "outline",
    "paragraphs",
    "sections",
    "stats",
    "styles",
    "find",
    "find_paragraphs",
    "location",
    "nearest_heading",
    "read_text",
    "read_section",
    "between",
    "read_bookmark",
    "read_cc",
    "to_markdown",
    "to_html",
    "digest",
    "table_list",
    "table_read",
    "table_records",
    "images",
    "read_image",
    "shapes",
    "equations",
    "charts",
    "footnotes",
    "endnotes",
    "hyperlinks",
    "fields",
    "comments",
    "revisions",
    "track",
    "properties",
    "variables",
    "theme",
    "themes",
    "format_info",
    "list_levels",
    "proofing",
    "lint",
    "checkpoint",
    "diff",
]
READ_COMMANDS: tuple[str, ...] = ReadCommand.__args__  # type: ignore[attr-defined]

WriteCommand = Literal["open", "save", "save_as", "export_pdf", "track"]
WRITE_COMMANDS: tuple[str, ...] = WriteCommand.__args__  # type: ignore[attr-defined]

# Reads whose payload is prose: returned as plain text, not a JSON string.
_TEXT_READS: dict[str, str] = {
    "guide": "guide",
    "to_markdown": "markdown",
    "to_html": "html",
    "digest": "digest",
}

# What to tell the model when a read is too large to return whole.
_MORE_HINTS: dict[str, str] = {
    "paragraphs": "Pass start and count to page through the paragraphs.",
    "outline": "Without all_paragraphs the outline lists headings only; "
    "use paragraphs with start and count for the rest.",
    "to_markdown": "Pass within=<heading anchor id> to convert one section, "
    "or use digest for a budgeted overview.",
    "to_html": "Pass within=<heading anchor id> to convert one section.",
    "guide": "Read the full guide as the office-mcp://word/guide resource.",
}
_DEFAULT_MORE_HINT = (
    "Narrow the request (within, start and count, or an anchor_id) or use digest for an overview."
)

# What the model should do next, per error code.
_HINTS: dict[str, str] = {
    "anchor_not_found": "Positional ids (heading:N, para:N, image:N, ...) shift after edits. "
    "Re-read outline or paragraphs, then retry with a fresh id.",
    "style_not_found": 'List the styles defined in this document with word_read(command="styles").',
    "ambiguous_match": "Pick one of `matches`: add occurrence=N (1-based) to the op, "
    "or all=true to change every match.",
    "replace_verification": "The text moved under the match. Re-read it, then target it "
    'by anchor id, or scope the op with in="table:N:R:C".',
    "word_busy": "A dialog may be open in Word, or Word is still busy. "
    "Wait a moment and retry the same call.",
    "word_not_running": "Ask the user to start Word and open their document, or open a "
    'file with word_write(command="open", path=...).',
    "document_not_found": 'Call word_read(command="status") for the names of open documents.',
    "excel_not_available": "Charts need Microsoft Excel. Render the chart as an image and "
    "use insert_image instead.",
    "path_not_allowed": "Saving only writes inside the folders the server was started with "
    "(--save-dir). Ask the user which allowed folder to use. For images, pass base64 "
    "instead of a path.",
    "file_not_found": "Check the path with the user; it must be a local file.",
}
_WORD_NOT_RUNNING_READ_ONLY = "Ask the user to start Word and open their document."
_OP_ERROR_HINT = (
    'Fix the command or op fields. word_read(command="guide") lists every op with its '
    "required fields."
)

# Optional parameters are typed `X | SkipJsonSchema[None]`: a client may still
# send null, but the schema shows a plain `X`. `anyOf [X, null]` trips some
# clients (Gemini models) and adds bytes; Codex strips descriptions from any
# tool schema over 5,000 bytes.
DocName = Annotated[
    str | SkipJsonSchema[None],
    Field(description='Document name from word_read(command="status"); omit for the active one.'),
]


def _error(exc: BaseException, *, read_only: bool = False) -> CallToolResult:
    """Turn an exception into an actionable tool error.

    With `read_only`, hints never point at tools the server did not register.
    """
    if isinstance(exc, WordcomError):
        code, retryable = classify(exc)
        hint = _HINTS.get(code)
        if read_only and code == "word_not_running":
            hint = _WORD_NOT_RUNNING_READ_ONLY
        if hint is None and isinstance(exc, OpError):
            hint = _OP_ERROR_HINT
        details: dict[str, Any] = {
            "failure": getattr(exc, "failure", None),
            "ops_run": getattr(exc, "ops_run", None),
            "rolled_back": getattr(exc, "rolled_back", None),
        }
        if isinstance(exc, AmbiguousMatchError):
            details["matches"] = exc.matches
        if isinstance(exc, ComError) and exc.hresult is not None:
            details["hresult"] = f"0x{exc.hresult & 0xFFFFFFFF:08X}"
        return response.error_result(
            code, str(exc), retryable=retryable, hint=hint, details=details
        )
    if isinstance(exc, FileNotFoundError):
        return response.error_result("file_not_found", str(exc), hint=_HINTS["file_not_found"])
    if isinstance(exc, (ValueError, TypeError, IsADirectoryError)):
        return response.error_result("error", str(exc), hint=_OP_ERROR_HINT)
    log.exception("unexpected error in a Word tool")
    return response.error_result("internal_error", f"{type(exc).__name__}: {exc}")


def _guarded(fn: Callable[[], CallToolResult], *, read_only: bool = False) -> CallToolResult:
    try:
        return fn()
    except Exception as exc:  # noqa: BLE001 — every failure becomes a readable tool error
        return _error(exc, read_only=read_only)


def _without_none(**values: Any) -> dict[str, Any]:
    return {k: v for k, v in values.items() if v is not None}


class WordToolset:
    """The Word app's tools, resources and server instructions.

    The `WordController` (and with it the COM thread) is created on the first
    tool call, so an idle Word toolset costs nothing. `close` stops that thread;
    it never closes Word.
    """

    app = "word"
    instructions = (
        "Word tools (word_*) edit the Word documents open on this PC. Start with "
        'word_read(command="status") and word_read(command="outline"). Before your first '
        'edit, call word_read(command="guide") once: it explains anchor ids and lists every '
        "word_exec op with its fields. Make every edit with word_exec; a failed batch is "
        "rolled back as a whole."
    )

    def __init__(
        self,
        config: ServerConfig,
        *,
        controller_factory: Callable[[], WordController] | None = None,
    ) -> None:
        self._config = config
        self._factory = controller_factory or (
            lambda: WordController(
                save_dirs=list(config.save_dirs), image_dirs=list(config.image_dirs)
            )
        )
        self._controller: WordController | None = None
        self._lock = threading.Lock()

    @property
    def controller(self) -> WordController:
        with self._lock:
            if self._controller is None:
                self._controller = self._factory()
            return self._controller

    def close(self) -> None:
        with self._lock:
            if self._controller is not None:
                self._controller.close()
                self._controller = None

    # -- tools -------------------------------------------------------------

    def tools(self) -> list[ToolDefinition]:
        return [
            ToolDefinition(
                name="word_read",
                title="Read Word document",
                description=_READ_DESCRIPTION,
                handler=self._read_handler(),
                read_only=True,
            ),
            ToolDefinition(
                name="word_exec",
                title="Edit Word document",
                description=_EXEC_DESCRIPTION,
                handler=self._exec_handler(),
                read_only=False,
                destructive=True,
            ),
            ToolDefinition(
                name="word_snapshot",
                title="Render Word pages",
                description=_SNAPSHOT_DESCRIPTION,
                handler=self._snapshot_handler(),
                read_only=True,
            ),
            ToolDefinition(
                name="word_write",
                title="Open, save or export a Word document",
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
                name="word-guide",
                title="Word tools guide",
                description="Anchor ids, every word_read command and every word_exec op.",
                mime_type="text/markdown",
                read=lambda: skill_body("mcp"),
            )
        ]

    # -- handlers ----------------------------------------------------------

    def _read_handler(self) -> Callable[..., CallToolResult]:
        max_chars = self._config.max_output_chars

        def word_read(
            command: Annotated[ReadCommand, Field(description="What to read.")],
            doc: DocName = None,
            anchor_id: Annotated[
                str | SkipJsonSchema[None],
                Field(
                    description="Anchor id such as heading:12, para:3, range:40-55, "
                    "bookmark:NAME, cc:NAME, table:1:2:3 or image:1. For read_text, "
                    "read_section, location, nearest_heading, format_info, list_levels, "
                    "read_image."
                ),
            ] = None,
            text: Annotated[
                str | SkipJsonSchema[None],
                Field(description="Text to look for (find, find_paragraphs)."),
            ] = None,
            name: Annotated[
                str | SkipJsonSchema[None],
                Field(description="Bookmark or content-control name (read_bookmark, read_cc)."),
            ] = None,
            heading: Annotated[
                str | SkipJsonSchema[None],
                Field(description="Heading text, instead of anchor_id (read_section)."),
            ] = None,
            start: Annotated[
                int | SkipJsonSchema[None],
                Field(ge=1, description="First paragraph, 1-based (paragraphs)."),
            ] = None,
            count: Annotated[
                int | SkipJsonSchema[None],
                Field(ge=1, description="How many paragraphs (paragraphs)."),
            ] = None,
            all_paragraphs: Annotated[
                bool | SkipJsonSchema[None],
                Field(description="List every paragraph, not just headings (outline)."),
            ] = None,
            in_anchor: Annotated[
                str | SkipJsonSchema[None],
                Field(description="Only search inside this anchor (find)."),
            ] = None,
            limit: Annotated[
                int | SkipJsonSchema[None],
                Field(ge=1, description="Most candidates, default 5 (find_paragraphs)."),
            ] = None,
            min_score: Annotated[
                float | SkipJsonSchema[None],
                Field(
                    ge=0, le=1, description="Lowest similarity, 0-1, default 0.6 (find_paragraphs)."
                ),
            ] = None,
            direction: Annotated[
                Literal["before", "after"] | SkipJsonSchema[None],
                Field(description="Search direction (nearest_heading)."),
            ] = None,
            view: Annotated[
                Literal["raw", "final", "original", "segments"] | SkipJsonSchema[None],
                Field(
                    description="raw text, as if tracked changes were accepted (final) or "
                    "rejected (original), or per-revision segments (read_text)."
                ),
            ] = None,
            start_anchor: Annotated[
                str | SkipJsonSchema[None],
                Field(description="Anchor the span starts at (between)."),
            ] = None,
            end_anchor: Annotated[
                str | SkipJsonSchema[None], Field(description="Anchor the span ends at (between).")
            ] = None,
            inclusive: Annotated[
                bool | SkipJsonSchema[None],
                Field(description="Include both end anchors (between)."),
            ] = None,
            within: Annotated[
                str | SkipJsonSchema[None],
                Field(
                    description="Limit to one anchor, usually a heading's section "
                    "(to_markdown, to_html, lint, checkpoint)."
                ),
            ] = None,
            budget: Annotated[
                int | SkipJsonSchema[None],
                Field(ge=200, description="Token budget, default 6000 (digest)."),
            ] = None,
            depth: Annotated[
                int | SkipJsonSchema[None],
                Field(ge=1, le=9, description="Deepest heading level (digest)."),
            ] = None,
            table: Annotated[
                int | SkipJsonSchema[None],
                Field(ge=1, description="Table number, 1-based (table_read, table_records)."),
            ] = None,
            rules: Annotated[
                list[str] | SkipJsonSchema[None],
                Field(description="Lint rules to run; default all (lint)."),
            ] = None,
            include: Annotated[
                Literal["text", "text+style", "text+format"] | SkipJsonSchema[None],
                Field(description="What the fingerprint covers (checkpoint)."),
            ] = None,
            checkpoint: Annotated[
                Any,
                Field(description="A checkpoint to compare with the document now (diff)."),
            ] = None,
            cp_a: Annotated[Any, Field(description="Older checkpoint, with cp_b (diff).")] = None,
            cp_b: Annotated[Any, Field(description="Newer checkpoint, with cp_a (diff).")] = None,
        ) -> CallToolResult:
            params = _without_none(
                doc=doc,
                anchor_id=anchor_id,
                text=text,
                name=name,
                heading=heading,
                start=start,
                count=count,
                all_paragraphs=all_paragraphs,
                in_anchor=in_anchor,
                limit=limit,
                min_score=min_score,
                direction=direction,
                view=view,
                start_anchor=start_anchor,
                end_anchor=end_anchor,
                inclusive=inclusive,
                within=within,
                budget=budget,
                depth=depth,
                table=table,
                rules=rules,
                include=include,
                checkpoint=checkpoint,
                cp_a=cp_a,
                cp_b=cp_b,
            )

            def run() -> CallToolResult:
                result = self.controller.read(command, **params)
                more = _MORE_HINTS.get(command, _DEFAULT_MORE_HINT)
                if command == "read_image":
                    return _image_result(result)
                if command in _TEXT_READS and isinstance(result, dict):
                    return response.text_result(
                        str(result.get(_TEXT_READS[command], "")),
                        max_chars=max_chars,
                        more_hint=more,
                        label=command,
                    )
                return response.json_result(result, max_chars=max_chars, more_hint=more)

            return _guarded(run, read_only=self._config.read_only)

        return word_read

    def _exec_handler(self) -> Callable[..., CallToolResult]:
        max_chars = self._config.max_output_chars

        def word_exec(
            ops: Annotated[
                list[dict[str, Any]],
                Field(
                    min_length=1,
                    description='Ops to apply in order. Each is {"op": "<name>", ...fields}; '
                    'see word_read(command="guide").',
                ),
            ],
            label: Annotated[
                str | SkipJsonSchema[None],
                Field(description="Name of the undo step the user sees in Word."),
            ] = None,
            tracked: Annotated[
                bool, Field(description="Record the whole batch as tracked changes.")
            ] = False,
            doc: DocName = None,
        ) -> CallToolResult:
            def run() -> CallToolResult:
                result = self.controller.exec(ops, doc=doc, label=label, tracked=tracked)
                return response.json_result(result, max_chars=max_chars)

            return _guarded(run, read_only=self._config.read_only)

        return word_exec

    def _snapshot_handler(self) -> Callable[..., CallToolResult]:
        default_max_dim = self._config.snapshot_max_dim
        max_pages = self._config.snapshot_max_pages

        def word_snapshot(
            pages: Annotated[
                str | SkipJsonSchema[None],
                Field(description='One page ("4") or a range ("2-5"). Omit for every page.'),
            ] = None,
            anchor: Annotated[
                str | SkipJsonSchema[None],
                Field(
                    description="Render the page(s) this anchor is on instead; a heading "
                    "covers its whole section."
                ),
            ] = None,
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
            markup: Annotated[
                Literal["none", "all"],
                Field(description="all = show tracked changes and comments."),
            ] = "none",
            doc: DocName = None,
        ) -> CallToolResult:
            def run() -> CallToolResult:
                if pages and anchor:
                    raise OpError("pass pages or anchor, not both")
                shots = self.controller.snapshot(
                    doc=doc,
                    pages=pages,
                    anchor=anchor,
                    dpi=dpi,
                    markup=markup,
                    max_dim=max_dim if max_dim is not None else default_max_dim,
                )
                return _snapshot_result(shots, max_pages)

            return _guarded(run, read_only=self._config.read_only)

        return word_snapshot

    def _write_handler(self) -> Callable[..., CallToolResult]:
        max_chars = self._config.max_output_chars

        def word_write(
            command: Annotated[WriteCommand, Field(description="What to do.")],
            path: Annotated[
                str | SkipJsonSchema[None],
                Field(
                    description="Local file path: the document to open (open), or the "
                    "target file (save_as, export_pdf)."
                ),
            ] = None,
            overwrite: Annotated[
                bool | SkipJsonSchema[None],
                Field(description="Replace an existing file (save_as)."),
            ] = None,
            from_page: Annotated[
                int | SkipJsonSchema[None],
                Field(ge=1, description="First page to export (export_pdf)."),
            ] = None,
            to_page: Annotated[
                int | SkipJsonSchema[None],
                Field(ge=1, description="Last page to export (export_pdf)."),
            ] = None,
            on: Annotated[
                bool | SkipJsonSchema[None],
                Field(description="Turn Track Changes on or off (track)."),
            ] = None,
            doc: DocName = None,
        ) -> CallToolResult:
            def run() -> CallToolResult:
                if command == "open":
                    if not path:
                        raise OpError("write command 'open' requires 'path'")
                    result = self.controller.open(path)
                else:
                    params = _without_none(
                        doc=doc,
                        path=path,
                        overwrite=overwrite,
                        from_page=from_page,
                        to_page=to_page,
                        on=on,
                    )
                    result = self.controller.write(command, **params)
                return response.json_result(result, max_chars=max_chars)

            return _guarded(run, read_only=self._config.read_only)

        return word_write


def _image_result(result: dict[str, Any]) -> CallToolResult:
    """`read_image` as a real image block plus its metadata."""
    import base64  # noqa: PLC0415

    data = base64.b64decode(result.get("base64", ""))
    meta = {k: v for k, v in result.items() if k != "base64"}
    return CallToolResult(
        content=[
            response.text_block(response.dumps(meta)),
            response.image_block(data, str(result.get("mime") or "image/png")),
        ]
    )


def _snapshot_result(shots: list[tuple[int, bytes]], max_pages: int) -> CallToolResult:
    """One `Page N` label and one PNG per page, capped at `max_pages`."""
    if not shots:
        return response.error_result(
            "error", "nothing was rendered", hint="Check the pages range against stats."
        )
    shown = shots[:max_pages]
    content: list[ContentBlock] = []
    for page, png in shown:
        content.append(response.text_block(f"Page {page}"))
        content.append(response.image_block(png, "image/png"))
    if len(shots) > len(shown):
        rest = [p for p, _ in shots[len(shown) :]]
        content.append(
            response.text_block(
                f"[truncated] Showing {len(shown)} of {len(shots)} pages. "
                f'Ask for the rest with pages="{rest[0]}-{rest[-1]}".'
            )
        )
    return CallToolResult(content=content)


_READ_DESCRIPTION = """\
Read the Word document open on this PC. Never changes it.

Start with status (open documents), then outline (headings with their anchor ids).
Anchor ids from these reads are what word_exec ops target.

Commands, with the fields they take:
- Orient: status, outline (all_paragraphs), paragraphs (start, count), sections, stats, styles
- Locate: find (text, in_anchor) returns range: ids; find_paragraphs (text, limit, min_score) \
fuzzy; location (anchor_id); nearest_heading (anchor_id, direction)
- Text: read_text (anchor_id, view), read_section (anchor_id or heading), between \
(start_anchor, end_anchor, inclusive), read_bookmark (name), read_cc (name)
- Whole document: to_markdown (within), to_html (within), digest (budget, depth) for a large \
document
- Tables: table_list, table_read (table), table_records (table)
- Objects: images, read_image (anchor_id) returns the picture, shapes, equations, charts, \
footnotes, endnotes, hyperlinks, fields, comments, revisions, track
- Design: properties, variables, theme, themes, format_info (anchor_id), list_levels (anchor_id)
- Quality: proofing, lint (rules, within)
- Changes: checkpoint (include, within) returns a token; diff (checkpoint, or cp_a and cp_b) \
lists what changed since
- guide: how anchors work and every word_exec op. Read it once before your first edit.

Large results are truncated with a note saying how to narrow the request."""

_EXEC_DESCRIPTION = """\
Edit the Word document: apply a list of ops as one undo step for the user.

Each op is an object {"op": "<name>", ...fields}. Ops target anchor ids from word_read \
(heading:N, para:N, range:S-E, bookmark:NAME, table:N:R:C, start, end ...). Ops run in \
order and each resolves its anchor against the document as it is at that moment.

If any op fails, every op before it is rolled back, so the document is unchanged. The \
error names the failing op (failure.index) and says how to fix it; fix that op and resend \
the whole batch. On success the result lists each op's outputs (new table index, pin id, \
shape id ...) and warns about fields an op ignored.

Common ops: insert_paragraph, insert_block, insert_markdown, replace, find_replace, \
delete_paragraph, apply_style, format_run, format_paragraph, apply_list, create_table, \
set_cell, insert_image, add_comment, write_header, insert_toc.
The full op list with required fields is in word_read(command="guide"); read it once \
before your first edit.

Example: {"ops": [{"op": "insert_paragraph", "anchor_id": "heading:4", \
"text": "New paragraph."}, {"op": "find_replace", "find": "Q3", "text": "Q4", "all": true}]}"""

_SNAPSHOT_DESCRIPTION = """\
Render pages of the Word document as PNG images, so you can check layout, styling and \
placement the way the user sees it. Never changes the document.

Pass pages ("4" or "2-5") or an anchor (renders the page(s) it sits on). Without either, \
every page is rendered, up to the server's page limit. Use markup="all" to see tracked \
changes and comments. Images cost tokens: render only the pages you need."""

_WRITE_DESCRIPTION = """\
Open, save or export a Word document, or turn Track Changes on or off.

- open (path): open a local .docx in Word (starts Word if needed); activates it if already open.
- save: save the document to its current file.
- save_as (path, overwrite): save to a new .docx.
- export_pdf (path, from_page, to_page): write a PDF copy.
- track (on): turn Track Changes on or off.

Saving and exporting only write inside the folders the server was configured to allow; \
otherwise they fail with path_not_allowed. Edits go through word_exec, not this tool."""
