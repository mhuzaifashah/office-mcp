"""Build tool results: compact JSON, size limits, images and actionable errors.

Every tool returns a `CallToolResult` built here, so all apps share one output
shape:

- **Success** is compact JSON text (or plain text for prose such as Markdown).
- **Too-large output** is cut to `max_chars` and followed by a `[truncated]`
  note saying what was dropped and how to ask for a narrower slice. A list
  keeps whole items, so the JSON stays valid.
- **Errors** are tool results with `isError: true` (not protocol errors), so the
  model can read them and correct itself. The body is JSON:
  `{"error": {"code", "message", "retryable", "hint"?, ...details}}`.
"""

from __future__ import annotations

import base64
import json
from typing import Any

from mcp_types import CallToolResult, ContentBlock, ImageContent, TextContent

__all__ = [
    "dumps",
    "error_result",
    "image_block",
    "json_result",
    "text_block",
    "text_result",
]

# A shortened string field keeps this much slack under the limit, so the
# escaping JSON adds back does not push the text over it again.
_FIELD_SLACK = 200


def dumps(value: Any) -> str:
    """Compact JSON: no indentation, real Unicode, unknown objects as strings."""
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=str)


def text_block(text: str) -> TextContent:
    return TextContent(text=text)


def image_block(data: bytes, mime_type: str) -> ImageContent:
    return ImageContent(data=base64.b64encode(data).decode("ascii"), mime_type=mime_type)


def _note(summary: str, more_hint: str | None) -> TextContent:
    text = f"[truncated] {summary}"
    if more_hint:
        text += f" {more_hint}"
    return text_block(text)


def text_result(
    text: str, *, max_chars: int, more_hint: str | None = None, label: str = "text"
) -> CallToolResult:
    """Plain text (Markdown, HTML, a guide), cut at `max_chars` if needed."""
    if len(text) <= max_chars:
        return CallToolResult(content=[text_block(text)])
    summary = f"Showing the first {max_chars:,} of {len(text):,} characters of {label}."
    return CallToolResult(content=[text_block(text[:max_chars]), _note(summary, more_hint)])


def _fit_list(value: list[Any], max_chars: int) -> tuple[str, str | None]:
    """Largest prefix of whole items whose JSON fits in `max_chars`."""
    lo, hi = 0, len(value)
    while lo < hi:
        mid = (lo + hi + 1) // 2
        if len(dumps(value[:mid])) <= max_chars:
            lo = mid
        else:
            hi = mid - 1
    shown = lo
    summary = (
        f"Showing items 1-{shown} of {len(value)}; the rest did not fit the "
        f"{max_chars:,}-character output limit."
    )
    return dumps(value[:shown]), summary


def _fit_dict(value: dict[str, Any], max_chars: int) -> tuple[str, str | None]:
    """Shorten the longest string fields until the JSON fits."""
    out = dict(value)
    cut: list[str] = []
    text = dumps(out)
    for _ in range(len(out)):
        over = len(text) - max_chars
        if over <= 0:
            break
        strings = [k for k, v in out.items() if isinstance(v, str)]
        if not strings:
            break
        key = max(strings, key=lambda k: len(out[k]))
        keep = len(out[key]) - over - _FIELD_SLACK
        if keep <= 0:
            out[key] = ""
        else:
            out[key] = out[key][:keep] + "…"
        if key not in cut:
            cut.append(key)
        text = dumps(out)
    if len(text) > max_chars:
        # Nothing left to shorten field by field: cut the raw text.
        return text[:max_chars], (
            f"Showing the first {max_chars:,} of {len(dumps(value)):,} characters; "
            "the JSON is cut off."
        )
    fields = ", ".join(f"`{k}`" for k in cut)
    return text, f"Shortened {fields} to fit the {max_chars:,}-character output limit."


def json_result(value: Any, *, max_chars: int, more_hint: str | None = None) -> CallToolResult:
    """A JSON value, kept under `max_chars` without breaking the JSON where possible."""
    text = dumps(value)
    if len(text) <= max_chars:
        return CallToolResult(content=[text_block(text)])
    if isinstance(value, list):
        text, summary = _fit_list(value, max_chars)
    elif isinstance(value, dict):
        text, summary = _fit_dict(value, max_chars)
    else:
        text, summary = (
            text[:max_chars],
            f"Showing the first {max_chars:,} of {len(text):,} characters.",
        )
    content: list[ContentBlock] = [text_block(text)]
    if summary:
        content.append(_note(summary, more_hint))
    return CallToolResult(content=content)


def error_result(
    code: str,
    message: str,
    *,
    retryable: bool = False,
    hint: str | None = None,
    details: dict[str, Any] | None = None,
) -> CallToolResult:
    """A tool error the model can act on: a stable code, what happened, what to do."""
    error: dict[str, Any] = {"code": code, "message": message, "retryable": retryable}
    if hint:
        error["hint"] = hint
    if details:
        error.update({k: v for k, v in details.items() if v is not None})
    return CallToolResult(content=[text_block(dumps({"error": error}))], is_error=True)
