"""Fast paragraph lookups: by index, by position, by span and by outline level.

Stepping through `doc.Paragraphs` from Python costs one cross-process COM call
per paragraph, and reading `OutlineLevel` costs another 10-20 ms each, so a
walk of a 40-page document takes 10-30 seconds. Word answers the same
questions itself in one call or a handful:

- `Paragraphs(n)` jumps to paragraph n.
- `Range(0, pos + 1).Paragraphs.Count` is the index of the paragraph holding
  `pos`.
- `Range(a, b).Paragraphs` walks only the paragraphs in a span.
- Find with a paragraph-format filter lists the paragraphs at one outline level.

Find only narrows the candidates. A match can begin on the previous
paragraph's mark, and Find's idea of a paragraph's level can differ from
`OutlineLevel` (a level-1 heading found only by the level-2 search), so every
level is searched and each candidate's `OutlineLevel` is read directly before
it counts as a heading. Where Find is unavailable (its `Execute` returns
something other than a bool), heading lookups walk the paragraphs instead.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from typing import Any

from ..constants import WdCollapseDirection

__all__ = [
    "count",
    "headings",
    "index_at",
    "paragraph",
    "paragraphs_between",
    "paragraphs_starting_in",
    "texts",
]

_BODY_LEVEL = 10  # WdOutlineLevel for body text; 1-9 are heading levels
_WD_FIND_STOP = 0
# A paragraph ends at a CR, at a table cell's CR + end mark, or at a section
# break, which is a paragraph of its own (no CR after it). A manual page break
# is followed by the CR of its paragraph.
_PARAGRAPH_END = re.compile(r"\r\x07|\r|\x0c(?!\r)")
_TEXT_SAMPLES = 8


def count(doc_com: Any) -> int:
    """Number of paragraphs in the main story."""
    return int(doc_com.Paragraphs.Count)


def paragraph(doc_com: Any, index: int) -> Any | None:
    """Paragraph `index` (1-based), or None when there is no such paragraph."""
    if index < 1:
        return None
    try:
        return doc_com.Paragraphs(index)
    except Exception:  # noqa: BLE001 - Word raises a COM error past the last paragraph
        return None


def index_at(doc_com: Any, pos: int) -> int:
    """1-based index of the paragraph that holds document position `pos`."""
    pos = max(0, int(pos))
    end = int(doc_com.Content.End)
    if end > 0 and pos >= end:
        pos = end - 1
    return int(doc_com.Range(0, pos + 1).Paragraphs.Count)


def paragraphs_between(doc_com: Any, first: int, last: int) -> Iterator[tuple[int, Any]]:
    """`(index, paragraph)` for paragraphs `first`..`last` (1-based, inclusive)."""
    first = max(1, first)
    if last < first:
        return
    head = paragraph(doc_com, first)
    if head is None:
        return
    tail = paragraph(doc_com, last)
    if tail is None:
        last = count(doc_com)
        tail = paragraph(doc_com, last)
        if tail is None:
            return
    span = doc_com.Range(int(head.Range.Start), int(tail.Range.End))
    for offset, para in enumerate(span.Paragraphs):
        if first + offset > last:
            break
        yield first + offset, para


def paragraphs_starting_in(doc_com: Any, start: int, end: int) -> Iterator[tuple[int, Any]]:
    """`(index, paragraph)` for every paragraph whose start lies in `[start, end)`."""
    if end <= start:
        return
    span = doc_com.Range(start, end)
    index = index_at(doc_com, start)
    for para in span.Paragraphs:
        p_start = int(para.Range.Start)
        if p_start >= end:
            break
        if p_start >= start:
            yield index, para
        index += 1


def _level(para: Any) -> int:
    try:
        return int(para.OutlineLevel)
    except Exception:  # noqa: BLE001 - unreadable paragraphs count as body text
        return _BODY_LEVEL


def headings(
    doc_com: Any,
    *,
    start: int = 0,
    end: int | None = None,
    max_level: int = 9,
    first: bool = False,
) -> list[tuple[int, int, Any]]:
    """Heading paragraphs that start in `[start, end)`, in document order.

    Each row is `(index, level, paragraph)`. Only levels `1..max_level` count.
    With `first=True` only the earliest such heading is returned (or none).
    """
    doc_end = int(doc_com.Content.End)
    stop = doc_end if end is None else min(end, doc_end)
    if stop <= start:
        return []
    found = _headings_by_find(doc_com, start, stop, max_level, first)
    if found is None:
        found = _headings_by_walk(doc_com, start, stop, max_level, first)
    return found


def _headings_by_walk(
    doc_com: Any, start: int, end: int, max_level: int, first: bool
) -> list[tuple[int, int, Any]]:
    out: list[tuple[int, int, Any]] = []
    for index, para in enumerate(doc_com.Paragraphs, start=1):
        p_start = int(para.Range.Start)
        if p_start < start:
            continue
        if p_start >= end:
            break
        level = _level(para)
        if level <= max_level:
            out.append((index, level, para))
            if first:
                break
    return out


def _headings_by_find(
    doc_com: Any, start: int, end: int, max_level: int, first: bool
) -> list[tuple[int, int, Any]] | None:
    """Heading candidates from Find, verified by reading each one's level.

    Returns None when Find is unavailable, so the caller can walk instead.
    """
    found: dict[int, tuple[int, Any]] = {}  # paragraph start -> (level, paragraph)
    for level in range(1, 10):
        rng = doc_com.Range(start, end)
        finder = rng.Find
        finder.ClearFormatting()
        finder.Text = ""
        finder.Forward = True
        finder.Wrap = _WD_FIND_STOP
        finder.Format = True
        finder.MatchWildcards = False
        finder.ParagraphFormat.OutlineLevel = level
        try:
            last = -1
            while True:
                hit = finder.Execute()
                if not isinstance(hit, bool):
                    return None
                if not hit:
                    break
                hit_start = int(rng.Start)
                if hit_start <= last or hit_start >= end:
                    break
                last = hit_start
                matched = False
                for para in rng.Paragraphs:
                    p_start = int(para.Range.Start)
                    if p_start < start or p_start >= end or p_start in found:
                        continue
                    p_level = _level(para)
                    if p_level <= max_level:
                        found[p_start] = (p_level, para)
                        matched = True
                if first and matched:
                    break
                if int(rng.End) >= end:
                    break
                rng.Collapse(int(WdCollapseDirection.END))
        finally:
            # Leave Word's shared Find state as a user expects it.
            finder.ClearFormatting()
            finder.Text = ""
    starts = sorted(found)
    if first:
        starts = starts[:1]
    out: list[tuple[int, int, Any]] = []
    prev_start: int | None = None
    prev_index = 0
    for p_start in starts:
        if prev_start is None:
            index = index_at(doc_com, p_start)
        else:
            # Paragraphs from the previous heading through this one, inclusive.
            index = prev_index + int(doc_com.Range(prev_start, p_start + 1).Paragraphs.Count) - 1
        level, para = found[p_start]
        out.append((index, level, para))
        prev_start, prev_index = p_start, index
    return out


def texts(doc_com: Any) -> list[str] | None:
    """Every paragraph's raw text from one read of the document, or None.

    `Content.Text` is split at paragraph ends. The split is trusted only when
    its length matches `Paragraphs.Count` and a sample of paragraphs read
    directly agrees; otherwise None, so the caller can walk instead.
    """
    try:
        full = str(doc_com.Content.Text or "")
        total = count(doc_com)
    except Exception:  # noqa: BLE001 - fall back to a walk
        return None
    parts = _PARAGRAPH_END.split(full)
    if parts and parts[-1] == "":
        parts.pop()
    if len(parts) != total or total == 0:
        return None
    step = max(1, total // _TEXT_SAMPLES)
    for index in {1, total, *range(1, total + 1, step)}:
        para = paragraph(doc_com, index)
        if para is None:
            return None
        direct = str(para.Range.Text or "").rstrip("\r\x07").replace("\x0c", "")
        if direct != parts[index - 1].replace("\x0c", ""):
            return None
    return parts
