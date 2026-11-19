"""Block-level Markdown — a deliberately *tiny* subset mapped to Word structure.

`insert_markdown` lets an agent hand wordcom a chunk of Markdown and get back
real Word paragraphs, headings, and lists in one op. This module does the
COM-free part: classify each source line into a `Block`. The actual insertion
(and the inline `**bold**`/`*italic*` parsing, which is reused as-is from
`_inline_runs.parse_markup`) lives in `_anchors.insert_markdown`.

The dialect is **a documented subset, not CommonMark** — Word is not Markdown.
Supported, on its own line:

- ``# `` / ``## `` / ``### `` → `Heading 1` / `Heading 2` / `Heading 3`
  (1–3 hashes only; ``#### `` and deeper fall through to a plain paragraph).
- ``- `` / ``* `` → a bulleted list item.
- ``1. `` (one-or-more digits, a dot, a space) → a numbered list item.
- a blank line separates paragraphs; consecutive plain lines join into one
  `Normal` paragraph (a soft wrap, CommonMark-style).
- a GFM pipe table (a header row, a ``| --- |`` separator row, then body
  rows) → a Word table. Cells keep their inline markup; a row with extra
  cells widens the table.
- a horizontal rule (``---``, ``***``, ``___``) separates paragraphs and
  adds nothing.
- inline spans (`**bold**`, `*italic*`, `***both***`) are left intact in the
  block text — `insert_block` runs them through `_inline_runs.parse_markup` downstream.

Explicitly **out of scope** (deferred until asked): code fences, nested /
mixed lists, block quotes, links, images. Anything that isn't a heading,
list item, or blank line is plain paragraph text — so a stray ``> quote`` or
```` ``` ```` fence lands as literal text rather than being silently dropped.

This module is pure and COM-free so the parser is unit-testable off-Windows,
mirroring `_inline_runs.py`.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

# Block kinds. `heading` carries a 1–3 `level`; the rest leave it `None`.
HEADING = "heading"
BULLET = "bullet"
NUMBER = "number"
NORMAL = "normal"
TABLE = "table"

# Marker patterns, matched after stripping leading whitespace (the subset is
# flat — there are no nested list levels in v1, so indentation is insignificant).
_HEADING = re.compile(r"(#{1,3}) +(.*)$")
_BULLET = re.compile(r"[-*] +(.*)$")
_NUMBER = re.compile(r"\d+\. +(.*)$")
_THEMATIC_BREAK = re.compile(r"(-{3,}|\*{3,}|_{3,})")
_TABLE_SEPARATOR = re.compile(r"\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?")


@dataclass(frozen=True)
class Block:
    """One source-derived block in document order.

    `kind` is `heading`/`bullet`/`number`/`normal`; `text` is the block's content
    with inline markup left intact for the run parser. `level` is the heading
    level (1–3) for a heading, `None` otherwise.
    """

    kind: str
    text: str
    level: int | None = None
    # A table's rows (header first), each a tuple of cell texts.
    rows: tuple[tuple[str, ...], ...] | None = None


def _table_cells(line: str) -> list[str]:
    """Split one pipe-table row into cell texts (``\\|`` is a literal pipe)."""
    body = line.strip()
    if body.startswith("|"):
        body = body[1:]
    if body.endswith("|") and not body.endswith("\\|"):
        body = body[:-1]
    cells = re.split(r"(?<!\\)\|", body)
    return [c.strip().replace("\\|", "|") for c in cells]


def _table_block(lines: list[str]) -> Block | None:
    """A TABLE block from consecutive ``|`` lines, or None if they aren't a GFM table."""
    if len(lines) < 2 or not _TABLE_SEPARATOR.fullmatch(lines[1].strip()):
        return None
    cells = [_table_cells(lines[0]), *(_table_cells(line) for line in lines[2:])]
    # A row with more cells than the header widens the table rather than losing
    # data; short rows are padded.
    width = max(len(row) for row in cells)
    rows = tuple(tuple(row + [""] * (width - len(row))) for row in cells)
    return Block(TABLE, "", rows=rows)


def parse_markdown(md: str) -> list[Block]:
    """Classify `md` into a flat list of `Block`s in document order.

    A blank line flushes the current plain-text paragraph; consecutive plain
    lines join with a single space. Headings and list items each flush the
    pending paragraph and emit their own block. Returns ``[]`` for empty input.
    """
    blocks: list[Block] = []
    para: list[str] = []

    def flush() -> None:
        if para:
            blocks.append(Block(NORMAL, " ".join(para)))
            para.clear()

    lines = md.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1
        if not line or _THEMATIC_BREAK.fullmatch(line):
            flush()  # a horizontal rule has no Word equivalent: just a break
            continue
        if line.startswith("|"):
            run = [line]
            while i < len(lines) and lines[i].strip().startswith("|"):
                run.append(lines[i].strip())
                i += 1
            table = _table_block(run)
            if table is not None:
                flush()
                blocks.append(table)
            else:
                para.extend(run)  # not a table: literal text, as before
            continue
        if m := _HEADING.fullmatch(line):
            flush()
            blocks.append(Block(HEADING, m.group(2), level=len(m.group(1))))
        elif m := _BULLET.fullmatch(line):
            flush()
            blocks.append(Block(BULLET, m.group(1)))
        elif m := _NUMBER.fullmatch(line):
            flush()
            blocks.append(Block(NUMBER, m.group(1)))
        else:
            para.append(line)
    flush()
    return blocks
