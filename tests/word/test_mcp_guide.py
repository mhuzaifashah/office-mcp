"""The MCP guide must document the real op and read-command vocabulary (W14).

The guide is the only reference an MCP client gets, so it is checked against
the tables the code itself validates with: every exec op needs a row in the op
reference, every required field must be marked bold in that row, no row may
name an op that doesn't exist, and every `read` command must be mentioned.
"""

from __future__ import annotations

import inspect
import re

from office_mcp.word import tools as _tools
from office_mcp.word.wordcom._guide import skill_body
from office_mcp.word.wordcom.exceptions import classify
from office_mcp.word.wordcom.runtime import _dispatch
from office_mcp.word.wordcom.runtime._ops import OP_REQUIRED_FIELDS

GUIDE = skill_body("mcp")
ROWS = [line for line in GUIDE.splitlines() if line.startswith("| `")]


def _op_cell(row: str) -> list[str]:
    """Op names in a table row's first column (rows may group several ops)."""
    first = row.split("|")[1]
    return re.findall(r"`([a-z_]+)`", first)


def _rows_for(op: str) -> list[str]:
    return [r for r in ROWS if op in _op_cell(r)]


def _read_commands() -> set[str]:
    src = inspect.getsource(_dispatch.dispatch_read)
    return set(re.findall(r'command == "([a-z_]+)"', src))


def _extra_codes() -> set[str]:
    """Codes from `classify` that have no hint of their own."""
    src = inspect.getsource(classify)
    return set(re.findall(r'return "([a-z_]+)", ', src))


def test_every_exec_op_has_a_row():
    missing = sorted(op for op in OP_REQUIRED_FIELDS if not _rows_for(op))
    assert not missing, f"exec ops missing from the MCP guide's op reference: {missing}"


def test_required_fields_are_marked_bold():
    problems = []
    for op, required in OP_REQUIRED_FIELDS.items():
        rows = " ".join(_rows_for(op))
        for field in required:
            if f"**{field}**" not in rows:
                problems.append(f"{op}.{field}")
    assert not problems, f"required fields not bold in their op row: {problems}"


def test_no_row_names_an_unknown_op():
    # Rows in the anchor/error tables start with an anchor id or code, not a
    # bare snake_case word, so only op-table rows contribute names here.
    known = set(OP_REQUIRED_FIELDS)
    named = {name for r in ROWS for name in _op_cell(r)}
    op_like = {n for n in named if "_" in n or n in known}
    # Error codes the tools explain (their table shares the row shape).
    non_ops = {*_tools._HINTS, "error", "internal_error", *_extra_codes()}
    unknown = sorted(op_like - known - non_ops)
    assert not unknown, f"guide rows name ops that don't exist: {unknown}"


def test_every_read_command_is_documented():
    commands = _read_commands()
    assert len(commands) > 30  # the regex still finds the dispatch table
    missing = sorted(c for c in commands if f"`{c}`" not in GUIDE)
    assert not missing, f"read commands missing from the MCP guide: {missing}"


def test_guide_read_serves_the_mcp_guide():
    out = _dispatch.dispatch_read(None, "guide", {})  # needs no worker
    assert out["guide"] == GUIDE
    assert "exec" in GUIDE and "rolled_back" in GUIDE
