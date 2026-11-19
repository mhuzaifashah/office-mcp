"""The generated tool reference matches the tools the server registers."""

from __future__ import annotations

import runpy
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def test_tool_reference_is_up_to_date():
    gen = runpy.run_path(str(ROOT / "scripts" / "gen_tool_reference.py"))
    committed = (ROOT / "docs" / "tool-reference.md").read_text(encoding="utf-8")
    assert committed == gen["render"](), (
        "docs/tool-reference.md is out of date; run `uv run python scripts/gen_tool_reference.py`"
    )
