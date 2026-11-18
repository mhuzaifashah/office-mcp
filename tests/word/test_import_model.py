"""The engine exists once, as office_mcp.word.wordcom (W22).

The old sys.path bootstrap made it importable twice (as top-level `wordcom`
and as `office_mcp.word.wordcom`), giving two copies of every exception class,
so an error raised through one path was not caught by an `except` written
against the other.
"""

from __future__ import annotations

import importlib.util
import sys

from office_mcp.word import controller
from office_mcp.word.wordcom import exceptions


def test_no_top_level_wordcom_module():
    import office_mcp.word  # noqa: F401  (must not put wordcom on sys.path)

    assert importlib.util.find_spec("wordcom") is None
    assert "wordcom" not in sys.modules


def test_controller_and_engine_share_exception_classes():
    assert controller.WordNotRunningError is exceptions.WordNotRunningError


def test_bundled_guide_loads_from_the_package():
    from office_mcp.word.wordcom._guide import skill_body

    assert skill_body("mcp").startswith("# Word (MCP)")


def test_version_comes_from_the_distribution():
    from importlib.metadata import version

    from office_mcp.word import wordcom

    assert wordcom.__version__ == version("office-mcp")
