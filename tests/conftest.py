"""Root test config.

Importing ``office_mcp.word`` runs its sys.path bootstrap so the engine is
importable as top-level ``wordcom`` (needed by importlib.resources and tests).

Tests that exercise the removed command-line frontend (``wordcom.cli``) are
skipped at collection time rather than deleted, so the suite stays a faithful
behavior baseline for the engine.
"""

from __future__ import annotations

import inspect
import sys
import types

import pytest

import office_mcp.word  # noqa: F401  (side effect: wordcom on sys.path)


def _install_cli_stub() -> None:
    """Register a stand-in ``wordcom.cli.main`` so test modules import cleanly.

    ``EXIT_*`` names resolve to distinct ints; any other attribute is a callable
    that skips the calling test.
    """
    pkg = types.ModuleType("wordcom.cli")
    mod = types.ModuleType("wordcom.cli.main")
    codes: dict[str, int] = {}

    def _getattr(name: str):
        if name.startswith("EXIT_"):
            return codes.setdefault(name, 100 + len(codes))
        if name.startswith("__"):
            raise AttributeError(name)

        def _skip(*_a, **_k):
            pytest.skip("CLI frontend not shipped")

        return _skip

    mod.__getattr__ = _getattr  # type: ignore[attr-defined]
    pkg.main = mod  # type: ignore[attr-defined]
    sys.modules.setdefault("wordcom.cli", pkg)
    sys.modules.setdefault("wordcom.cli.main", mod)

    # CLI tests drive the frontend through click's CliRunner; stub it the same way.
    click = types.ModuleType("click")
    testing = types.ModuleType("click.testing")

    class CliRunner:
        def __init__(self, *_a, **_k) -> None:
            pass

        def invoke(self, *_a, **_k):
            pytest.skip("CLI frontend not shipped")

    testing.CliRunner = CliRunner  # type: ignore[attr-defined]
    click.testing = testing  # type: ignore[attr-defined]
    sys.modules.setdefault("click", click)
    sys.modules.setdefault("click.testing", testing)


_install_cli_stub()


def pytest_collection_modifyitems(config, items):
    skip_cli = pytest.mark.skip(reason="CLI frontend not shipped")
    for item in items:
        fn = getattr(item, "function", None)
        if fn is None:
            continue
        try:
            src = inspect.getsource(fn)
        except (OSError, TypeError):
            continue
        if "wordcom.cli" in src:
            item.add_marker(skip_cli)
