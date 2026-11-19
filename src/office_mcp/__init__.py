"""Office MCP: live Microsoft Office (Word, Excel, PowerPoint) for MCP clients."""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError
from importlib.metadata import version as _pkg_version

try:
    __version__ = _pkg_version("office-mcp")
except PackageNotFoundError:  # running from a source tree without an install
    __version__ = "0.0.0"

__all__ = ["__version__", "main"]


def main() -> None:
    """Entry point of the `office-mcp` command."""
    from .server.cli import main as _main  # noqa: PLC0415 — keep package import light

    _main()
