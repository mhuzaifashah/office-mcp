"""Assemble the MCP server from the enabled apps' toolsets."""

from __future__ import annotations

import importlib
from collections.abc import AsyncIterator, Sequence
from contextlib import asynccontextmanager

from mcp.server import MCPServer

from .. import __version__
from .config import ServerConfig
from .tool_definition import ResourceDefinition, Toolset

__all__ = ["build_server", "load_toolset"]

# Where each app's toolset lives, imported only when that app is enabled so a
# Word-only server never imports the Excel or PowerPoint code.
_TOOLSETS: dict[str, str] = {
    "word": "office_mcp.word.tools:WordToolset",
    "excel": "office_mcp.excel.tools:ExcelToolset",
}

_INTRO = (
    "Office MCP drives the Microsoft Office apps running on this Windows machine. "
    "It works on the documents the user has open right now, and each editing call "
    "is a single undo step for the user. Tool names start with the app they act on."
)

_READ_ONLY_NOTE = (
    "This server is read-only: tools that change documents or write files are not "
    "available. Tell the user if they ask for an edit."
)


def load_toolset(app: str, config: ServerConfig) -> Toolset:
    """Import and construct the toolset for one app."""
    module_name, _, attr = _TOOLSETS[app].partition(":")
    factory = getattr(importlib.import_module(module_name), attr)
    return factory(config)


def _instructions(config: ServerConfig, toolsets: Sequence[Toolset]) -> str:
    parts = [_INTRO, *(t.instructions for t in toolsets if t.instructions)]
    if config.read_only:
        parts.append(_READ_ONLY_NOTE)
    return "\n\n".join(parts)


def _add_resource(server: MCPServer, res: ResourceDefinition) -> None:
    def read() -> str:
        return res.read()

    server.resource(
        res.uri,
        name=res.name,
        title=res.title,
        description=res.description,
        mime_type=res.mime_type,
    )(read)


def build_server(config: ServerConfig, *, toolsets: Sequence[Toolset] | None = None) -> MCPServer:
    """Build the server for `config`.

    `toolsets` overrides the per-app toolsets (tests use this); by default one
    is loaded for each app in `config.apps`. With `config.read_only`, only tools
    marked read-only are registered.
    """
    sets = (
        list(toolsets) if toolsets is not None else [load_toolset(a, config) for a in config.apps]
    )

    @asynccontextmanager
    async def lifespan(_server: MCPServer) -> AsyncIterator[None]:
        try:
            yield None
        finally:
            for toolset in sets:
                toolset.close()

    server = MCPServer(
        name="office-mcp",
        title="Office MCP",
        version=__version__,
        instructions=_instructions(config, sets),
        lifespan=lifespan,
    )
    for toolset in sets:
        for tool in toolset.tools():
            if config.read_only and not tool.read_only:
                continue
            server.add_tool(
                tool.handler,
                name=tool.name,
                title=tool.title,
                description=tool.description,
                annotations=tool.annotations(),
                structured_output=False,
            )
        for res in toolset.resources():
            _add_resource(server, res)
    return server
