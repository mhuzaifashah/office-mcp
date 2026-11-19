"""The MCP server: one process exposing the tools of every enabled Office app.

- `config.ServerConfig`: which apps to expose and the limits tools obey.
- `tool_definition`: how an app declares tools (`ToolDefinition`, `Toolset`).
- `response`: shared result builder (compact JSON, truncation, errors).
- `build.build_server`: assembles an `MCPServer` from the enabled toolsets.
- `cli.main`: the `office-mcp` command (flags + env vars, stdio transport).
"""

from __future__ import annotations

from .build import build_server
from .config import ServerConfig

__all__ = ["ServerConfig", "build_server"]
