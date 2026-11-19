"""How an app declares its tools to the server.

Each app package (`office_mcp.word`, later Excel and PowerPoint) exposes a
toolset: an object that lists its `ToolDefinition`s, its share of the server
instructions, and any resources. The server registers whatever the enabled
toolsets declare, so adding an app never touches the server core.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from typing import Any, Protocol

from mcp_types import ToolAnnotations


@dataclass(frozen=True)
class ToolDefinition:
    """One MCP tool.

    `handler` is a plain function. Its typed, annotated parameters become the
    tool's input schema, and it returns a `CallToolResult` built with
    `office_mcp.server.response`. Handlers are synchronous: the SDK runs them on
    a worker thread, and each app serialises its COM calls onto its own thread.

    `read_only` decides whether the tool survives `--read-only`. The hints map
    to the MCP tool annotations clients use to decide when to ask the user.
    """

    name: str
    title: str
    description: str
    handler: Callable[..., Any]
    read_only: bool
    destructive: bool = False
    idempotent: bool = False

    def annotations(self) -> ToolAnnotations:
        if self.read_only:
            return ToolAnnotations(title=self.title, read_only_hint=True, open_world_hint=False)
        return ToolAnnotations(
            title=self.title,
            read_only_hint=False,
            destructive_hint=self.destructive,
            idempotent_hint=self.idempotent,
            open_world_hint=False,
        )


@dataclass(frozen=True)
class ResourceDefinition:
    """A static, read-only resource (for example an app's guide)."""

    uri: str
    name: str
    title: str
    description: str
    mime_type: str
    read: Callable[[], str]


class Toolset(Protocol):
    """What an app package hands the server.

    `instructions` is the app's paragraph of the server instructions every
    client shows the model at connect time. `close` releases the app's COM
    thread when the server stops; it must not close the Office app itself.
    """

    app: str
    instructions: str

    def tools(self) -> list[ToolDefinition]: ...

    def resources(self) -> list[ResourceDefinition]: ...

    def close(self) -> None: ...
