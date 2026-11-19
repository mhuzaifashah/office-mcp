"""Server assembly: registration, read-only filtering, annotations, resources, shutdown."""

from __future__ import annotations

import anyio
from mcp import Client
from mcp_types import CallToolResult, TextContent

from office_mcp.server import ServerConfig, build_server
from office_mcp.server.tool_definition import ResourceDefinition, ToolDefinition


class _StubToolset:
    app = "word"
    instructions = "Stub instructions."

    def __init__(self) -> None:
        self.closed = False

    def tools(self):
        def stub_read(x: int = 1) -> CallToolResult:
            return CallToolResult(content=[TextContent(text=f"read {x}")])

        def stub_edit() -> CallToolResult:
            return CallToolResult(content=[TextContent(text="edited")])

        return [
            ToolDefinition("stub_read", "Read", "Reads.", stub_read, read_only=True),
            ToolDefinition(
                "stub_edit", "Edit", "Edits.", stub_edit, read_only=False, destructive=True
            ),
        ]

    def resources(self):
        return [
            ResourceDefinition(
                uri="office-mcp://stub/guide",
                name="stub-guide",
                title="Stub guide",
                description="A guide.",
                mime_type="text/markdown",
                read=lambda: "# Guide",
            )
        ]

    def close(self) -> None:
        self.closed = True


def _run(server, fn):
    async def go():
        async with Client(server) as client:
            return await fn(client)

    return anyio.run(go)


def test_tools_annotations_and_instructions():
    toolset = _StubToolset()
    server = build_server(ServerConfig(), toolsets=[toolset])

    async def check(client):
        tools = {t.name: t for t in (await client.list_tools()).tools}
        return tools, client.instructions

    tools, instructions = _run(server, check)
    assert set(tools) == {"stub_read", "stub_edit"}
    read, edit = tools["stub_read"].annotations, tools["stub_edit"].annotations
    assert read.read_only_hint is True and read.open_world_hint is False
    assert edit.read_only_hint is False and edit.destructive_hint is True
    assert edit.idempotent_hint is False
    assert "Stub instructions." in instructions
    assert "read-only" not in instructions


def test_read_only_registers_only_read_only_tools():
    server = build_server(ServerConfig(read_only=True), toolsets=[_StubToolset()])

    async def check(client):
        return [t.name for t in (await client.list_tools()).tools], client.instructions

    names, instructions = _run(server, check)
    assert names == ["stub_read"]
    assert "This server is read-only" in instructions


def test_resource_is_readable():
    server = build_server(ServerConfig(), toolsets=[_StubToolset()])

    async def check(client):
        listed = [str(r.uri) for r in (await client.list_resources()).resources]
        body = await client.read_resource("office-mcp://stub/guide")
        return listed, body.contents[0].text

    listed, text = _run(server, check)
    assert listed == ["office-mcp://stub/guide"]
    assert text == "# Guide"


def test_toolsets_are_closed_when_the_server_stops():
    toolset = _StubToolset()
    server = build_server(ServerConfig(), toolsets=[toolset])

    async def call(client):
        return await client.call_tool("stub_read", {"x": 2})

    result = _run(server, call)
    assert result.content[0].text == "read 2"
    assert toolset.closed


def test_default_build_loads_the_word_toolset():
    server = build_server(ServerConfig())

    async def check(client):
        return sorted(t.name for t in (await client.list_tools()).tools)

    assert _run(server, check) == ["word_exec", "word_read", "word_snapshot", "word_write"]
