"""The Excel MCP tools: schemas, guide, error codes and hints, file rules."""

from __future__ import annotations

import json
import re

import anyio
import pytest
from mcp import Client

from office_mcp.comthread import InlineRunner
from office_mcp.excel import dispatch, service
from office_mcp.excel.errors import ExcelNotRunningError
from office_mcp.excel.service import ExcelService
from office_mcp.excel.tools import READ_COMMANDS, ExcelToolset, guide_text
from office_mcp.server import ServerConfig, build_server


def _toolset(config: ServerConfig | None = None) -> ExcelToolset:
    config = config or ServerConfig(apps=("excel",))
    return ExcelToolset(
        config,
        service_factory=lambda: ExcelService(save_dirs=config.save_dirs, runner=InlineRunner()),
    )


def _call(name: str, args: dict, *, config: ServerConfig | None = None):
    config = config or ServerConfig(apps=("excel",))
    server = build_server(config, toolsets=[_toolset(config)])

    async def go():
        async with Client(server) as client:
            return await client.call_tool(name, args)

    return anyio.run(go)


def _error(result):
    assert result.is_error, result.content
    return json.loads(result.content[0].text)["error"]


@pytest.fixture
def with_book(monkeypatch, book):
    monkeypatch.setattr(service, "find_book", lambda name=None: book)
    return book


def test_read_commands_match_the_dispatcher():
    assert set(READ_COMMANDS) == {"guide", *dispatch.READS}


def test_guide_documents_every_op_and_read():
    guide = guide_text()
    documented = set(re.findall(r"`([a-z_0-9]+)`", guide))
    assert set(dispatch.OPS) <= documented, set(dispatch.OPS) - documented
    assert set(dispatch.READS) <= documented, set(dispatch.READS) - documented


def test_schemas_stay_portable_across_clients():
    async def go():
        server = build_server(ServerConfig(apps=("excel",)), toolsets=[_toolset()])
        async with Client(server) as client:
            return (await client.list_tools()).tools

    tools = anyio.run(go)
    assert {t.name for t in tools} == {"excel_read", "excel_exec", "excel_snapshot", "excel_write"}
    for tool in tools:
        schema = json.dumps(tool.input_schema, separators=(",", ":"))
        assert "anyOf" not in schema and "$ref" not in schema, tool.name
        assert len(schema) < 4500, (tool.name, len(schema))


def test_guide_read_is_plain_markdown():
    result = _call("excel_read", {"command": "guide"})
    assert not result.is_error
    assert result.content[0].text.startswith("# Excel tools guide")


def test_excel_not_running_suggests_opening_a_file(monkeypatch):
    def no_excel(name=None):
        raise ExcelNotRunningError("Excel is not running")

    monkeypatch.setattr(service, "find_book", no_excel)
    err = _error(_call("excel_read", {"command": "sheets"}))
    assert err["code"] == "excel_not_running"
    assert "excel_write" in err["hint"]


def test_unknown_sheet_points_at_the_sheets_read(with_book):
    err = _error(_call("excel_read", {"command": "used_range", "sheet": "Missing"}))
    assert err["code"] == "not_found"
    assert 'command="sheets"' in err["hint"]


def test_a_bad_op_is_reported_with_its_index(with_book):
    err = _error(_call("excel_exec", {"ops": [{"op": "merge"}]}))
    assert err["code"] == "op_error"
    assert err["failure"]["index"] == 0
    assert err["ops_run"] == 0


def test_save_as_is_refused_without_save_dirs(with_book, tmp_path):
    err = _error(_call("excel_write", {"command": "save_as", "path": str(tmp_path / "out.xlsx")}))
    assert err["code"] == "path_not_allowed"


def test_save_as_checks_the_extension(with_book, tmp_path):
    config = ServerConfig(apps=("excel",), save_dirs=(str(tmp_path),))
    err = _error(
        _call(
            "excel_write", {"command": "save_as", "path": str(tmp_path / "out.txt")}, config=config
        )
    )
    assert err["code"] == "op_error"


def test_open_refuses_network_paths():
    err = _error(_call("excel_write", {"command": "open", "path": "\\\\server\\share\\a.xlsx"}))
    assert err["code"] == "path_not_allowed"


def test_open_missing_file(tmp_path):
    err = _error(_call("excel_write", {"command": "open", "path": str(tmp_path / "none.xlsx")}))
    assert err["code"] == "file_not_found"


def test_save_refuses_a_never_saved_workbook(with_book):
    with_book.fullname = "Book1"
    err = _error(_call("excel_write", {"command": "save"}))
    assert "save_as" in err["message"]


def test_read_only_hides_edit_tools():
    async def go():
        config = ServerConfig(apps=("excel",), read_only=True)
        async with Client(build_server(config, toolsets=[_toolset(config)])) as client:
            return sorted(t.name for t in (await client.list_tools()).tools)

    assert anyio.run(go) == ["excel_read", "excel_snapshot"]


def test_excel_is_opt_in():
    async def go(config):
        async with Client(build_server(config)) as client:
            return sorted(t.name for t in (await client.list_tools()).tools)

    assert not any(n.startswith("excel_") for n in anyio.run(go, ServerConfig()))
    both = anyio.run(go, ServerConfig(apps=("word", "excel")))
    assert {"excel_read", "word_read"} <= set(both)
