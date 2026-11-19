"""The Word MCP tools end to end (in-process client, mocked Word)."""

from __future__ import annotations

import inspect
import json
import re

import anyio
import pytest
from mcp import Client

from office_mcp.server import ServerConfig, build_server
from office_mcp.word.controller import WordController
from office_mcp.word.tools import READ_COMMANDS, WordToolset
from office_mcp.word.wordcom.runtime import _dispatch
from office_mcp.word.wordcom.runtime._worker import InlineWorker


def _toolset(config: ServerConfig | None = None) -> WordToolset:
    config = config or ServerConfig()
    return WordToolset(
        config,
        controller_factory=lambda: WordController(
            save_dirs=list(config.save_dirs),
            image_dirs=list(config.image_dirs),
            worker=InlineWorker(),
        ),
    )


def _call(name: str, args: dict, *, toolset: WordToolset | None = None, config=None):
    config = config or ServerConfig()
    server = build_server(config, toolsets=[toolset or _toolset(config)])

    async def go():
        async with Client(server) as client:
            return await client.call_tool(name, args)

    return anyio.run(go)


def _json(result):
    return json.loads(result.content[0].text)


def _error(result):
    assert result.is_error, result.content
    return _json(result)["error"]


def test_read_commands_match_the_dispatcher():
    src = inspect.getsource(_dispatch.dispatch_read)
    assert set(READ_COMMANDS) == set(re.findall(r'command == "([a-z_]+)"', src))


def test_schemas_stay_portable_across_clients():
    async def go():
        async with Client(build_server(ServerConfig(), toolsets=[_toolset()])) as client:
            return (await client.list_tools()).tools

    for tool in anyio.run(go):
        schema = json.dumps(tool.input_schema, separators=(",", ":"))
        assert "anyOf" not in schema and "$ref" not in schema, tool.name
        # Codex drops parameter descriptions from schemas over 5,000 bytes.
        assert len(schema) < 4500, (tool.name, len(schema))


def test_optional_params_accept_null(fake_word):
    result = _call("word_read", {"command": "status", "doc": None, "anchor_id": None})
    assert not result.is_error


def test_status_lists_open_documents(fake_word):
    result = _call("word_read", {"command": "status"})
    assert not result.is_error
    docs = _json(result)
    assert isinstance(docs, list) and docs and "name" in docs[0]


def test_outline_returns_heading_anchor_ids(fake_word):
    rows = _json(_call("word_read", {"command": "outline"}))
    assert any(str(r.get("anchor_id", "")).startswith("heading:") for r in rows)


def test_guide_is_plain_markdown():
    result = _call("word_read", {"command": "guide"})
    assert result.content[0].text.startswith("# Word (MCP)")


def test_read_image_returns_an_image_block(fake_word):
    result = _call("word_read", {"command": "read_image", "anchor_id": "image:1"})
    meta, image = result.content
    assert "base64" not in json.loads(meta.text)
    assert image.type == "image" and image.mime_type == "image/png"


def test_exec_applies_a_batch(fake_word):
    result = _call(
        "word_exec",
        {"ops": [{"op": "write_bookmark", "name": "Address", "text": "1 Main St"}], "label": "t"},
    )
    body = _json(result)
    assert body["ok"] is True and body["ops_run"] == 1


def test_exec_failure_reports_the_failing_op_and_rollback(fake_word):
    result = _call(
        "word_exec",
        {
            "ops": [
                {"op": "write_bookmark", "name": "Address", "text": "x"},
                {"op": "no_such_op"},
            ]
        },
    )
    error = _error(result)
    assert error["code"] == "error"
    assert error["failure"]["index"] == 1
    assert error["ops_run"] == 1
    assert "rolled_back" in error
    assert "guide" in error["hint"]


def test_exec_needs_at_least_one_op(fake_word):
    assert _call("word_exec", {"ops": []}).is_error


def test_stale_anchor_gets_a_reread_hint(fake_word):
    error = _error(_call("word_read", {"command": "read_text", "anchor_id": "heading:99"}))
    assert error["code"] == "anchor_not_found"
    assert "Re-read" in error["hint"]


def test_word_not_running_suggests_opening_a_file(no_word):
    error = _error(_call("word_read", {"command": "status"}))
    assert error["code"] == "word_not_running"
    assert 'word_write(command="open"' in error["hint"]


def test_save_as_is_denied_without_save_dirs(fake_word, tmp_path):
    error = _error(_call("word_write", {"command": "save_as", "path": str(tmp_path / "out.docx")}))
    assert error["code"] == "path_not_allowed"
    assert "--save-dir" in error["hint"]


def test_open_requires_a_path(fake_word):
    error = _error(_call("word_write", {"command": "open"}))
    assert "requires 'path'" in error["message"]


def test_open_missing_file(fake_word, tmp_path):
    error = _error(_call("word_write", {"command": "open", "path": str(tmp_path / "x.docx")}))
    assert error["code"] == "file_not_found"


def test_track_toggle(fake_word):
    body = _json(_call("word_write", {"command": "track", "on": True}))
    assert body == {"ok": True, "command": "track", "track_changes": True}


def test_large_read_is_truncated_with_a_hint(fake_word):
    config = ServerConfig(max_output_chars=1000)
    toolset = _toolset(config)
    rows = [{"anchor_id": f"para:{i}", "text": "y" * 40} for i in range(1, 200)]
    toolset.controller.read = lambda command, **p: rows  # type: ignore[method-assign]
    result = _call("word_read", {"command": "paragraphs"}, toolset=toolset, config=config)
    body, note = result.content
    assert len(json.loads(body.text)) < len(rows)
    assert "start and count" in note.text


def test_snapshot_caps_pages_and_says_how_to_get_the_rest(fake_word):
    config = ServerConfig(snapshot_max_pages=2)
    toolset = _toolset(config)
    seen = {}

    def fake_snapshot(**kwargs):
        seen.update(kwargs)
        return [(1, b"one"), (2, b"two"), (3, b"three"), (4, b"four")]

    toolset.controller.snapshot = fake_snapshot  # type: ignore[method-assign]
    result = _call("word_snapshot", {}, toolset=toolset, config=config)
    kinds = [c.type for c in result.content]
    assert kinds == ["text", "image", "text", "image", "text"]
    assert result.content[0].text == "Page 1"
    assert 'pages="3-4"' in result.content[-1].text
    assert seen["max_dim"] == config.snapshot_max_dim


def test_snapshot_rejects_pages_and_anchor_together(fake_word):
    error = _error(_call("word_snapshot", {"pages": "1", "anchor": "heading:1"}))
    assert "not both" in error["message"]


@pytest.mark.parametrize("read_only", [False, True])
def test_read_only_hides_edit_tools(read_only):
    config = ServerConfig(read_only=read_only)
    server = build_server(config, toolsets=[_toolset(config)])

    async def go():
        async with Client(server) as client:
            return sorted(t.name for t in (await client.list_tools()).tools)

    names = anyio.run(go)
    if read_only:
        assert names == ["word_read", "word_snapshot"]
    else:
        assert names == ["word_exec", "word_read", "word_snapshot", "word_write"]


def test_controller_is_created_lazily_and_closed_on_shutdown(fake_word):
    toolset = _toolset()
    assert toolset._controller is None
    _call("word_read", {"command": "status"}, toolset=toolset)
    assert toolset._controller is None  # closed by the server lifespan


def test_every_error_code_with_a_hint_is_in_the_guide():
    from office_mcp.word import tools
    from office_mcp.word.wordcom._guide import skill_body

    guide = skill_body("mcp")
    for code in [*tools._HINTS, "error", "internal_error"]:
        assert f"| `{code}` |" in guide, code


def test_read_only_hint_does_not_mention_missing_tools(no_word):
    config = ServerConfig(read_only=True)
    error = _error(_call("word_read", {"command": "status"}, config=config))
    assert error["code"] == "word_not_running"
    assert "word_write" not in error["hint"]
