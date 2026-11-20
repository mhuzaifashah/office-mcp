"""PowerPoint ops and tools: generated op table, checks before running, errors."""

from __future__ import annotations

import inspect
import json
import re
from types import SimpleNamespace

import anyio
import pytest
from mcp import Client

from office_mcp.comthread import InlineRunner
from office_mcp.ppt import dispatch, service
from office_mcp.ppt.controller import PowerPointController
from office_mcp.ppt.dispatch import OPS, run_batch, validate_ops
from office_mcp.ppt.errors import PowerPointNotRunningError, PptOpError
from office_mcp.ppt.guide import guide_text
from office_mcp.ppt.service import PowerPointService
from office_mcp.ppt.tools import READ_COMMANDS, PptToolset
from office_mcp.server import ServerConfig, build_server

# -- a fake deck -----------------------------------------------------------------


class _Collection:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]


def _deck(shapes_per_slide=((1, 2), (5,))):
    slides = [
        SimpleNamespace(Shapes=_Collection([SimpleNamespace(Id=i) for i in ids]))
        for ids in shapes_per_slide
    ]
    pres = SimpleNamespace(
        Slides=_Collection(slides),
        Name="Deck.pptx",
        FullName="C:\\decks\\Deck.pptx",
        Path="C:\\decks",
        Saved=True,
    )
    pres.Application = SimpleNamespace(Presentations=_Collection([pres]))
    return pres


@pytest.fixture
def calls(monkeypatch):
    """Record controller calls instead of touching PowerPoint."""
    seen: list[tuple[str, dict]] = []

    def recorder(name):
        def method(self, **kwargs):
            seen.append((name, kwargs))
            if kwargs.get("text") == "fail":
                raise RuntimeError("PowerPoint said no")
            return {"success": True, "shape_id": 99} if name.endswith("add") else {"success": True}

        return method

    for spec in OPS.values():
        monkeypatch.setattr(PowerPointController, spec.method, recorder(spec.method))
    return seen


# -- the op table ----------------------------------------------------------------


def test_every_op_maps_onto_a_controller_method():
    for name, spec in OPS.items():
        method = getattr(PowerPointController, spec.method)
        params = set(inspect.signature(method).parameters) - {"self"}
        assert set(spec.params.values()) == params, name


def test_ops_use_short_field_names():
    assert OPS["set_text"].required == ("slide", "shape", "text")
    assert OPS["group"].required == ("slide", "shapes")
    assert OPS["add_picture"].paths == ("path",)


def test_pictures_are_never_fetched_from_urls():
    assert all(spec.method != "picture_add_from_url" for spec in OPS.values())


def test_guide_documents_every_op_and_read():
    guide = guide_text()
    documented = set(re.findall(r"`([a-z_]+)`", guide))
    assert set(OPS) <= documented
    assert set(READ_COMMANDS) - {"guide"} <= documented
    assert set(READ_COMMANDS) == {"guide", *dispatch.READS}


# -- checks before anything runs -------------------------------------------------


@pytest.mark.parametrize(
    ("op", "message"),
    [
        ({"op": "nope"}, "unknown op"),
        ({"op": "set_text", "slide": 1, "shape": 1}, "needs text"),
        ({"op": "z_order", "slide": 1, "shape": 1, "cmd": "top"}, "cmd must be one of"),
        ({"op": "set_text", "slide": 9, "shape": 1, "text": "x"}, "slide 9 doesn't exist"),
        ({"op": "set_text", "slide": 1, "shape": 42, "text": "x"}, "shape id(s) [42]"),
        ({"op": "group", "slide": 1, "shapes": [1, 7]}, "shape id(s) [7]"),
    ],
)
def test_bad_ops_fail_before_anything_runs(calls, op, message):
    with pytest.raises(PptOpError) as info:
        run_batch(_deck(), [{"op": "set_notes", "slide": 1, "notes": "hi"}, op])
    assert message in str(info.value)
    assert info.value.details["applied"] == []
    assert calls == []


def test_later_slides_and_shapes_may_come_from_earlier_ops(calls):
    run_batch(
        _deck(),
        [
            {"op": "add_slide", "layout": "blank"},
            {"op": "set_notes", "slide": 3, "notes": "new slide"},
            {"op": "add_text_box", "slide": 1, "left": 0, "top": 0, "width": 10, "height": 10},
            {"op": "set_text", "slide": 1, "shape": 99, "text": "new box"},
        ],
    )
    assert [c[0] for c in calls] == [
        "slide_add",
        "slide_set_notes",
        "shape_add_text_box",
        "text_set",
    ]


def test_fields_are_passed_under_the_controller_names(calls):
    out = run_batch(_deck(), [{"op": "set_text", "slide": 1, "shape": 2, "text": "Hello"}])
    assert calls == [("text_set", {"slide_index": 1, "shape_id": 2, "text": "Hello"})]
    assert out == {"ok": True, "ops_run": 1}


def test_a_failure_reports_the_ops_already_applied(calls):
    with pytest.raises(PptOpError) as info:
        run_batch(
            _deck(),
            [
                {"op": "set_notes", "slide": 1, "notes": "a"},
                {"op": "set_text", "slide": 1, "shape": 1, "text": "ok"},
                {"op": "set_text", "slide": 1, "shape": 2, "text": "fail"},
            ],
        )
    d = info.value.details
    assert d["failure"]["index"] == 2
    assert d["applied"] == [0, 1]
    assert d["rolled_back"] is False


def test_picture_paths_are_screened_before_anything_runs(calls):
    svc = PowerPointService(runner=InlineRunner())
    with pytest.raises(Exception, match="network"):
        run_batch(
            _deck(),
            [
                {"op": "set_notes", "slide": 1, "notes": "a"},
                {
                    "op": "add_picture",
                    "slide": 1,
                    "path": "\\\\evil\\share\\x.png",
                    "left": 0,
                    "top": 0,
                },
            ],
            screen_path=svc.screen_picture,
        )
    assert calls == []


def test_image_dirs_limit_picture_sources(tmp_path):
    allowed = tmp_path / "img"
    allowed.mkdir()
    svc = PowerPointService(image_dirs=(str(allowed),), runner=InlineRunner())
    svc.screen_picture(str(allowed / "a.png"))
    with pytest.raises(Exception, match="outside the allowed image folders"):
        svc.screen_picture(str(tmp_path / "a.png"))


def test_ignored_fields_are_warned_about():
    assert validate_ops([{"op": "clear_text", "slide": 1, "shape": 1, "colour": "red"}]) == [
        {"index": 0, "op": "clear_text", "ignored_fields": ["colour"]}
    ]


def test_slide_ranges_parse():
    assert service._parse_slides(None, 3) == [1, 2, 3]
    assert service._parse_slides("2-3", 3) == [2, 3]
    with pytest.raises(PptOpError):
        service._parse_slides("4", 3)


# -- the tools -------------------------------------------------------------------


def _toolset(config=None):
    config = config or ServerConfig(apps=("ppt",))
    return PptToolset(config, service_factory=lambda: PowerPointService(runner=InlineRunner()))


def _call(name, args, config=None):
    config = config or ServerConfig(apps=("ppt",))

    async def go():
        async with Client(build_server(config, toolsets=[_toolset(config)])) as client:
            return await client.call_tool(name, args)

    return anyio.run(go)


def test_schemas_stay_portable_across_clients():
    async def go():
        async with Client(build_server(ServerConfig(apps=("ppt",)), toolsets=[_toolset()])) as c:
            return (await c.list_tools()).tools

    tools = anyio.run(go)
    assert {t.name for t in tools} == {"ppt_read", "ppt_exec", "ppt_snapshot", "ppt_write"}
    for tool in tools:
        schema = json.dumps(tool.input_schema, separators=(",", ":"))
        assert "anyOf" not in schema and "$ref" not in schema, tool.name
        assert len(schema) < 4500, (tool.name, len(schema))


def test_guide_read():
    result = _call("ppt_read", {"command": "guide"})
    assert result.content[0].text.startswith("# PowerPoint tools guide")


def test_not_running_suggests_opening_a_file(monkeypatch):
    def no_ppt(name=None):
        raise PowerPointNotRunningError("PowerPoint is not running")

    monkeypatch.setattr(service, "find_presentation", no_ppt)
    result = _call("ppt_read", {"command": "slides"})
    err = json.loads(result.content[0].text)["error"]
    assert err["code"] == "powerpoint_not_running"
    assert "ppt_write" in err["hint"]


def test_failed_batch_hint_says_not_to_resend_everything(monkeypatch, calls):
    monkeypatch.setattr(service, "find_presentation", lambda name=None: _deck())
    result = _call(
        "ppt_exec",
        {
            "ops": [
                {"op": "set_notes", "slide": 1, "notes": "a"},
                {"op": "set_text", "slide": 1, "shape": 1, "text": "fail"},
            ]
        },
    )
    err = json.loads(result.content[0].text)["error"]
    assert err["applied"] == [0]
    assert "not the whole batch" in err["hint"]


def test_save_as_is_refused_without_save_dirs(tmp_path):
    result = _call("ppt_write", {"command": "save_as", "path": str(tmp_path / "a.pptx")})
    assert json.loads(result.content[0].text)["error"]["code"] == "path_not_allowed"
