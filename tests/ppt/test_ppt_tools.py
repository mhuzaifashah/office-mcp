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
        SimpleNamespace(SlideID=256 + n, Shapes=_Collection([SimpleNamespace(Id=i) for i in ids]))
        for n, ids in enumerate(shapes_per_slide)
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


class _Tags:
    def __init__(self):
        self.values: dict[str, str] = {}
        self.log: list[tuple[str, str]] = []

    def Add(self, name, value):
        self.values[name] = value
        self.log.append(("add", name))

    def Delete(self, name):
        self.values.pop(name, None)
        self.log.append(("delete", name))

    def __call__(self, name):
        return self.values.get(name, "")


def _undoable_deck(*, windows=1, active=True, undo_works=True):
    """A deck whose PowerPoint supports undo entries; Undo drops the marker tag."""
    pres = _deck()
    pres.Tags = _Tags()
    pressed: list[str] = []

    def execute_mso(name):
        pressed.append(name)
        if undo_works:
            pres.Tags.values.clear()

    other = SimpleNamespace(FullName="C:\\decks\\Other.pptx")
    pres.Application = SimpleNamespace(
        Presentations=_Collection([pres]),
        StartNewUndoEntry=lambda: pressed.append("start"),
        CommandBars=SimpleNamespace(ExecuteMso=execute_mso),
        ActivePresentation=pres if active else other,
    )
    pres.Windows = _Collection([SimpleNamespace(Activate=lambda: None)] * windows)
    return pres, pressed


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


_FAILING_BATCH = [
    {"op": "set_notes", "slide": 1, "notes": "a"},
    {"op": "set_text", "slide": 1, "shape": 1, "text": "ok"},
    {"op": "set_text", "slide": 1, "shape": 2, "text": "fail"},
]


def test_a_successful_batch_is_one_undo_entry(calls):
    pres, pressed = _undoable_deck()
    out = run_batch(pres, [{"op": "set_text", "slide": 1, "shape": 2, "text": "Hello"}])
    assert out["ok"] is True
    assert pressed == ["start"]  # entry opened, Undo never pressed
    assert pres.Tags.log == [("add", "OFFICE_MCP_BATCH"), ("delete", "OFFICE_MCP_BATCH")]


def test_a_failed_batch_is_undone(calls):
    pres, pressed = _undoable_deck()
    with pytest.raises(PptOpError) as info:
        run_batch(pres, _FAILING_BATCH)
    d = info.value.details
    assert pressed == ["start", "Undo"]
    assert d["rolled_back"] is True
    assert d["applied"] == []
    assert d["ops_run"] == 2
    assert "rollback_error" not in d


@pytest.mark.parametrize(
    ("deck_kwargs", "reason"),
    [
        ({"windows": 0}, "no window"),
        ({"active": False}, "another presentation is active"),
        ({"undo_works": False}, "marker is still there"),
    ],
)
def test_when_undo_cant_run_the_applied_ops_are_reported(calls, deck_kwargs, reason):
    pres, _ = _undoable_deck(**deck_kwargs)
    with pytest.raises(PptOpError) as info:
        run_batch(pres, _FAILING_BATCH)
    d = info.value.details
    assert d["rolled_back"] is False
    assert d["applied"] == [0, 1]
    assert reason in d["rollback_error"]
    assert pres.Tags("OFFICE_MCP_BATCH") == ""  # marker cleaned up


def test_undo_that_leaves_a_different_layout_is_not_trusted(calls, monkeypatch):
    pres, _ = _undoable_deck()
    layouts = iter([[(256, 2), (257, 1)], [(256, 3), (257, 1)]])
    monkeypatch.setattr(dispatch, "_layout", lambda p: next(layouts))
    with pytest.raises(PptOpError) as info:
        run_batch(pres, _FAILING_BATCH)
    assert info.value.details["rolled_back"] is False
    assert "don't match" in info.value.details["rollback_error"]


def test_rolled_back_batch_hint_says_to_resend_it(monkeypatch, calls):
    pres, _ = _undoable_deck()
    monkeypatch.setattr(service, "find_presentation", lambda name=None: pres)
    result = _call("ppt_exec", {"ops": _FAILING_BATCH})
    err = json.loads(result.content[0].text)["error"]
    assert err["rolled_back"] is True
    assert "send the batch again" in err["hint"]


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
