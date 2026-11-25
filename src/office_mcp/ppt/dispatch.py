"""The PowerPoint read commands and edit ops behind the MCP tools.

Each op maps onto one `PowerPointController` method. Its required and optional
fields come from that method's signature, with the parameter names agents see
shortened (`slide_index` -> `slide`, `shape_id` -> `shape`, ...), so the op
table can't drift from the controller.

A batch:

1. checks every op before any runs: name, fields, allowed values, and that the
   slides and shapes it names exist (unless an earlier op in the batch adds
   slides or shapes);
2. runs the ops as one PowerPoint undo entry (`StartNewUndoEntry`), opened by a
   marker tag on the presentation, so one Undo reverses the whole batch;
3. on a failure part-way, presses Undo once (`ExecuteMso("Undo")`, the ribbon
   button: PowerPoint has no Undo method) and checks the result: the marker
   must be gone and the slides and shapes must match the state before the
   batch. If the presentation has no window, or the check fails, it reports
   which ops stayed applied instead.
"""

from __future__ import annotations

import inspect
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

from . import constants as C
from .controller import PowerPointController
from .errors import PowerPointError, PptOpError
from .session import controller_for, presentations

__all__ = ["OPS", "READS", "dispatch_read", "run_batch", "validate_ops"]

# Controller parameter -> the field name agents use.
_FIELD_NAMES: dict[str, str] = {
    "slide_index": "slide",
    "shape_id": "shape",
    "shape_ids": "shapes",
    "file_path": "path",
    "picture_path": "picture",
    "connector_id": "connector",
    "begin_shape_id": "begin_shape",
    "end_shape_id": "end_shape",
    "source_id": "source",
    "target_id": "target",
}


def com_message(exc: BaseException) -> str:
    """PowerPoint's own description of a COM error, else the exception text."""
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3 and args[2][2]:
        return str(args[2][2]).strip()
    if len(args) >= 2 and isinstance(args[1], str):
        return args[1]
    return str(exc)


def _is_com_error(exc: BaseException) -> bool:
    return type(exc).__name__ == "com_error"


def _strip(result: Any) -> Any:
    if isinstance(result, dict):
        return {k: v for k, v in result.items() if k != "success"} or None
    return result


# -- edit ops --------------------------------------------------------------------


@dataclass(frozen=True)
class OpSpec:
    method: str
    required: tuple[str, ...]
    optional: tuple[str, ...]
    params: dict[str, str]  # field -> controller parameter
    choices: dict[str, tuple[str, ...]] = field(default_factory=dict)
    paths: tuple[str, ...] = ()  # fields naming a local file to screen
    creates_slides: bool = False
    creates_shapes: bool = False


def _op(
    method: str,
    *,
    choices: dict[str, Any] | None = None,
    paths: tuple[str, ...] = (),
    creates_slides: bool = False,
    creates_shapes: bool = False,
) -> OpSpec:
    sig = inspect.signature(getattr(PowerPointController, method))
    required: list[str] = []
    optional: list[str] = []
    params: dict[str, str] = {}
    for p in sig.parameters.values():
        if p.name == "self":
            continue
        name = _FIELD_NAMES.get(p.name, p.name)
        params[name] = p.name
        (required if p.default is inspect.Parameter.empty else optional).append(name)
    return OpSpec(
        method,
        tuple(required),
        tuple(optional),
        params,
        {k: tuple(v) for k, v in (choices or {}).items()},
        paths,
        creates_slides,
        creates_shapes,
    )


OPS: dict[str, OpSpec] = {
    # slides
    "add_slide": _op("slide_add", creates_slides=True, creates_shapes=True),
    "delete_slide": _op("slide_delete"),
    "duplicate_slide": _op("slide_duplicate", creates_slides=True, creates_shapes=True),
    "move_slide": _op("slide_move"),
    "set_layout": _op("slide_set_layout", creates_shapes=True),
    "hide_slide": _op("slide_set_hidden"),
    "set_notes": _op("slide_set_notes"),
    "go_to_slide": _op("slide_goto"),
    "set_background": _op("slide_set_background", paths=("picture",)),
    "set_footer": _op("slide_set_footer"),
    # text
    "set_placeholder": _op("placeholder_set_text", choices={"placeholder": C.PLACEHOLDER_NAME_MAP}),
    "set_text": _op("text_set"),
    "append_text": _op("text_append"),
    "clear_text": _op("text_clear"),
    "format_text": _op("text_set_font"),
    "format_paragraph": _op(
        "text_set_paragraph",
        choices={"alignment": C.PARAGRAPH_ALIGN_MAP, "bullet": C.BULLET_TYPE_MAP},
    ),
    # shapes
    "add_shape": _op(
        "shape_add", choices={"shape_type": C.AUTOSHAPE_NAME_MAP}, creates_shapes=True
    ),
    "add_text_box": _op(
        "shape_add_text_box",
        choices={"orientation": C.TEXT_ORIENTATION_MAP},
        creates_shapes=True,
    ),
    "delete_shape": _op("shape_delete"),
    "move_shape": _op("shape_move"),
    "resize_shape": _op("shape_resize"),
    "fill": _op("shape_set_fill_color"),
    "fill_transparency": _op("shape_set_fill_transparency"),
    "remove_fill": _op("shape_remove_fill"),
    "line_color": _op("shape_set_line_color"),
    "line_weight": _op("shape_set_line_weight"),
    "line_dash": _op("shape_set_line_dash_style", choices={"dash_style": C.DASH_STYLE_MAP}),
    "rotate": _op("shape_set_rotation"),
    "z_order": _op("shape_z_order", choices={"cmd": C.Z_ORDER_MAP}),
    "group": _op("shape_group", creates_shapes=True),
    "ungroup": _op("shape_ungroup", creates_shapes=True),
    "align": _op("shape_align", choices={"alignment": C.ALIGN_MAP}),
    "distribute": _op("shape_distribute", choices={"direction": C.DISTRIBUTE_MAP}),
    "copy_format": _op("shape_copy_format"),
    "shadow": _op("shape_set_shadow"),
    # pictures
    "add_picture": _op("picture_add_from_file", paths=("path",), creates_shapes=True),
    "replace_picture": _op("picture_replace", paths=("path",), creates_shapes=True),
    "crop_picture": _op("picture_crop"),
    "picture_border": _op("picture_set_border", choices={"dash_style": C.DASH_STYLE_MAP}),
    # tables
    "add_table": _op("table_add", creates_shapes=True),
    "set_cell": _op("table_set_cell_text"),
    "cell_fill": _op("table_set_cell_fill"),
    "cell_font": _op("table_set_cell_font"),
    "cell_align": _op(
        "table_set_cell_alignment",
        choices={"horizontal": C.PARAGRAPH_ALIGN_MAP, "vertical": C.VERTICAL_ALIGN_MAP},
    ),
    "column_width": _op("table_set_column_width"),
    "row_height": _op("table_set_row_height"),
    "merge_cells": _op("table_merge_cells"),
    "table_add_row": _op("table_add_row"),
    "table_add_column": _op("table_add_column"),
    "table_style": _op("table_apply_style", choices={"style": C.TABLE_STYLE_MAP}),
    # charts
    "add_chart": _op("chart_add", choices={"chart_type": C.CHART_TYPE_MAP}, creates_shapes=True),
    "chart_data": _op("chart_set_data"),
    "chart_title": _op("chart_set_title"),
    "chart_type": _op("chart_set_type", choices={"chart_type": C.CHART_TYPE_MAP}),
    "series_color": _op("chart_set_series_color"),
    "data_labels": _op("chart_set_data_labels"),
    "chart_axis": _op("chart_set_axis", choices={"axis": C.AXIS_TYPE_MAP}),
    "chart_legend": _op("chart_set_legend", choices={"position": C.LEGEND_POSITION_MAP}),
    "gridlines": _op("chart_set_gridlines", choices={"axis": C.AXIS_TYPE_MAP}),
    "chart_style": _op("chart_set_style"),
    # connectors and links
    "add_connector": _op(
        "connector_add", choices={"connector_type": C.CONNECTOR_TYPE_MAP}, creates_shapes=True
    ),
    "connect": _op("connector_connect"),
    "disconnect": _op("connector_disconnect", choices={"end": ("begin", "end", "both")}),
    "reroute": _op("connector_reroute"),
    "set_link": _op("hyperlink_set"),
    "remove_link": _op("hyperlink_remove"),
}


def validate_ops(ops: list[Any]) -> list[dict[str, Any]]:
    """Check names, fields and allowed values. Returns warnings for ignored fields."""
    if not isinstance(ops, list) or not ops:
        raise PptOpError("ops must be a non-empty list")
    warnings: list[dict[str, Any]] = []
    for i, op in enumerate(ops):
        if not isinstance(op, dict):
            raise _bad(i, op, "each op must be an object")
        name = op.get("op")
        spec = OPS.get(str(name))
        if spec is None:
            raise _bad(i, op, f"unknown op {name!r}")
        missing = [f for f in spec.required if op.get(f) is None]
        if missing:
            raise _bad(i, op, f"op {name!r} needs {', '.join(missing)}")
        for key, allowed in spec.choices.items():
            value = op.get(key)
            if isinstance(value, str) and value.lower() not in allowed:
                raise _bad(i, op, f"{key} must be one of: {', '.join(allowed)}")
        ignored = sorted(set(op) - {"op", *spec.params})
        if ignored:
            warnings.append({"index": i, "op": name, "ignored_fields": ignored})
    return warnings


def _bad(
    index: int,
    op: Any,
    message: str,
    *,
    applied: list[int] | None = None,
    rolled_back: bool = False,
    rollback_error: str | None = None,
) -> PptOpError:
    details: dict[str, Any] = {
        "failure": {"index": index, "op": op, "error": message},
        "ops_run": len(applied or []),
        "rolled_back": rolled_back,
        "applied": [] if rolled_back else (applied or []),
    }
    if rollback_error:
        details["rollback_error"] = rollback_error
    return PptOpError(f"op {index}: {message}", **details)


# -- rollback through PowerPoint's own undo --------------------------------------

_MARKER = "OFFICE_MCP_BATCH"


def _layout(pres: Any) -> list[tuple[int, int]]:
    """(SlideID, shape count) per slide: cheap to read, and any added, deleted,
    moved or re-shaped slide changes it."""
    slides = pres.Slides
    return [
        (int(slides(i).SlideID), int(slides(i).Shapes.Count))
        for i in range(1, int(slides.Count) + 1)
    ]


def _begin_undo_entry(pres: Any) -> str | None:
    """Start the batch's undo entry with a marker tag; None if PowerPoint can't."""
    token = f"{id(pres):x}-{time.monotonic_ns():x}"
    try:
        pres.Application.StartNewUndoEntry()
        pres.Tags.Add(_MARKER, token)
    except Exception:  # noqa: BLE001 - PowerPoint 2007, or a protected presentation
        return None
    return token


def _end_undo_entry(pres: Any) -> None:
    # Joins the batch's undo entry, so one Ctrl+Z still reverses the whole batch.
    try:
        pres.Tags.Delete(_MARKER)
    except Exception:  # noqa: BLE001
        pass


def _undo_batch(pres: Any, token: str | None, before: list[tuple[int, int]]) -> str | None:
    """Undo the batch's entry. None on success, else why the batch stays applied."""
    if token is None:
        return "this PowerPoint can't group automation into one undo entry"
    app = pres.Application
    try:
        if int(pres.Windows.Count) < 1:
            return "the presentation has no window, so PowerPoint's Undo can't reach it"
        pres.Windows(1).Activate()
        if str(app.ActivePresentation.FullName) != str(pres.FullName):
            return "another presentation is active, so Undo was not pressed"
        if str(pres.Tags(_MARKER)) != token:
            return "the batch's undo entry was not found"
        app.CommandBars.ExecuteMso("Undo")
    except Exception as exc:  # noqa: BLE001
        return f"Undo failed: {com_message(exc) if _is_com_error(exc) else exc}"
    try:
        marker = str(pres.Tags(_MARKER))
        layout = _layout(pres)
    except Exception as exc:  # noqa: BLE001
        return f"couldn't check the presentation after Undo: {exc}"
    if marker:
        return "Undo didn't reverse the batch (its marker is still there)"
    if layout != before:
        return "after Undo the slides and shapes don't match the state before the batch"
    return None


def _shape_ids(slide: Any) -> set[int]:
    ids: set[int] = set()
    shapes = slide.Shapes
    for i in range(1, int(shapes.Count) + 1):
        try:
            ids.add(int(shapes(i).Id))
        except Exception:  # noqa: BLE001
            continue
    return ids


def _precheck(pres: Any, ops: list[dict[str, Any]]) -> None:
    """Slides and shapes an op names must exist, unless an earlier op may create them."""
    count = int(pres.Slides.Count)
    slides_may_change = False
    shapes_may_change = False
    shape_cache: dict[int, set[int]] = {}
    for i, op in enumerate(ops):
        spec = OPS[op["op"]]
        slide = op.get("slide")
        if isinstance(slide, int) and not slides_may_change:
            if not 1 <= slide <= count:
                raise _bad(i, op, f"slide {slide} doesn't exist (the deck has {count})")
            if not shapes_may_change:
                if slide not in shape_cache:
                    shape_cache[slide] = _shape_ids(pres.Slides(slide))
                ids = shape_cache[slide]
                wanted = [op.get(f) for f in ("shape", "connector", "source", "target")]
                wanted += list(op.get("shapes") or [])
                missing = [s for s in wanted if isinstance(s, int) and s not in ids]
                if missing:
                    raise _bad(
                        i,
                        op,
                        f"shape id(s) {missing} not on slide {slide}; "
                        "read the slide for current ids",
                    )
        slides_may_change = (
            slides_may_change
            or spec.creates_slides
            or op["op"]
            in (
                "delete_slide",
                "move_slide",
            )
        )
        shapes_may_change = shapes_may_change or spec.creates_shapes


def run_batch(
    pres: Any,
    ops: list[dict[str, Any]],
    *,
    screen_path: Callable[[str], None] | None = None,
) -> dict[str, Any]:
    """Apply `ops` in order; stop at the first failure and report what applied."""
    warnings = validate_ops(ops)
    for i, op in enumerate(ops):
        for key in OPS[op["op"]].paths:
            if op.get(key) and screen_path is not None:
                try:
                    screen_path(str(op[key]))
                except Exception as exc:  # noqa: BLE001
                    exc.details = {  # type: ignore[attr-defined]
                        "failure": {"index": i, "op": op, "error": str(exc)},
                        "ops_run": 0,
                        "applied": [],
                    }
                    raise
    _precheck(pres, ops)
    ctrl = controller_for(pres)
    before = _layout(pres)
    token = _begin_undo_entry(pres)
    outputs: list[dict[str, Any]] = []
    for i, op in enumerate(ops):
        spec = OPS[op["op"]]
        kwargs = {spec.params[k]: v for k, v in op.items() if k in spec.params}
        try:
            result = _strip(getattr(ctrl, spec.method)(**kwargs))
        except Exception as exc:  # noqa: BLE001 - rolled back, or reported as applied
            message = com_message(exc) if _is_com_error(exc) else str(exc)
            problem = _undo_batch(pres, token, before)
            if problem is not None and token is not None:
                _end_undo_entry(pres)  # don't leave the marker behind
            err = _bad(
                i,
                op,
                message,
                applied=list(range(i)),
                rolled_back=problem is None,
                rollback_error=problem if i else None,
            )
            raise err from exc
        if result:
            outputs.append(
                {"index": i, **result}
                if isinstance(result, dict)
                else {
                    "index": i,
                    "result": result,
                }
            )
    if token is not None:
        _end_undo_entry(pres)
    out: dict[str, Any] = {"ok": True, "ops_run": len(ops)}
    if outputs:
        out["outputs"] = outputs
    if warnings:
        out["warnings"] = warnings
    return out


# -- reads -----------------------------------------------------------------------


def _status(pres: Any, p: dict[str, Any]) -> dict[str, Any]:
    app = pres.Application
    items = []
    for x in presentations(app):
        items.append(
            {
                "name": str(x.Name),
                "path": str(x.FullName),
                "active": str(x.FullName) == str(pres.FullName),
                "slides": int(x.Slides.Count),
                "saved": bool(x.Saved),
            }
        )
    out: dict[str, Any] = {
        "presentations": items,
        "active_presentation": str(pres.Name),
        "slide_width": float(pres.PageSetup.SlideWidth),
        "slide_height": float(pres.PageSetup.SlideHeight),
    }
    try:
        out["current_slide"] = int(app.ActiveWindow.View.Slide.SlideIndex)
    except Exception:  # noqa: BLE001 - no window, or not in a slide view
        pass
    return out


def _all_text(pres: Any, p: dict[str, Any]) -> list[dict[str, Any]]:
    ctrl = controller_for(pres)
    if p.get("slide"):
        return [ctrl.slide_get_text(int(p["slide"]))]
    return [ctrl.slide_get_text(i) for i in range(1, int(pres.Slides.Count) + 1)]


@dataclass(frozen=True)
class ReadSpec:
    run: Callable[[Any, dict[str, Any]], Any]
    required: tuple[str, ...] = ()


def _ctrl_read(method: str, *fields: str) -> Callable[[Any, dict[str, Any]], Any]:
    def run(pres: Any, p: dict[str, Any]) -> Any:
        ctrl = controller_for(pres)
        args = {_param(f): p[f] for f in fields if p.get(f) is not None}
        return getattr(ctrl, method)(**args)

    return run


def _param(name: str) -> str:
    return {v: k for k, v in _FIELD_NAMES.items()}.get(name, name)


READS: dict[str, ReadSpec] = {
    "status": ReadSpec(_status),
    "slides": ReadSpec(_ctrl_read("slide_list")),
    "slide": ReadSpec(_ctrl_read("slide_get_info", "slide"), ("slide",)),
    "text": ReadSpec(_all_text),
    "runs": ReadSpec(_ctrl_read("text_get_runs", "slide", "shape"), ("slide", "shape")),
    "notes": ReadSpec(_ctrl_read("slide_get_notes", "slide"), ("slide",)),
    "layouts": ReadSpec(_ctrl_read("list_layouts")),
    "table": ReadSpec(_ctrl_read("table_get_info", "slide", "shape"), ("slide", "shape")),
    "link": ReadSpec(
        _ctrl_read("hyperlink_get", "slide", "shape", "paragraph"), ("slide", "shape")
    ),
}


def dispatch_read(pres: Any, command: str, params: dict[str, Any]) -> Any:
    spec = READS.get(command)
    if spec is None:
        raise PptOpError(f"unknown read command {command!r}")
    missing = [f for f in spec.required if params.get(f) is None]
    if missing:
        raise PptOpError(f"read {command!r} needs {', '.join(missing)}")
    try:
        return spec.run(pres, params)
    except PowerPointError:
        raise
    except Exception as exc:  # noqa: BLE001
        if _is_com_error(exc):
            raise PowerPointError(com_message(exc)) from exc
        raise
