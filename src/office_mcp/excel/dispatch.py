"""The Excel read commands and edit ops behind the MCP tools.

`dispatch_read(book, command, params)` answers one read. `run_batch(book, ops)`
applies a list of edit ops. Both run on the Excel COM thread with a workbook
from `session.find_book`.

Excel can't undo changes made through automation (they also clear the user's
own undo history). A batch therefore:

1. checks every op (name, required fields, allowed values) before changing
   anything, so a malformed op fails with the workbook untouched;
2. saves the cell contents (values and formulas) an op is about to overwrite;
3. when an op fails part-way, writes those contents back in reverse order and
   reports which earlier ops it could not undo (formatting and structure:
   sheets, rows, tables, charts, pivots, names).
"""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

from .constants import CHART_TYPE_MAP
from .controller import ExcelController, ExcelError
from .errors import ExcelOpError, NotFoundError
from .session import controller_for, open_books

__all__ = ["OPS", "READS", "dispatch_read", "run_batch", "validate_ops"]

# Largest area whose contents a batch saves for rollback. Bigger writes still
# run, but are reported as not restorable if a later op fails.
_MAX_SAVED_CELLS = 1_000_000

# Most cells one `read` returns before it asks the agent to narrow the range.
_MAX_READ_CELLS = 20_000

_COLOR_RE = re.compile(r"^#?([0-9A-Fa-f]{6})$")

_CHART_TYPE_NAMES = {v: k for k, v in CHART_TYPE_MAP.items()}


# -- small helpers ---------------------------------------------------------------


def com_message(exc: BaseException) -> str:
    """Excel's own description of a COM error, else the exception text."""
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3 and args[2][2]:
        return str(args[2][2]).strip()
    if len(args) >= 2 and isinstance(args[1], str):
        return args[1]
    return str(exc)


def _is_com_error(exc: BaseException) -> bool:
    return type(exc).__name__ == "com_error"


def _rgb(value: Any, field_name: str) -> tuple[int, int, int]:
    """`"#1F4E79"`, `"1F4E79"` or `[31, 78, 121]` -> (r, g, b)."""
    if isinstance(value, str):
        m = _COLOR_RE.match(value.strip())
        if m:
            n = int(m.group(1), 16)
            return (n >> 16) & 0xFF, (n >> 8) & 0xFF, n & 0xFF
    elif isinstance(value, (list, tuple)) and len(value) == 3:
        if all(isinstance(c, int) and 0 <= c <= 255 for c in value):
            return int(value[0]), int(value[1]), int(value[2])
    raise ExcelOpError(f'{field_name} must be a color like "#1F4E79" or [31, 78, 121]')


def _sheet_name(ctrl: ExcelController, name: str | None) -> str:
    """The sheet's exact name (matched case-insensitively), or the active sheet's."""
    if not name:
        return str(ctrl.book.sheets.active.name)
    wanted = name.lower()
    for sheet in ctrl.book.sheets:
        if str(sheet.name).lower() == wanted:
            return str(sheet.name)
    raise NotFoundError("sheet", name)


def _check(found: bool, kind: str, name: str) -> None:
    if not found:
        raise NotFoundError(kind, name)


def _table_exists(ctrl: ExcelController, name: str) -> bool:
    return any(t["name"].lower() == name.lower() for t in ctrl.table_list())


def _strip(result: Any) -> Any:
    """Drop the controller's `success: True` noise from a result dict."""
    if isinstance(result, dict):
        out = {k: v for k, v in result.items() if k != "success"}
        return out or None
    return result


def _shape(values: Any) -> tuple[int, int]:
    if not isinstance(values, list) or not values:
        return 1, 1
    if not isinstance(values[0], list):
        return 1, len(values)
    return len(values), max((len(r) for r in values), default=1)


# -- reads -----------------------------------------------------------------------


@dataclass(frozen=True)
class ReadSpec:
    run: Callable[[Any, dict[str, Any]], Any]
    fields: tuple[str, ...] = ()


def _status(book: Any, p: dict[str, Any]) -> dict[str, Any]:
    books = []
    active_name = str(book.name)
    for b in open_books():
        try:
            saved = bool(b.api.Saved)
        except Exception:  # noqa: BLE001
            saved = None
        books.append(
            {
                "name": str(b.name),
                "path": str(b.fullname),
                "active": str(b.name) == active_name,
                "saved": saved,
                "sheets": [str(s.name) for s in b.sheets],
            }
        )
    out: dict[str, Any] = {"workbooks": books, "active_workbook": active_name}
    try:
        out["active_sheet"] = str(book.sheets.active.name)
        sel = book.app.api.Selection
        out["selection"] = str(sel.Address) if sel is not None else None
    except Exception:  # noqa: BLE001 - a chart or shape may be selected
        pass
    return out


def _sheets(book: Any, p: dict[str, Any]) -> list[dict[str, Any]]:
    ctrl = controller_for(book)
    rows = ctrl.sheet_list()
    for row in rows:
        try:
            row["used_range"] = str(book.sheets[row["name"]].api.UsedRange.Address)
        except Exception:  # noqa: BLE001
            row["used_range"] = None
    return rows


def _used_range(book: Any, p: dict[str, Any]) -> dict[str, Any]:
    ctrl = controller_for(book)
    sheet = _sheet_name(ctrl, p.get("sheet"))
    return {"sheet": sheet, **ctrl.range_used(sheet)}


def _read(book: Any, p: dict[str, Any]) -> dict[str, Any]:
    ctrl = controller_for(book)
    sheet = _sheet_name(ctrl, p.get("sheet"))
    address = p.get("range")
    if not address:
        used = ctrl.range_used(sheet)
        if not used["row_count"]:
            return {"sheet": sheet, "address": None, "values": [], "note": "the sheet is empty"}
        address = used["address"]
    rng = ctrl._resolve_range(sheet, address, clip_to_used=True)
    cells = int(rng.api.Count)
    if cells > _MAX_READ_CELLS:
        raise ExcelOpError(
            f"{rng.address} holds {cells:,} cells; read at most {_MAX_READ_CELLS:,} at a time. "
            "Read a smaller range, or use summary for statistics."
        )
    out = ctrl.range_read(sheet, address)
    if p.get("formulas"):
        out["formulas"] = ctrl.range_read_formulas(sheet, address)["formulas"]
    return out


def _find(book: Any, p: dict[str, Any]) -> dict[str, Any]:
    """Every cell whose value or formula contains `text` (up to `limit`)."""
    ctrl = controller_for(book)
    text = str(p["text"])
    limit = int(p.get("limit") or 50)
    sheets = [_sheet_name(ctrl, p["sheet"])] if p.get("sheet") else [s.name for s in book.sheets]
    matches: list[dict[str, Any]] = []
    truncated = False
    for name in sheets:
        ws = book.sheets[name]
        area = ws.range(p["range"]).api if p.get("range") else ws.api.UsedRange
        # Pass every Find setting: Excel otherwise reuses whatever the user last
        # chose in its Find dialog.
        first = area.Find(
            What=text,
            After=area.Cells(area.Cells.Count),
            LookIn=-4123,  # xlFormulas: formulas and constants
            LookAt=1 if p.get("whole_cell") else 2,
            SearchOrder=1,
            SearchDirection=1,
            MatchCase=bool(p.get("match_case")),
        )
        if first is None:
            continue
        first_address = str(first.Address)
        cell = first
        while True:
            matches.append(
                {
                    "sheet": str(name),
                    "address": str(cell.Address),
                    "value": cell.Value2,
                    "formula": str(cell.Formula) if cell.HasFormula else None,
                }
            )
            if len(matches) >= limit:
                truncated = True
                break
            cell = area.FindNext(cell)
            if cell is None or str(cell.Address) == first_address:
                break
        if truncated:
            break
    return {"matches": matches, "truncated": truncated}


def _summary(book: Any, p: dict[str, Any]) -> dict[str, Any]:
    ctrl = controller_for(book)
    return ctrl.range_summary(_sheet_name(ctrl, p.get("sheet")), p["range"])


def _tables(book: Any, p: dict[str, Any]) -> list[dict[str, Any]]:
    return controller_for(book).table_list()


def _table_data(book: Any, p: dict[str, Any]) -> dict[str, Any]:
    ctrl = controller_for(book)
    _check(_table_exists(ctrl, p["table"]), "table", p["table"])
    return ctrl.table_get_data(p["table"], visible_only=bool(p.get("visible_only")))


def _charts(book: Any, p: dict[str, Any]) -> list[dict[str, Any]]:
    charts = controller_for(book).chart_list()
    for c in charts:
        c["chart_type"] = _CHART_TYPE_NAMES.get(c.get("chart_type"), c.get("chart_type"))
    return charts


def _pivots(book: Any, p: dict[str, Any]) -> list[dict[str, Any]]:
    return controller_for(book).pivot_list()


def _pivot_fields(book: Any, p: dict[str, Any]) -> dict[str, Any]:
    ctrl = controller_for(book)
    _check(
        any(x["name"].lower() == p["pivot"].lower() for x in ctrl.pivot_list()),
        "pivot table",
        p["pivot"],
    )
    return ctrl.pivot_list_fields(p["pivot"])


def _names(book: Any, p: dict[str, Any]) -> list[dict[str, Any]]:
    return controller_for(book).named_range_list()


def _calc_mode(book: Any, p: dict[str, Any]) -> dict[str, Any]:
    return controller_for(book).calc_get_mode()


READS: dict[str, ReadSpec] = {
    "status": ReadSpec(_status),
    "sheets": ReadSpec(_sheets),
    "used_range": ReadSpec(_used_range, ("sheet",)),
    "read": ReadSpec(_read, ("sheet", "range", "formulas")),
    "find": ReadSpec(_find, ("text", "sheet", "range", "match_case", "whole_cell", "limit")),
    "summary": ReadSpec(_summary, ("sheet", "range")),
    "tables": ReadSpec(_tables),
    "table_data": ReadSpec(_table_data, ("table", "visible_only")),
    "charts": ReadSpec(_charts),
    "pivots": ReadSpec(_pivots),
    "pivot_fields": ReadSpec(_pivot_fields, ("pivot",)),
    "names": ReadSpec(_names),
    "calc_mode": ReadSpec(_calc_mode),
}

_READ_REQUIRED: dict[str, tuple[str, ...]] = {
    "find": ("text",),
    "summary": ("range",),
    "table_data": ("table",),
    "pivot_fields": ("pivot",),
}


def dispatch_read(book: Any, command: str, params: dict[str, Any]) -> Any:
    """Run one read command against `book`."""
    spec = READS.get(command)
    if spec is None:
        raise ExcelOpError(f"unknown read command {command!r}")
    missing = [f for f in _READ_REQUIRED.get(command, ()) if params.get(f) in (None, "")]
    if missing:
        raise ExcelOpError(f"read {command!r} needs {', '.join(missing)}")
    try:
        return spec.run(book, params)
    except ExcelError:
        raise
    except Exception as exc:  # noqa: BLE001
        if _is_com_error(exc):
            raise ExcelError(com_message(exc)) from exc
        raise


# -- edit ops --------------------------------------------------------------------


@dataclass(frozen=True)
class OpSpec:
    """One edit op.

    `run(ctrl, op)` applies it. `saves(ctrl, op)` names the areas whose cell
    contents it may overwrite, as `(sheet, address)` pairs; ops without it
    change formatting or structure, which a failed batch can't restore.
    """

    run: Callable[[ExcelController, dict[str, Any]], Any]
    required: tuple[str, ...] = ()
    optional: tuple[str, ...] = ()
    choices: dict[str, tuple[str, ...]] = field(default_factory=dict)
    colors: tuple[str, ...] = ()
    saves: Callable[[ExcelController, dict[str, Any]], list[tuple[str, str]]] | None = None


def _s(ctrl: ExcelController, op: dict[str, Any]) -> str:
    return _sheet_name(ctrl, op.get("sheet"))


def _need_table(ctrl: ExcelController, name: str) -> str:
    _check(_table_exists(ctrl, name), "table", name)
    return name


def _need_chart(ctrl: ExcelController, name: str) -> str:
    _check(any(c["name"].lower() == name.lower() for c in ctrl.chart_list()), "chart", name)
    return name


def _need_pivot(ctrl: ExcelController, name: str) -> str:
    _check(any(x["name"].lower() == name.lower() for x in ctrl.pivot_list()), "pivot table", name)
    return name


# Areas each content op may overwrite.


def _area_of_write(ctrl: ExcelController, op: dict[str, Any]) -> list[tuple[str, str]]:
    sheet = _s(ctrl, op)
    rng = ctrl.book.sheets[sheet].range(op["range"])
    rows, cols = _shape(op["values"])
    rows = max(rows, int(rng.rows.count))
    cols = max(cols, int(rng.columns.count))
    return [(sheet, str(rng[0, 0].resize(rows, cols).address))]


def _area_of_range(ctrl: ExcelController, op: dict[str, Any]) -> list[tuple[str, str]]:
    sheet = _s(ctrl, op)
    return [(sheet, str(ctrl._resolve_range(sheet, op["range"], clip_to_used=True).address))]


def _area_of_replace(ctrl: ExcelController, op: dict[str, Any]) -> list[tuple[str, str]]:
    sheet = _s(ctrl, op)
    if op.get("range"):
        return _area_of_range(ctrl, op)
    return [(sheet, str(ctrl.book.sheets[sheet].api.UsedRange.Address))]


def _area_of_fill(ctrl: ExcelController, op: dict[str, Any]) -> list[tuple[str, str]]:
    sheet = _s(ctrl, op)
    return [(sheet, str(ctrl.book.sheets[sheet].range(op["range"]).address))]


def _sheet_rel(ctrl: ExcelController, op: dict[str, Any], key: str) -> str | None:
    return _sheet_name(ctrl, op[key]) if op.get(key) else None


_ALIGN_H = ("left", "center", "right")
_ALIGN_V = ("top", "center", "bottom")

OPS: dict[str, OpSpec] = {
    # cells
    "write": OpSpec(
        lambda c, o: c.range_write(_s(c, o), o["range"], _as_rows(o["values"])),
        ("range", "values"),
        ("sheet",),
        saves=_area_of_write,
    ),
    "fill_formula": OpSpec(
        lambda c, o: c.range_fill_formula(_s(c, o), o["range"], o["formula"]),
        ("range", "formula"),
        ("sheet",),
        saves=_area_of_fill,
    ),
    "clear": OpSpec(
        lambda c, o: c.range_clear(_s(c, o), o["range"], bool(o.get("contents_only"))),
        ("range",),
        ("sheet", "contents_only"),
        saves=_area_of_range,
    ),
    "replace": OpSpec(
        lambda c, o: c.range_replace(_s(c, o), o["find"], o["text"], o.get("range")),
        ("find", "text"),
        ("sheet", "range"),
        saves=_area_of_replace,
    ),
    "sort": OpSpec(
        lambda c, o: c.range_sort(
            _s(c, o),
            o["range"],
            int(o.get("key_column", 1)),
            bool(o.get("ascending", True)),
            bool(o.get("has_header", True)),
            o.get("custom_order"),
        ),
        ("range",),
        ("sheet", "key_column", "ascending", "has_header", "custom_order"),
        saves=_area_of_range,
    ),
    "format": OpSpec(
        lambda c, o: c.range_format(
            _s(c, o),
            o["range"],
            bold=o.get("bold"),
            italic=o.get("italic"),
            font_size=o.get("font_size"),
            font_color=_rgb(o["font_color"], "font_color") if "font_color" in o else None,
            bg_color=_rgb(o["fill"], "fill") if "fill" in o else None,
            number_format=o.get("number_format"),
            h_align=o.get("align"),
            v_align=o.get("valign"),
            wrap_text=o.get("wrap"),
        ),
        ("range",),
        (
            "sheet",
            "bold",
            "italic",
            "font_size",
            "font_color",
            "fill",
            "number_format",
            "align",
            "valign",
            "wrap",
        ),
        choices={"align": _ALIGN_H, "valign": _ALIGN_V},
        colors=("font_color", "fill"),
    ),
    "merge": OpSpec(lambda c, o: c.range_merge(_s(c, o), o["range"]), ("range",), ("sheet",)),
    "unmerge": OpSpec(lambda c, o: c.range_unmerge(_s(c, o), o["range"]), ("range",), ("sheet",)),
    "autofit": OpSpec(
        lambda c, o: c.range_autofit(
            _s(c, o), o["range"], bool(o.get("columns", True)), bool(o.get("rows", True))
        ),
        ("range",),
        ("sheet", "columns", "rows"),
    ),
    "insert_rows": OpSpec(
        lambda c, o: c.range_insert_rows(_s(c, o), o["range"], int(o.get("count", 1))),
        ("range",),
        ("sheet", "count"),
    ),
    "delete_rows": OpSpec(
        lambda c, o: c.range_delete_rows(_s(c, o), o["range"]), ("range",), ("sheet",)
    ),
    "insert_columns": OpSpec(
        lambda c, o: c.range_insert_columns(_s(c, o), o["range"], int(o.get("count", 1))),
        ("range",),
        ("sheet", "count"),
    ),
    "delete_columns": OpSpec(
        lambda c, o: c.range_delete_columns(_s(c, o), o["range"]), ("range",), ("sheet",)
    ),
    "hyperlink": OpSpec(
        lambda c, o: c.range_set_hyperlink(_s(c, o), o["range"], o["url"], o.get("text")),
        ("range", "url"),
        ("sheet", "text"),
    ),
    "validation": OpSpec(
        lambda c, o: c.range_set_validation(
            _s(c, o), o["range"], o.get("type", "list"), str(o.get("formula", ""))
        ),
        ("range", "formula"),
        ("sheet", "type"),
        choices={"type": ("list", "whole", "decimal", "date", "text_length")},
    ),
    "cond_format": OpSpec(
        lambda c, o: c.cond_format_add_rule(
            _s(c, o),
            o["range"],
            o["rule"],
            operator=o.get("operator"),
            formula=o.get("formula"),
            value1=o.get("value1"),
            value2=o.get("value2"),
            font_color=_rgb(o["font_color"], "font_color") if "font_color" in o else None,
            bg_color=_rgb(o["fill"], "fill") if "fill" in o else None,
        ),
        ("range", "rule"),
        ("sheet", "operator", "formula", "value1", "value2", "font_color", "fill"),
        choices={
            "rule": ("cell_value", "formula", "duplicate", "unique", "top10", "above_average"),
            "operator": (
                "between",
                "not_between",
                "equal",
                "not_equal",
                "greater",
                "less",
                "greater_equal",
                "less_equal",
            ),
        },
        colors=("font_color", "fill"),
    ),
    "cond_format_clear": OpSpec(
        lambda c, o: c.cond_format_clear(_s(c, o), o["range"]), ("range",), ("sheet",)
    ),
    # sheets
    "add_sheet": OpSpec(
        lambda c, o: c.sheet_create(
            o["name"], _sheet_rel(c, o, "after"), _sheet_rel(c, o, "before")
        ),
        ("name",),
        ("after", "before"),
    ),
    "rename_sheet": OpSpec(lambda c, o: c.sheet_rename(_s(c, o), o["name"]), ("sheet", "name")),
    "delete_sheet": OpSpec(lambda c, o: c.sheet_delete(_s(c, o)), ("sheet",)),
    "copy_sheet": OpSpec(lambda c, o: c.sheet_copy(_s(c, o), o["name"]), ("sheet", "name")),
    "move_sheet": OpSpec(
        lambda c, o: c.sheet_move(_s(c, o), _sheet_rel(c, o, "before"), _sheet_rel(c, o, "after")),
        ("sheet",),
        ("before", "after"),
    ),
    "sheet_visibility": OpSpec(
        lambda c, o: c.sheet_visibility(
            _s(c, o), o["visibility"] == "visible", o["visibility"] == "very_hidden"
        ),
        ("sheet", "visibility"),
        choices={"visibility": ("visible", "hidden", "very_hidden")},
    ),
    "tab_color": OpSpec(
        lambda c, o: c.sheet_tab_color(_s(c, o), _rgb(o["color"], "color")),
        ("sheet", "color"),
        colors=("color",),
    ),
    # tables
    "create_table": OpSpec(
        lambda c, o: c.table_create(
            _s(c, o), o["range"], o["name"], bool(o.get("has_headers", True)), o.get("style")
        ),
        ("range", "name"),
        ("sheet", "has_headers", "style"),
    ),
    "table_add_row": OpSpec(
        lambda c, o: c.table_add_row(_need_table(c, o["table"]), o.get("values")),
        ("table",),
        ("values",),
    ),
    "table_add_column": OpSpec(
        lambda c, o: c.table_add_column(_need_table(c, o["table"]), o["name"], o.get("position")),
        ("table", "name"),
        ("position",),
    ),
    "rename_table": OpSpec(
        lambda c, o: c.table_rename(_need_table(c, o["table"]), o["name"]), ("table", "name")
    ),
    "delete_table": OpSpec(lambda c, o: c.table_delete(_need_table(c, o["table"])), ("table",)),
    "resize_table": OpSpec(
        lambda c, o: c.table_resize(_need_table(c, o["table"]), o["range"]), ("table", "range")
    ),
    "table_totals": OpSpec(
        lambda c, o: c.table_toggle_totals(_need_table(c, o["table"]), bool(o["show"])),
        ("table", "show"),
    ),
    "table_column_total": OpSpec(
        lambda c, o: c.table_set_column_total(
            _need_table(c, o["table"]), o["column"], o.get("function", "sum")
        ),
        ("table", "column"),
        ("function",),
    ),
    "sort_table": OpSpec(
        lambda c, o: c.table_sort(
            _need_table(c, o["table"]),
            o["column"],
            bool(o.get("ascending", True)),
            o.get("custom_order"),
        ),
        ("table", "column"),
        ("ascending", "custom_order"),
    ),
    "filter_table": OpSpec(
        lambda c, o: c.table_filter(_need_table(c, o["table"]), o["column"], str(o["criteria"])),
        ("table", "column", "criteria"),
    ),
    "clear_table_filters": OpSpec(
        lambda c, o: c.table_clear_filters(_need_table(c, o["table"])), ("table",)
    ),
    "table_style": OpSpec(
        lambda c, o: c.table_set_style(_need_table(c, o["table"]), o["style"]), ("table", "style")
    ),
    # charts
    "create_chart": OpSpec(
        lambda c, o: _create_chart(c, o),
        ("type",),
        ("sheet", "range", "table", "name", "left", "top", "width", "height", "at", "title"),
        choices={"type": tuple(CHART_TYPE_MAP)},
    ),
    "chart_title": OpSpec(
        lambda c, o: c.chart_set_title(_need_chart(c, o["chart"]), o["title"]), ("chart", "title")
    ),
    "chart_type": OpSpec(
        lambda c, o: c.chart_set_type(_need_chart(c, o["chart"]), o["type"]),
        ("chart", "type"),
        choices={"type": tuple(CHART_TYPE_MAP)},
    ),
    "series_type": OpSpec(
        lambda c, o: c.chart_set_series_type(_need_chart(c, o["chart"]), o["series"], o["type"]),
        ("chart", "series", "type"),
        choices={"type": tuple(CHART_TYPE_MAP)},
    ),
    "series_axis": OpSpec(
        lambda c, o: c.chart_set_series_axis(_need_chart(c, o["chart"]), o["series"], o["axis"]),
        ("chart", "series", "axis"),
        choices={"axis": ("primary", "secondary")},
    ),
    "chart_source": OpSpec(
        lambda c, o: c.chart_set_source(_need_chart(c, o["chart"]), _s(c, o), o["range"]),
        ("chart", "range"),
        ("sheet",),
    ),
    "move_chart": OpSpec(
        lambda c, o: c.chart_move(
            _need_chart(c, o["chart"]), o.get("left"), o.get("top"), o.get("width"), o.get("height")
        ),
        ("chart",),
        ("left", "top", "width", "height"),
    ),
    "delete_chart": OpSpec(lambda c, o: c.chart_delete(_need_chart(c, o["chart"])), ("chart",)),
    "chart_legend": OpSpec(
        lambda c, o: c.chart_set_legend(
            _need_chart(c, o["chart"]), bool(o["show"]), o.get("position")
        ),
        ("chart", "show"),
        ("position",),
        choices={"position": ("top", "bottom", "left", "right", "corner")},
    ),
    "axis_title": OpSpec(
        lambda c, o: c.chart_set_axis_title(
            _need_chart(c, o["chart"]), "x" if o["axis"] in ("category", "x") else "y", o["title"]
        ),
        ("chart", "axis", "title"),
        choices={"axis": ("category", "value", "x", "y")},
    ),
    # pivot tables
    "create_pivot": OpSpec(
        lambda c, o: _create_pivot(c, o),
        ("name", "dest_sheet", "dest_cell"),
        ("sheet", "range", "table"),
    ),
    "pivot_add_field": OpSpec(
        lambda c, o: c.pivot_add_field(
            _need_pivot(c, o["pivot"]),
            o["field"],
            o["area"],
            o.get("position"),
            o.get("function", "sum"),
            o.get("label"),
            o.get("number_format"),
        ),
        ("pivot", "field", "area"),
        ("position", "function", "label", "number_format"),
        choices={
            "area": ("row", "column", "filter", "value"),
            "function": ("sum", "count", "average", "max", "min", "count_nums", "stddev", "var"),
        },
    ),
    "pivot_remove_field": OpSpec(
        lambda c, o: c.pivot_remove_field(_need_pivot(c, o["pivot"]), o["field"]),
        ("pivot", "field"),
    ),
    "refresh_pivot": OpSpec(lambda c, o: c.pivot_refresh(_need_pivot(c, o["pivot"])), ("pivot",)),
    "pivot_layout": OpSpec(
        lambda c, o: c.pivot_set_layout(_need_pivot(c, o["pivot"]), o["layout"]),
        ("pivot", "layout"),
        choices={"layout": ("compact", "tabular", "outline")},
    ),
    "pivot_subtotals": OpSpec(
        lambda c, o: c.pivot_set_subtotals(_need_pivot(c, o["pivot"]), o["field"], bool(o["show"])),
        ("pivot", "field", "show"),
    ),
    "pivot_grand_totals": OpSpec(
        lambda c, o: c.pivot_set_grand_totals(
            _need_pivot(c, o["pivot"]), o.get("rows"), o.get("columns")
        ),
        ("pivot",),
        ("rows", "columns"),
    ),
    "pivot_sort": OpSpec(
        lambda c, o: c.pivot_sort_field(
            _need_pivot(c, o["pivot"]), o["field"], bool(o.get("ascending", True))
        ),
        ("pivot", "field"),
        ("ascending",),
    ),
    "pivot_filter": OpSpec(
        lambda c, o: c.pivot_filter_field(
            _need_pivot(c, o["pivot"]), o["field"], [str(v) for v in o["values"]]
        ),
        ("pivot", "field", "values"),
    ),
    "pivot_number_format": OpSpec(
        lambda c, o: c.pivot_set_field_number_format(
            _need_pivot(c, o["pivot"]), o["field"], o["number_format"]
        ),
        ("pivot", "field", "number_format"),
    ),
    "delete_pivot": OpSpec(lambda c, o: c.pivot_delete(_need_pivot(c, o["pivot"])), ("pivot",)),
    # names and calculation
    "add_name": OpSpec(
        lambda c, o: c.named_range_create(
            o["name"],
            o["refers_to"],
            "sheet" if o.get("sheet") else "workbook",
            _sheet_rel(c, o, "sheet"),
        ),
        ("name", "refers_to"),
        ("sheet",),
    ),
    "update_name": OpSpec(
        lambda c, o: c.named_range_update(o["name"], o["refers_to"]), ("name", "refers_to")
    ),
    "delete_name": OpSpec(lambda c, o: c.named_range_delete(o["name"]), ("name",)),
    "calc_mode": OpSpec(
        lambda c, o: c.calc_set_mode(o["mode"]),
        ("mode",),
        choices={"mode": ("automatic", "manual", "semiautomatic")},
    ),
    "calculate": OpSpec(
        lambda c, o: c.calc_calculate(
            "sheet" if o.get("sheet") else "workbook", _sheet_rel(c, o, "sheet")
        ),
        (),
        ("sheet",),
    ),
}


def _as_rows(values: Any) -> list[list[Any]]:
    """`values` as rows: a scalar is one cell, a flat list is one row."""
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def _create_chart(c: ExcelController, o: dict[str, Any]) -> Any:
    if ("range" in o) == ("table" in o):
        raise ExcelOpError("create_chart needs exactly one of range or table")
    size = {k: o[k] for k in ("left", "top", "width", "height") if k in o}
    if "table" in o:
        table = _need_table(c, o["table"])
        sheet = _s(c, o) if o.get("sheet") else _table_sheet(c, table)
        out = c.chart_create_from_table(table, sheet, o["type"], o.get("name"), **size)
    else:
        out = c.chart_create(
            _s(c, o), o["range"], o["type"], o.get("name"), target_range=o.get("at"), **size
        )
    if o.get("title"):
        c.chart_set_title(out["name"], o["title"])
    return out


def _table_sheet(c: ExcelController, table: str) -> str:
    for t in c.table_list():
        if t["name"].lower() == table.lower():
            return str(t["sheet"])
    raise NotFoundError("table", table)


def _create_pivot(c: ExcelController, o: dict[str, Any]) -> Any:
    if ("range" in o) == ("table" in o):
        raise ExcelOpError("create_pivot needs exactly one of range or table")
    dest = _sheet_name(c, o["dest_sheet"])
    if "table" in o:
        return c.pivot_create_from_table(
            _need_table(c, o["table"]), dest, o["dest_cell"], o["name"]
        )
    return c.pivot_create_from_range(_s(c, o), o["range"], dest, o["dest_cell"], o["name"])


# -- validation ------------------------------------------------------------------


def validate_ops(ops: list[Any]) -> list[dict[str, Any]]:
    """Check every op before any runs. Returns warnings for fields an op ignores.

    Raises `ExcelOpError` naming the first bad op; nothing has changed yet.
    """
    if not isinstance(ops, list) or not ops:
        raise ExcelOpError("ops must be a non-empty list")
    warnings: list[dict[str, Any]] = []
    for i, op in enumerate(ops):
        if not isinstance(op, dict):
            raise _bad(i, op, "each op must be an object")
        name = op.get("op")
        spec = OPS.get(str(name))
        if spec is None:
            raise _bad(i, op, f"unknown op {name!r}")
        missing = [f for f in spec.required if op.get(f) in (None, "")]
        if missing:
            raise _bad(i, op, f"op {name!r} needs {', '.join(missing)}")
        for key, allowed in spec.choices.items():
            if key in op and op[key] is not None and op[key] not in allowed:
                raise _bad(i, op, f"{key} must be one of: {', '.join(allowed)}")
        for key in spec.colors:
            if key in op:
                try:
                    _rgb(op[key], key)
                except ExcelOpError as exc:
                    raise _bad(i, op, str(exc)) from None
        known = {"op", *spec.required, *spec.optional}
        ignored = sorted(set(op) - known)
        if ignored:
            warnings.append({"index": i, "op": name, "ignored_fields": ignored})
    return warnings


def _bad(index: int, op: Any, message: str) -> ExcelOpError:
    return ExcelOpError(
        f"op {index}: {message}",
        failure={"index": index, "op": op, "error": message},
        ops_run=0,
        rolled_back=False,
        not_rolled_back=[],
    )


# -- the batch -------------------------------------------------------------------


@dataclass
class _Saved:
    sheet: str
    address: str
    prop: str
    contents: Any


def _save(ctrl: ExcelController, areas: list[tuple[str, str]]) -> list[_Saved] | None:
    """Current contents of `areas`, or None when one is too large to keep."""
    saved: list[_Saved] = []
    for sheet, address in areas:
        rng = ctrl.book.sheets[sheet].range(address).api
        if int(rng.Count) > _MAX_SAVED_CELLS:
            return None
        try:
            saved.append(_Saved(sheet, address, "Formula2", rng.Formula2))
        except Exception:  # noqa: BLE001 - Excel before 365 has no Formula2
            saved.append(_Saved(sheet, address, "Formula", rng.Formula))
    return saved


def _restore(ctrl: ExcelController, saved: list[_Saved]) -> bool:
    ok = True
    for item in reversed(saved):
        try:
            rng = ctrl.book.sheets[item.sheet].range(item.address).api
            setattr(rng, item.prop, item.contents)
        except Exception:  # noqa: BLE001 - keep restoring the rest
            ok = False
    return ok


def run_batch(book: Any, ops: list[dict[str, Any]]) -> dict[str, Any]:
    """Apply `ops` in order. On failure, restore saved cell contents and raise.

    The raised error carries `details`: `failure` (index, op, error),
    `ops_run`, `rolled_back` (True when every applied op's changes were
    undone) and `not_rolled_back` (indices of applied ops whose changes stay).
    """
    warnings = validate_ops(ops)
    ctrl = controller_for(book)
    app_api = book.app.api
    # Some ops (deleting a sheet) ask for confirmation; with no one to click,
    # the call would hang. Alerts go back to the user's setting afterwards.
    alerts = app_api.DisplayAlerts
    app_api.DisplayAlerts = False
    saved: list[_Saved] = []
    unrestorable: list[int] = []
    outputs: list[dict[str, Any]] = []
    try:
        for i, op in enumerate(ops):
            spec = OPS[op["op"]]
            try:
                if spec.saves is not None:
                    kept = _save(ctrl, spec.saves(ctrl, op))
                    if kept is None:
                        unrestorable.append(i)
                    else:
                        saved.extend(kept)
                else:
                    unrestorable.append(i)
                result = _strip(spec.run(ctrl, op))
            except Exception as exc:  # noqa: BLE001 - every failure is reported and undone
                if unrestorable and unrestorable[-1] == i:
                    unrestorable.pop()  # the failing op itself made no change we know of
                restored = _restore(ctrl, saved)
                message = com_message(exc) if _is_com_error(exc) else str(exc)
                details = {
                    "failure": {"index": i, "op": op, "error": message},
                    "ops_run": i,
                    "rolled_back": restored and not unrestorable,
                    "not_rolled_back": unrestorable,
                }
                err = exc if isinstance(exc, ExcelError) else ExcelOpError(message)
                if isinstance(err, ExcelOpError):
                    err.details = {**getattr(err, "details", {}), **details}
                else:
                    err.details = details  # type: ignore[attr-defined]
                if err is exc:
                    raise
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
    finally:
        try:
            app_api.DisplayAlerts = alerts
        except Exception:  # noqa: BLE001
            pass
    out: dict[str, Any] = {"ok": True, "ops_run": len(ops)}
    if outputs:
        out["outputs"] = outputs
    if warnings:
        out["warnings"] = warnings
    return out
