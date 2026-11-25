"""The PowerPoint tools guide, served by `ppt_read(command="guide")`.

The op tables are built from `dispatch.OPS`, so every op's fields always match
the controller method behind it. Notes add what a field list can't say.
"""

from __future__ import annotations

from .dispatch import OPS, READS

_INTRO = """\
# PowerPoint tools guide

The PowerPoint tools work on the presentations open in PowerPoint on this PC.
Every call acts on the active presentation unless you pass `presentation`
(its file name, as listed by `ppt_read(command="status")`).

## Addressing

- `slide`: a 1-based slide number. Numbers shift when slides are added, moved
  or deleted; re-read `slides` after such changes.
- `shape`: a shape's numeric id from `ppt_read(command="slide", slide=N)` or
  from the labels `ppt_snapshot` draws. Ids stay fixed while the shape exists.
  `shapes` is a list of ids (group, align, distribute).
- Positions and sizes (`left`, `top`, `width`, `height`) are in points from the
  slide's top-left corner; `status` gives the slide size (960 x 540 for 16:9).
- Colors: `"#1F4E79"`, `"1F4E79"` or `[31, 78, 121]`.

## Reads: `ppt_read(command=...)`

| Command | Fields | Returns |
|---|---|---|
| `status` | | open presentations, slide size, current slide |
| `slides` | | every slide: number, layout, shape count, notes or not |
| `slide` | **slide** | the slide's shapes with id, type, position, size and text |
| `text` | slide | the text of one slide, or of every slide |
| `runs` | **slide**, **shape** | a shape's text as runs, with their font settings |
| `notes` | **slide** | the speaker notes |
| `layouts` | | layout names `add_slide` and `set_layout` accept |
| `table` | **slide**, **shape** | a table's size and cell text |
| `link` | **slide**, **shape**, paragraph | a shape's hyperlink |

## Edits: `ppt_exec(ops=[...])`

Each op is `{"op": "<name>", ...fields}`; required fields are in bold. Ops run
in order. Before anything runs, every op is checked (name, fields, allowed
values, and that the slides and shapes it names exist). The batch is one undo
step in PowerPoint: if an op fails part-way, the whole batch is undone and the
error says `rolled_back: true`; fix the op and send the batch again. If Undo
can't run (the presentation has no window, or you switched to another one),
the ops before the failure stay applied: the error lists them in `applied` and
says why in `rollback_error`; then send only the failing op and the rest.
Ops that add a shape return its `shape_id`; use it in later calls.
"""

_NOTES: dict[str, str] = {
    "add_slide": "`position` 1-based (default: last); `layout` a name from `layouts` (default blank)",
    "set_layout": "`layout` a name from `layouts`",
    "set_placeholder": "fills the slide's title, body, subtitle ... placeholder",
    "format_text": "`start`/`length` (1-based characters) or `paragraph` limit it to part of the text",
    "format_paragraph": "`paragraph` limits it to one paragraph; `indent_level` 1-5",
    "add_shape": "`text` is optional",
    "fill_transparency": "0.0 (opaque) to 1.0",
    "z_order": "",
    "group": "returns the group's id",
    "copy_format": "`target` may be one id or a list",
    "add_picture": "a local image file; omit `width` or `height` to keep the aspect ratio",
    "crop_picture": "points to trim from each side",
    "add_table": "returns the table's shape id; cells are 1-based (`row`, `col`)",
    "add_chart": '`categories` like `["Q1", "Q2"]`; `series` like '
    '`[{"name": "North", "values": [120, 135]}]`. Needs Excel installed',
    "chart_data": "replaces the chart's categories and series",
    "data_labels": "`series` limits it to one series",
    "add_connector": "from (`begin_x`, `begin_y`) to (`end_x`, `end_y`)",
    "connect": "glues the connector's ends to shapes; sites number the connection points",
    "set_link": "`paragraph` links just that paragraph's text",
    "set_background": "a color, a local picture, or `follow_master=true`",
    "set_footer": "without `slide`, every slide",
}

_SECTIONS: list[tuple[str, list[str]]] = [
    (
        "Slides",
        [
            "add_slide",
            "duplicate_slide",
            "move_slide",
            "delete_slide",
            "set_layout",
            "hide_slide",
            "set_notes",
            "set_background",
            "set_footer",
            "go_to_slide",
        ],
    ),
    (
        "Text",
        [
            "set_placeholder",
            "set_text",
            "append_text",
            "clear_text",
            "format_text",
            "format_paragraph",
        ],
    ),
    (
        "Shapes",
        [
            "add_shape",
            "add_text_box",
            "move_shape",
            "resize_shape",
            "rotate",
            "fill",
            "fill_transparency",
            "remove_fill",
            "line_color",
            "line_weight",
            "line_dash",
            "shadow",
            "z_order",
            "group",
            "ungroup",
            "align",
            "distribute",
            "copy_format",
            "delete_shape",
        ],
    ),
    ("Pictures", ["add_picture", "replace_picture", "crop_picture", "picture_border"]),
    (
        "Tables",
        [
            "add_table",
            "set_cell",
            "cell_fill",
            "cell_font",
            "cell_align",
            "column_width",
            "row_height",
            "merge_cells",
            "table_add_row",
            "table_add_column",
            "table_style",
        ],
    ),
    (
        "Charts",
        [
            "add_chart",
            "chart_data",
            "chart_title",
            "chart_type",
            "series_color",
            "data_labels",
            "chart_axis",
            "chart_legend",
            "gridlines",
            "chart_style",
        ],
    ),
    (
        "Connectors and links",
        ["add_connector", "connect", "disconnect", "reroute", "set_link", "remove_link"],
    ),
]

_OUTRO = """\
## Seeing the result: `ppt_snapshot`

Renders slides as PNG. By default each shape is outlined and labelled with its
id, so you can find the shape to edit; pass `shape_ids=false` for a clean
picture of what the audience sees. `slides` takes `"4"` or `"2-5"`.

## Files: `ppt_write`

`open` (path), `new`, `save`, `save_as` (path, overwrite: .pptx or .pptm),
`export_pdf` (path, overwrite). Writing only works inside the folders the
server allows; a presentation that has never been saved needs `save_as`.

## Example

```json
{"ops": [
  {"op": "add_slide", "layout": "title_only"},
  {"op": "set_placeholder", "slide": 2, "placeholder": "title", "text": "Q2 results"},
  {"op": "add_chart", "slide": 2, "chart_type": "column_clustered", "left": 60, "top": 120,
   "width": 840, "height": 380, "categories": ["Q1", "Q2"],
   "series": [{"name": "North", "values": [120, 135]}, {"name": "South", "values": [98, 110]}]}
]}
```
"""


def _row(name: str) -> str:
    spec = OPS[name]
    fields = [f"**{f}**" for f in spec.required] + list(spec.optional)
    notes = [_NOTES.get(name, "")]
    for key, allowed in spec.choices.items():
        notes.append(f"`{key}`: {', '.join(allowed)}")
    note = "; ".join(n for n in notes if n)
    return f"| `{name}` | {', '.join(fields)} | {note} |"


def guide_text() -> str:
    parts = [_INTRO]
    covered: set[str] = set()
    for title, names in _SECTIONS:
        parts.append(f"### {title}\n\n| Op | Fields | Notes |\n|---|---|---|")
        parts.extend(_row(n) for n in names)
        parts.append("")
        covered.update(names)
    rest = sorted(set(OPS) - covered)
    if rest:
        parts.append("### Other\n\n| Op | Fields | Notes |\n|---|---|---|")
        parts.extend(_row(n) for n in rest)
        parts.append("")
    parts.append(_OUTRO)
    return "\n".join(parts)


__all__ = ["READS", "guide_text"]
