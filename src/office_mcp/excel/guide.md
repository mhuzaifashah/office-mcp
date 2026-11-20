# Excel tools guide

The Excel tools work on the workbooks open in Excel on this PC. Every call acts
on the active workbook unless you pass `workbook` (its file name, as listed by
`excel_read(command="status")`).

## Addressing

- `sheet`: a sheet name, matched without regard to case. Omit it for the active sheet.
- `range`: an A1 address (`B2`, `A1:D20`, `C:C`, `3:3`) or a defined name.
  Whole columns and rows are clipped to the data when read.
- `table`, `chart`, `pivot`: names from the `tables`, `charts` and `pivots` reads.
- Colors: `"#1F4E79"` or `[31, 78, 121]`.

## Reads: `excel_read(command=...)`

| Command | Fields | Returns |
|---|---|---|
| `status` | | open workbooks, active workbook and sheet, selection |
| `sheets` | | each sheet's name, index, visibility, used range |
| `used_range` | sheet | the real data extent (ignores cleared cells) |
| `read` | sheet, range, formulas | values as rows; `formulas=true` adds the formulas. Omit `range` for the used range. At most 20,000 cells per call |
| `find` | **text**, sheet, range, match_case, whole_cell, limit | every matching cell (value or formula), across all sheets unless `sheet` is given |
| `summary` | sheet, **range** | count, sum, min, max, average, median, computed by Excel |
| `tables` | | tables with their sheet, range, columns and style |
| `table_data` | **table**, visible_only | a table's header and rows |
| `charts` · `pivots` · `names` | | charts, pivot tables, defined names |
| `pivot_fields` | **pivot** | a pivot table's fields and their areas |
| `calc_mode` | | automatic, manual or semiautomatic |

## Edits: `excel_exec(ops=[...])`

Each op is `{"op": "<name>", ...fields}`; required fields are in bold. Ops run
in order, and each sees the workbook as the earlier ops left it.

**Before anything runs**, every op is checked (name, required fields, allowed
values). A bad op fails the call with the workbook untouched.

**Excel can't undo automation.** Ctrl+Z won't undo these edits, and running
them clears Excel's undo history. If an op fails part-way through a batch, the
cell contents (values and formulas) that earlier ops overwrote are written back.
Formatting and structure (sheets, rows, columns, tables, charts, pivots, names)
are not; the error lists those ops in `not_rolled_back`. Before a risky change
to data the user cares about, read it first, or ask them to save a copy.

### Cells

| Op | Fields | Notes |
|---|---|---|
| `write` | **range**, **values**, sheet | `values` is rows (`[[1, 2], [3, 4]]`), one row (`[1, 2]`) or one value. Written from the range's top-left cell. Strings starting with `=` are formulas |
| `fill_formula` | **range**, **formula**, sheet | puts the formula in the first cell and fills it across the range, so references adjust row by row |
| `clear` | **range**, sheet, contents_only | `contents_only=true` keeps the formatting |
| `replace` | **find**, **text**, sheet, range | find and replace inside cells; the whole sheet without `range` |
| `sort` | **range**, sheet, key_column, ascending, has_header, custom_order | `key_column` is 1-based within the range |
| `format` | **range**, sheet, bold, italic, font_size, font_color, fill, number_format, align, valign, wrap | `align`: left/center/right; `valign`: top/center/bottom; `number_format` like `"#,##0.00"`, `"0%"`, `"yyyy-mm-dd"` |
| `merge` · `unmerge` | **range**, sheet | |
| `autofit` | **range**, sheet, columns, rows | |
| `insert_rows` · `insert_columns` | **range**, sheet, count | inserted before the range |
| `delete_rows` · `delete_columns` | **range**, sheet | |
| `hyperlink` | **range**, **url**, sheet, text | |
| `validation` | **range**, **formula**, sheet, type | `type`: list (default; `formula` like `"Yes,No"`), whole, decimal, date, text_length |
| `cond_format` | **range**, **rule**, sheet, operator, formula, value1, value2, font_color, fill | `rule`: cell_value, formula, duplicate, unique, top10, above_average. `operator` (cell_value): between, not_between, equal, not_equal, greater, less, greater_equal, less_equal |
| `cond_format_clear` | **range**, sheet | |

### Sheets

| Op | Fields | Notes |
|---|---|---|
| `add_sheet` | **name**, after, before | |
| `rename_sheet` | **sheet**, **name** | |
| `copy_sheet` | **sheet**, **name** | the copy goes last |
| `move_sheet` | **sheet**, before, after | |
| `delete_sheet` | **sheet** | can't be restored |
| `sheet_visibility` | **sheet**, **visibility** | visible, hidden, very_hidden |
| `tab_color` | **sheet**, **color** | |

### Tables

| Op | Fields | Notes |
|---|---|---|
| `create_table` | **range**, **name**, sheet, has_headers, style | a single cell grows to the data around it; `style` like `"TableStyleMedium2"` |
| `table_add_row` | **table**, values | one row of values |
| `table_add_column` | **table**, **name**, position | |
| `rename_table` · `delete_table` | **table** (+ **name**) | |
| `resize_table` | **table**, **range** | |
| `table_totals` | **table**, **show** | |
| `table_column_total` | **table**, **column**, function | sum, count, average, max, min, stddev, var, count_nums, none |
| `sort_table` | **table**, **column**, ascending, custom_order | |
| `filter_table` | **table**, **column**, **criteria** | criteria like `">100"` or `"North"` |
| `clear_table_filters` | **table** | |
| `table_style` | **table**, **style** | |

### Charts

| Op | Fields | Notes |
|---|---|---|
| `create_chart` | **type**, range or table, sheet, name, title, at, left, top, width, height | `type`: column_clustered, column_stacked, column_stacked_100, bar_clustered, bar_stacked, bar_stacked_100, line, line_markers, line_stacked, pie, pie_exploded, doughnut, area, area_stacked, scatter, scatter_lines, scatter_smooth, radar, combo, waterfall, treemap, histogram. `at` is a range the chart covers; without it the chart goes below the data. Returns the chart's `name` |
| `chart_title` | **chart**, **title** | |
| `chart_type` | **chart**, **type** | |
| `series_type` | **chart**, **series**, **type** | combo charts; `series` is a 1-based index or a series name |
| `series_axis` | **chart**, **series**, **axis** | primary or secondary |
| `chart_source` | **chart**, **range**, sheet | |
| `move_chart` | **chart**, left, top, width, height | points |
| `chart_legend` | **chart**, **show**, position | top, bottom, left, right, corner |
| `axis_title` | **chart**, **axis**, **title** | `axis`: category (x) or value (y) |
| `delete_chart` | **chart** | |

### Pivot tables

| Op | Fields | Notes |
|---|---|---|
| `create_pivot` | **name**, **dest_sheet**, **dest_cell**, range (+ sheet) or table | the source needs a header row |
| `pivot_add_field` | **pivot**, **field**, **area**, position, function, label, number_format | `area`: row, column, filter, value. `function` (value fields): sum, count, average, max, min, count_nums, stddev, var |
| `pivot_remove_field` | **pivot**, **field** | |
| `pivot_layout` | **pivot**, **layout** | compact, tabular, outline |
| `pivot_subtotals` | **pivot**, **field**, **show** | |
| `pivot_grand_totals` | **pivot**, rows, columns | |
| `pivot_sort` | **pivot**, **field**, ascending | |
| `pivot_filter` | **pivot**, **field**, **values** | the items to keep visible |
| `pivot_number_format` | **pivot**, **field**, **number_format** | |
| `refresh_pivot` · `delete_pivot` | **pivot** | |

### Names and calculation

| Op | Fields | Notes |
|---|---|---|
| `add_name` | **name**, **refers_to**, sheet | `refers_to` like `"=Sheet1!$A$1:$A$10"`; `sheet` makes it sheet-scoped |
| `update_name` | **name**, **refers_to** | |
| `delete_name` | **name** | |
| `calc_mode` | **mode** | automatic, manual, semiautomatic |
| `calculate` | sheet | recalculate the workbook, or one sheet |

## Seeing the result: `excel_snapshot`

Renders a range (default: the sheet's used range) as PNG, as it would print,
with gridlines and row and column headings so you can read addresses off it.
Wide ranges fit one page wide; long ones run to more pages. Use it to check
layout, formatting and charts after an edit.

## Files: `excel_write`

`open` (path), `new`, `save`, `save_as` (path, overwrite), `export_pdf` (path,
sheet, overwrite). Writing only works inside the folders the server allows; a
workbook that has never been saved needs `save_as`.

## Example

```json
{"ops": [
  {"op": "write", "sheet": "Sales", "range": "A1", "values": [["Region", "Q1", "Q2"], ["North", 120, 135], ["South", 98, 110]]},
  {"op": "fill_formula", "sheet": "Sales", "range": "D2:D3", "formula": "=B2+C2"},
  {"op": "write", "sheet": "Sales", "range": "D1", "values": "Total"},
  {"op": "format", "sheet": "Sales", "range": "A1:D1", "bold": true, "fill": "#DDEBF7"},
  {"op": "create_chart", "sheet": "Sales", "range": "A1:C3", "type": "column_clustered", "title": "Sales by region", "at": "F2:L16"}
]}
```
