---
name: wordcom-mcp
description: Read and edit the Microsoft Word document the user has open right now, through the Word MCP tools. Inspect structure (outline, paragraphs, tables), make edits (text, styles, lists, tables, images, charts, equations, comments, headers/footers), render pages to PNG so you can see the layout, and batch changes into one undo step that rolls back as a whole on failure. Use when the user wants to read, edit, or visually check a .docx that is open in Word on Windows.
---

# Word (MCP)

These tools drive a **running** Microsoft Word over COM (Windows only). They edit
the document the user has **open right now**, politely: the user's cursor,
selection and scroll position are preserved, and each call's edits collapse into
a single Ctrl-Z.

There are four calls:

| Call | Use it for |
| ---- | ---------- |
| `read(command, …)` | every read-only query (see **Reading**) |
| `exec(ops, label?, tracked?)` | **every edit**: a list of op objects applied as one undo step (see **Editing**) |
| `snapshot(pages? \| anchor?, …)` | render page(s) to PNG so you can *see* the layout |
| `write(command, …)` | saving (`save` / `save_as` / `export_pdf`) and toggling Track Changes (`track`) |

Every call takes an optional `doc` (a document name from `status`); it defaults
to the active document. This guide is also returned by `read("guide")`.

## First, orient yourself
1. `read("status")`: is Word reachable, and which documents are open.
2. `read("outline")`: heading tree with `heading:N` ids (`all_paragraphs: true` = every paragraph).
3. `read("paragraphs", start?, count?)`: every paragraph with `para:N` ids and character offsets.
4. `read("find", text)`: locate a phrase (tolerant of whitespace and smart quotes); returns `range:START-END` ids.

For a large document, `read("digest", budget?)` gives a token-budgeted overview
where every heading keeps its anchor id.

## Anchors: how you address things
Ops target stable **anchors**, never the live cursor. An anchor id is a short
string passed as `anchor_id`:

| Anchor id            | Refers to |
| -------------------- | --------- |
| `heading:N`          | the Nth *paragraph*, which must be a heading. Same index space as `para:N`, so the first heading is rarely `heading:1`: copy the id from `outline` |
| `para:N`             | the Nth paragraph, any kind |
| `bookmark:NAME`      | a bookmark |
| `pin:CODE`           | a **durable handle** minted by the `pin` / `pin_outline` ops; survives renumbering |
| `cc:NAME`            | a content control (by title, falling back to tag) |
| `footnote:N` / `endnote:N` | the Nth note's body (1-based) |
| `image:N`            | the Nth inline picture (1-based; see `images`) |
| `shape:N` / `textbox:N` | the Nth floating shape / text box (see `shapes`) |
| `equation:N` / `chart:N` | the Nth equation / chart, in document order |
| `table:N:R:C`        | row R, col C of the Nth table (all 1-based) |
| `table:N:row:R` / `table:N:col:C` | a whole row / column, as a styling target (`set_shading`, `set_borders`, `apply_style`, `format_run`) |
| `range:START-END`    | a raw character span (what `find` returns) |
| `header:S:WHICH` / `footer:S:WHICH` | header/footer of section S (`primary` / `first` / `even`) |
| `start` / `end`      | before the first / after the last paragraph (prepend / append targets) |

`heading:N`, `para:N`, `image:N`, `shape:N`, `equation:N` and `chart:N` are
**positional**: they renumber when an edit shifts the document. Re-read
(`outline` / `paragraphs` / `shapes` …) after a structural edit before reusing
them; an `anchor_not_found` error is the signal you skipped that, and it carries
a recovery hint. `bookmark:`, `cc:` and `pin:` ids are name-based and survive
edits. To get a durable handle on unnamed content, run the `pin` op on it
(`pin_outline` pins every heading at once).

## Reading
All of these are `read(command, …)`; parameters are keyword arguments.

- **Structure:** `status` · `outline` (`all_paragraphs`) · `paragraphs` (`start`, `count`) · `sections` (page setup per section) · `stats` (page/word/paragraph/table/image/comment/revision counts and `saved`) · `styles`.
- **Locate:** `find` (`text`, `in_anchor`): exact-ish locate, returns `range:` ids · `find_paragraphs` (`text`, `limit`, `min_score`): **fuzzy** paragraph search returning ranked `para:N` candidates · `location` (`anchor_id`): page span, line, column, `in_table` · `nearest_heading` (`anchor_id`, `direction`: `before`|`after`).
- **Text:** `read_text` (`anchor_id`, `view`: `raw`|`final`|`original`|`segments`). `final`/`original` resolve tracked changes as if accepted/rejected; `segments` is the per-run insert/delete breakdown · `read_section` (`anchor_id` **or** `heading` text): a heading's body · `between` (`start_anchor`, `end_anchor`, `inclusive`) · `read_bookmark` (`name`) · `read_cc` (`name`).
- **Whole document:** `to_markdown` / `to_html` (`within`): clean Markdown/HTML (headings, lists, bold/italic, GFM tables, `![alt](image:N)`, links); lossy by design · `digest` (`budget`, default ~6000 tokens; `depth`): headings verbatim, tables as stubs, body sampled to fit.
- **Tables:** `table_list` · `table_read` (`table`) · `table_records` (`table`): body rows as `{header: value}` dicts.
- **Objects:** `images` · `read_image` (`anchor_id`: `image:N`) returns base64 + mime · `shapes` · `equations` · `charts` · `footnotes` · `endnotes` · `hyperlinks` · `fields` · `comments` · `revisions` · `track` (is Track Changes on).
- **Metadata & design:** `properties` · `variables` · `theme` · `themes` (built-in theme names) · `format_info` (`anchor_id`): effective paragraph/character formatting with each value's style baseline and an `override` flag · `list_levels` (`anchor_id`).
- **Quality:** `proofing`: spelling/grammar errors (with `range:` ids) and readability stats · `lint` (`rules`, `within`): publishing defects (dangling headings, tables with no repeating header, split numbered lists, formatting drifted from the style), each fixable finding carrying the op that `regularize` would run.
- **Change tracking:** `checkpoint` (`include`: `text`|`text+style`|`text+format`, `within`) returns a fingerprint token; `diff` (`checkpoint`, or `cp_a` + `cp_b`) returns a change list (`replace`/`insert`/`delete`/`restyle`/`reformat`, `table_*`) with current anchors. This is how you answer "what changed" and verify your own edits landed; an unchanged doc returns `[]`.
- **This guide:** `guide`.

## Editing: `exec`
Every edit is an **op**: a JSON object with an `op` name and its fields. Send one
or many in a single `exec` call:

```json
{
  "label": "Update quarterly report",
  "ops": [
    {"op": "write_bookmark",   "name": "Address", "text": "123 Main St"},
    {"op": "insert_paragraph", "anchor_id": "heading:2", "text": "New section.", "style": "Body Text"},
    {"op": "insert_image",     "anchor_id": "heading:2", "base64": "<base64>", "wrap": "auto"},
    {"op": "find_replace",     "find": "Q3", "text": "Q4", "all": true}
  ]
}
```

- **One undo step.** All ops in a call become a single Ctrl-Z for the user,
  labelled with `label`. `tracked: true` records the whole batch as tracked changes.
- **All or nothing.** If any op fails, the ops before it are **rolled back**, so
  the document is exactly as it was before the call. The error reports `failure`
  (`index`, `op`, `error`, `type`; `matches` for an ambiguous find), `ops_run`,
  and `rolled_back`. Fix the failing op and resend the **whole** batch. (If
  `rolled_back` is `false`, Word could not undo the batch; re-read before retrying.)
- **Results.** A successful call returns `ok`, `ops_run`, and `outputs`: one entry
  per op that produced something (new table `index`, inserted block `range`,
  `pin`, new `shape:N`/`equation:N`/`chart:N`, footnote id …), each tagged with its
  op `index`.
- **Warnings, not silence.** A field an op doesn't use (a typo, or `style` on
  `append_inline`) is reported in `warnings` instead of being dropped silently.
- **Anchors resolve live.** Each op resolves its `anchor_id` against the document
  *as it is when that op runs*, so positional ids see the shifts made by earlier
  ops in the same batch. To insert several paragraphs after one fixed point,
  insert them in reverse order, use one `insert_block`, or anchor each to the
  previous insert's output.
- **Create, then target.** Add `bind: "slug"` (or `true` for a random code) to
  `insert_paragraph` / `insert_block` / `insert_section` / `insert_markdown` /
  `create_table` to mint a `pin:` handle on the new content (returned as `pin`).
  Any field written exactly as `$ops[N].field` is replaced with op N's output
  before the op runs: `[{"op":"create_table","anchor_id":"end","data":[["a","b"]]},
  {"op":"set_cell","table":"$ops[0].table","row":1,"col":1,"text":"x"}]`.
- **Before/after.** Insert ops add **after** their anchor by default. Pass
  `"before": true` (or `"where": "before"`) to insert above it.

### Value formats
- **Lengths:** points as a number (`12`) or a unit string (`"0.5in"`, `"2cm"`, `"14pt"`).
- **Colours:** a name (`"navy"`), hex (`"#1F3864"`), or `"r,g,b"`.
- **Text with inline formatting:** `runs` is `[{"text": "Bold", "bold": true}, {"text": " rest"}]`
  (`bold`, `italic`, `underline`, `style` per run).
- **Items** (`insert_block`, `insert_section.body`, `replace_section.body`): each is
  `"plain text"` or `{"text"|"runs": …, "style"?: …}`; item `text` accepts `**bold**`
  / `*italic*`.
- **Markdown** (`insert_markdown`, `replace_section.markdown`): a constrained subset:
  `#`/`##`/`###` headings, `-`/`*` bullets, `1.` numbers, blank-line paragraphs, inline
  `**bold**`/`*italic*`. No code fences, nested lists, or tables.

## Op reference
**Required fields in bold**, then optional ones. "± where" means the op also
accepts `before` / `after` / `where` to pick the side of the anchor.

### Text and paragraphs
| Op | Fields | Notes |
| -- | ------ | ----- |
| `insert_paragraph` | **anchor_id**, `text` \| `runs`, style, bind, ± where | new paragraph next to any anchor; exactly one of `text` / `runs` |
| `insert_block` | **anchor_id**, **items**, bind, ± where | several styled paragraphs in reading order; returns the block's `range` |
| `insert_section` | **anchor_id**, **heading**, **body**, level, bind, ± where | a `Heading {level}` plus its body (items shape) |
| `insert_markdown` | **anchor_id**, **markdown**, bind, ± where | constrained Markdown as real Word structure |
| `replace_section` | **anchor_id** (a heading), `body` \| `markdown` | rewrite a heading's body, keeping the heading |
| `delete_paragraph` | **anchor_id** | removes the paragraph(s) including the mark, so text closes up |
| `append` / `append_paragraph` | **text**, style | new last paragraph |
| `prepend` / `prepend_paragraph` | **text**, style | new first paragraph |
| `append_inline` / `prepend_inline` | **text** | continue the last / first paragraph (no style) |
| `replace` | **anchor_id**, **text** | overwrite an anchor's text |
| `find_replace` | **find**, **text**, in, all, occurrence | fuzzy find + replace. Several hits without `all`/`occurrence` fails with `ambiguous_match`. Inside a table, scope with `in: "table:N:R:C"` |
| `write_bookmark` | **name**, **text** | set an existing bookmark's text |
| `write_cc` | **name**, **text** | set a content control's text |
| `insert_break` | **anchor_id**, kind, ± where | `kind`: `page` (default) / `column` / `section_next` / `section_continuous` |
| `insert_footnote` / `insert_endnote` | **anchor_id**, **text**, ± where | returns the new `footnote:N` / `endnote:N` |

### Formatting and styles
| Op | Fields | Notes |
| -- | ------ | ----- |
| `apply_style` | **anchor_id**, **name** | style names: `read("styles")` |
| `format_paragraph` | **anchor_id**, alignment, left_indent, right_indent, first_line_indent, space_before, space_after, line_spacing, page_break_before, keep_together, keep_with_next, widow_control | `line_spacing`: a multiple (`1.5`), `single`/`double`, or an exact length (`"14pt"`). `page_break_before` is the reflow-safe way to start a paragraph on a new page |
| `format_run` | **anchor_id**, bold, italic, underline, strikethrough, font, size, color, highlight, subscript, superscript, small_caps, all_caps, spacing | character formatting; pair with a `range:` from `find` to style a phrase |
| `set_shading` | **anchor_id**, fill, pattern | paragraph, cell, row or column fill |
| `set_borders` | **anchor_id**, sides, style (alias line_style), weight, color | `sides`: `all`/`top`/`bottom`/`left`/`right`/…; `style`: `single`/`double`/`dot`/`dash` |
| `add_tab_stop` | **anchor_id**, **position**, align, leader | `leader`: `dots`/… |
| `drop_cap` | **anchor_id**, lines, position, distance, font | `position`: `dropped`/`margin`/`none` (removes it) |
| `add_style` | **name**, type, based_on, next_style | `type`: `paragraph`/`character` |
| `set_style` | **name**, + any `format_run` field except highlight, + any `format_paragraph` field, based_on, next_style | define the look once, then `apply_style` it everywhere |
| `regularize` | rules, within, dry_run | apply the fixable `lint` findings |

### Lists
| Op | Fields | Notes |
| -- | ------ | ----- |
| `apply_list` | **anchor_id**, type, continue_previous (alias continue) | `type`: `bulleted` (default) / `numbered` / `outline`. To number N paragraphs 1..N, apply once to a `range:` spanning them all; per-paragraph applies make N separate "1." lists |
| `apply_list_format` | **anchor_id**, **levels**, continue_previous | custom multi-level template. Each level: `kind` (`number`/`bullet`), `format` (`"%1."`, `"%1.%2"`, or a bullet glyph), `style` (`arabic`/`upper-roman`/`lower-letter`/…), `start_at`, `number_position`, `text_position`, `trailing`, `alignment`, `font`, `bold`, `italic`, `color` |
| `remove_list` · `restart_numbering` · `indent_list` · `outdent_list` | **anchor_id** | |

### Tables
`table` is the 1-based table index (`read("table_list")`). Cells are anchors
(`table:N:R:C`), so `replace`, `apply_style`, `format_paragraph`, `set_shading`
etc. work on them too. After `merge_cells`/`split_cell` the table is non-uniform
and `table:N:R:C` addresses *physical* cells.

| Op | Fields | Notes |
| -- | ------ | ----- |
| `create_table` | **anchor_id**, rows, cols, style, data, header, bind, ± where | `data`: a row-major 2-D array **or** records (objects whose keys become the header row); `rows`/`cols` are inferred from `data`. `style` defaults to `Table Grid`; `header` bolds row 1. New cells use `Normal`, never the anchor's heading style. Returns the new `table` index |
| `delete_table` | **table** | |
| `set_cell` | **table**, **row**, **col**, **text** | |
| `add_row` / `add_column` | **table**, values | |
| `delete_row` | **table**, **row** | |
| `delete_column` | **table**, **column** | |
| `append_record` | **table**, **record** | `{header: value}` → new row |
| `update_row` | **table**, **key**, **values**, column | update the first row whose key column equals `key` |
| `merge_cells` | **table**, **from**, **to** | `from`/`to` as `"R:C"` |
| `split_cell` | **table**, **cell**, rows, cols | `cell` as `"R:C"` |
| `set_heading_row` | **table**, row, heading, allow_break | repeat a header row across pages |
| `autofit_table` | **table**, mode | `content` (default) / `window` / `fixed` |
| `set_table_style` | **table**, **style** | restyle first, then shade cells (a style reapply overwrites shading) |
| `set_table_alignment` | **table**, **alignment** | `left`/`center`/`right` |
| `set_table_borders` | **table**, sides, style (alias line_style), weight, color | whole grid; interior lines via `horizontal`/`vertical` sides |
| `set_table_banding` | **table**, first_row, last_row, first_column, last_column, banded_rows, banded_columns | toggles table-style options |
| `set_cell_vertical_alignment` | **anchor_id** (a cell), **align** | `top`/`center`/`bottom` |

A column op on a table with merged or mixed-width cells raises an error; style
those cells one by one via `table:N:R:C`.

### Images and floating shapes
| Op | Fields | Notes |
| -- | ------ | ----- |
| `insert_image` | **anchor_id**, **wrap**, `path` \| `base64`, block, width, height, alt_text, lock_aspect, ± where | see below |
| `set_image_alt_text` | **anchor_id**, **text** | inline `image:N` |
| `set_image_size` | **anchor_id**, width, height, lock_aspect | inline `image:N` |
| `set_image_crop` | **anchor_id**, left, top, right, bottom (aliases crop_*) | inline `image:N` |
| `insert_text_box` | **anchor_id**, **text**, width, height, wrap, font, size, bold, italic, alignment, fill, border, ± where | a floating pull quote; returns its `shape:N` |
| `set_shape_wrap` | **anchor_id**, wrap, side, distance_top, distance_bottom, distance_left, distance_right | `side`: `both`/`left`/`right`/`largest` |
| `set_shape_position` | **anchor_id**, left, top, relative_to | lengths or `"center"`; `relative_to`: `margin`/`page` |
| `set_shape_size` | **anchor_id**, width, height, lock_aspect | |
| `set_shape_crop` | **anchor_id**, left, top, right, bottom (aliases crop_*) | picture shapes only |
| `format_shape` | **anchor_id**, fill, border, border_weight | `border`: a colour, `true` (default outline) or `false` (none) |
| `set_shape_alt_text` | **anchor_id**, **text** | |
| `set_shape_text` | **anchor_id**, **text** | text boxes only |
| `set_shape_text_frame` | **anchor_id**, margin_left, margin_right, margin_top, margin_bottom, word_wrap | text boxes only |
| `set_shape_rotation` | **anchor_id**, **degrees** | absolute angle |
| `set_shape_z_order` | **anchor_id**, **order** | `front`/`back`/`forward`/`backward`; **renumbers `shape:N`**, so re-read `shapes` |
| `replace_shape_image` | **anchor_id**, `path` \| `base64` | swap a floating picture in place, keeping wrap/position/size/alt text |
| `group_shapes` | **shapes** (list of ≥ 2 `shape:N`) | |
| `ungroup_shape` | **anchor_id** | returns the members' `shape:N` ids |
| `delete_shape` | **anchor_id** | |

- Give the image as `base64` (preferred when you hold the bytes) **or** a local
  `path`. Non-local paths (UNC `\\host\…`, `//host/…`, URLs, `file://`) are refused,
  and the server may restrict paths further to configured image directories.
- `wrap` is required: `inline` (stays in the text flow), `auto` (floats: square if
  small, else top-and-bottom), or `square`/`tight`/`through`/`top-bottom`/`behind`/`front`.
- `block: true` puts the image on its own `Normal` paragraph instead of in the
  anchor's text run (with `before`, it lands above a heading, not mid-line).
- A floating wrap turns the image into a **`shape:N`** (returned in `outputs`); an
  `inline` image stays `image:N`.

### Equations and charts
| Op | Fields | Notes |
| -- | ------ | ----- |
| `insert_equation` | **anchor_id**, `unicodemath` \| `latex` \| `mathml`, display, ± where | exactly one dialect. `unicodemath` is Word's native linear form (`"x=(-b±√(b^2-4ac))/(2a)"`); `latex` needs the `equations` extra. `display: true` = centred `Equation` style. Own paragraph either way. Returns `equation:N` |
| `insert_chart` | **anchor_id**, **kind**, **data**, title, ± where | `kind`: `bar`/`pie`/`line`/`scatter`. `data`: `{"Q1": 10, "Q2": 25}`, or `[[x, y], …]` pairs for `scatter` (`line` takes either). **Needs Excel installed** (else `excel_not_available`, document untouched). The data link is then broken, so the chart is static. Returns `chart:N` |
| `format_chart` | **anchor_id**, title, legend, legend_position, chart_style, background, plot_background, font, font_size, font_color, data_labels, data_label_format, chart_type, gap_width, overlap, data_table | |
| `format_axis` | **anchor_id**, **which**, title, minimum, maximum, scale, number_format, gridlines | `which`: `value`/`y`/`category`/`x`; `scale`: `linear`/`log` |
| `add_trendline` | **anchor_id**, series, kind, display_equation, display_r_squared, forward, backward, order, period | `kind`: `linear`/`exponential`/`logarithmic`/`moving_average` (`period`)/`polynomial` (`order`)/`power` |
| `set_series_color` | **anchor_id**, **color**, series, point | |
| `format_series` | **anchor_id**, series, point, marker, marker_size, smooth, explosion, data_labels, data_label_size, data_label_color | |
| `add_error_bars` | **anchor_id**, series, kind, amount, include, axis | `kind`: `fixed`/`percent`/`stdev`/`sterror` |

Chart formatting ops work on the static chart and do **not** need Excel. All
their fields are optional: only what you pass is changed.

### Fields, references and document furniture
| Op | Fields | Notes |
| -- | ------ | ----- |
| `insert_field` | **anchor_id**, **kind**, text, ± where | `kind`: `page`/`numpages`/`date`/`time`/`filename`/`author`/`title`, or `field` with a raw code in `text`. Page numbers in a footer: `anchor_id: "footer:1:primary"` |
| `update_fields` | | refresh every field (TOC, page refs, cross-refs …) |
| `insert_toc` | **anchor_id**, levels, use_heading_styles, hyperlinks, ± where | page numbers fill in after `update_fields` |
| `insert_table_of_figures` | **anchor_id**, label, include_label, hyperlinks, right_align_page_numbers, ± where | from the captions of one `label` |
| `insert_caption` | **anchor_id**, label, text, position, ± where | auto-numbered `Figure`/`Table`/…; on a cell it captions the whole table |
| `insert_cross_reference` | **anchor_id**, **target**, kind, hyperlink, ± where | `target`: `bookmark:`/`heading:`/`footnote:`/`endnote:`; `kind`: `text`/`page`/`number`/`above_below` |
| `add_bookmark` | **name**, **anchor_id** | name starts with a letter; letters, digits and `_` only |
| `pin` | **anchor_id**, name | returns `pin:CODE` |
| `pin_outline` | levels | pins every heading; returns the `{heading:N: pin:CODE}` map |
| `add_hyperlink` | **anchor_id**, `url` \| `bookmark`, text, screen_tip | `text` inserts new linked text; without it the anchor's range is linked |
| `set_hyperlink` | **index**, address, sub_address, text, screen_tip | edit an existing link (`read("hyperlinks")`) |
| `mark_index_entry` | **anchor_id**, **entry**, cross_reference, bold, italic | `"main:sub"` nests |
| `insert_index` | **anchor_id**, columns, run_in, right_align_page_numbers, ± where | builds the index from marked entries |
| `set_bibliography_style` | **style** | `APA`/`MLA`/`Chicago`/`IEEE`/`Turabian` |
| `add_source` | **source_type**, tag, author, title, year, publisher, city, journal_name, volume, issue, pages, url, edition, doi, xml | `source_type`: `book`/`journal_article`/`web_site`/`case`/…; `author` as `"Last, First"`; returns the `tag` |
| `insert_citation` | **anchor_id**, **tag**, pages, prefix, suffix, volume, suppress_author, suppress_year, suppress_title, locale, ± where | |
| `insert_bibliography` | **anchor_id**, ± where | |
| `mark_citation` | **anchor_id**, **long_citation**, short_citation, category, ± where | legal citations; `category`: `cases`/`statutes`/`other`/`rules`/`treatises`/`regulations`/`constitutional` or 1-16 |
| `insert_table_of_authorities` | **anchor_id**, category, passim, keep_entry_formatting, entry_separator, page_range_separator, ± where | |
| `create_content_control` | **anchor_id**, kind, title, tag, items, where, lock_contents, lock_control | `kind`: `rich_text` (default)/`text`/`picture`/`combo_box`/`dropdown`/`date`/`checkbox`/`building_block`/`group`/`repeating_section`. `where`: `wrap` (default, surrounds the anchor)/`before`/`after`. A titled control is addressable as `cc:TITLE` |
| `set_cc_properties` | **anchor_id**, title, tag, lock_contents, lock_control | |
| `set_cc_items` | **anchor_id**, **items** | strings or `{text, value}` |

### Page layout, headers and theme
| Op | Fields | Notes |
| -- | ------ | ----- |
| `set_page_setup` | **section**, margins, top_margin, bottom_margin, left_margin, right_margin, gutter, orientation, paper_size, columns, column_spacing | `paper_size`: `letter`/`legal`/`tabloid`/`a3`/`a4`/`a5` |
| `write_header` / `write_footer` | **section**, **text**, which | `which`: `primary` (default)/`first`/`even` |
| `set_watermark` | **text**, font, color, layout, semitransparent | `layout`: `diagonal`/`horizontal`; replaces any existing one |
| `remove_watermark` | | |
| `apply_theme` | **theme** | a built-in name (`read("themes")`) or a local `.thmx` path |
| `set_theme_colors` | scheme, colors | `colors`: `{"accent1": "#1A73E8", "text1": "navy", …}` (`text1`/`background1`/`text2`/`background2`/`accent1`..`accent6`/`hyperlink`/`followed_hyperlink`) |
| `set_theme_fonts` | scheme, major, minor | heading / body fonts |

### Review, metadata
| Op | Fields | Notes |
| -- | ------ | ----- |
| `add_comment` | **anchor_id**, **text**, author | |
| `resolve_comment` / `delete_comment` | **index** | 1-based, from `read("comments")` |
| `accept_revision` / `reject_revision` | **index** | from `read("revisions")` |
| `accept_all_revisions` / `reject_all_revisions` | anchor_id | optionally scoped to one anchor |
| `set_property` | **name**, **value**, custom | built-in (Title, Author, …) or, with `custom: true`, a custom property |
| `delete_property` | **name** | custom properties only |
| `set_variable` | **name**, **value** | hidden `{ DOCVARIABLE }` storage |
| `delete_variable` | **name** | |

## Snapshot: see the layout
`snapshot(pages?, anchor?, dpi?, max_dim?, markup?)` exports a pixel-faithful PDF
of the live document and returns the requested pages as PNG images: real fonts,
spacing and page geometry. Use it to judge layout and styling.

- Target **at most one** of `pages` (`"4"` or `"2-5"`) or `anchor` (the page(s) it
  occupies; a `heading:` expands to its whole section). Neither = whole document.
- `max_dim` caps each page's long edge in pixels (only ever lowers resolution);
  ~1000 stays legible and keeps image tokens predictable. `dpi` is 72-300 (default 150).
- `markup: "all"` shows tracked changes and comments; the default `none` shows the
  final document. The structured list is `read("revisions")`.

## Saving: gated
- `write("save")` · `write("save_as", path, overwrite?)` (a `.docx`) ·
  `write("export_pdf", path, from_page?, to_page?)` (a PDF, the recommended deliverable).
- **Default-deny:** saving only writes inside directories the server was
  configured to allow. With none configured, every save fails with
  `path_not_allowed`. `save` needs a document that was saved once before;
  `save_as` refuses to overwrite unless `overwrite: true`.
- `write("track", on)` turns Track Changes on or off.

## Errors
Failures carry a stable `code` and a `retryable` flag:

| Code | Meaning | What to do |
| ---- | ------- | ---------- |
| `anchor_not_found` | anchor id doesn't resolve (often a stale positional id), or `find` matched nothing | re-read `outline` / `paragraphs`, then retry |
| `style_not_found` | unknown style name | check `read("styles")` |
| `ambiguous_match` | `find_replace` hit several places | add `occurrence: N` or `all: true` |
| `replace_verification` | a whole-document match inside a table couldn't be verified | scope with `in: "table:N:R:C"` |
| `word_busy` | a modal dialog is open in Word | **retryable**: wait and retry |
| `word_not_running` | Word isn't open | ask the user to open Word |
| `document_not_found` | no open document with that `doc` name | check `read("status")` |
| `excel_not_available` | `insert_chart` without Excel | only after Excel is installed |
| `path_not_allowed` | save/image path outside the allowed directories, or non-local | use an allowed local path, or base64 for images |
| `error` | anything else (bad input, missing field) | fix the input |

## Typical workflow
1. `read("status")` to confirm Word and the target document.
2. `read("outline")` / `paragraphs` / `find` to get the anchor ids you need.
3. One `exec` call per logical change. It applies as one undo step and rolls back
   as a whole if any op fails.
4. Verify: read back, `diff` against a `checkpoint`, or `snapshot` the page.
