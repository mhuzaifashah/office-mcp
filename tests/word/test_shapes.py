"""Floating-shape anchor model — `shape:N`, the restyle handle for text boxes,
floating images, and WordArt.

Floating shapes are COM-heavy; these pin the wiring against the `fake_word`
MagicMock — that the right anchor comes back from `insert_text_box` / a floating
`insert_image`, resolves via `doc.shapes` / `anchor_by_id`, and that each mutator
op reaches the right COM property across Python / CLI / exec / MCP. Visual
correctness (and the delete+reinsert `replace_image` swap) is a live-Word concern,
covered by the smoke pass.
"""

from __future__ import annotations

import base64

import pytest
import wordcom
from wordcom._dispatch import _build_write_op, _read_impl, _write_impl
from wordcom._worker import InlineWorker
from wordcom.exceptions import AnchorNotFoundError, OpError

# A 1x1 transparent PNG (same as test_images).
_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYPhfDwAChwGA"
    "60e6kgAAAABJRU5ErkJggg=="
)
_B64 = base64.b64encode(_PNG).decode("ascii")

W = InlineWorker()


@pytest.fixture
def png_file(tmp_path):
    p = tmp_path / "pic.png"
    p.write_bytes(_PNG)
    return p


# --- insert returns a resolvable shape anchor -----------------------------------


def test_insert_text_box_returns_shape_anchor(fake_word):
    from wordcom._anchors import ShapeAnchor

    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = doc.bookmarks["Address"].insert_text_box("Pull quote")
    assert isinstance(shape, ShapeAnchor)
    assert shape.anchor_id == "shape:1"
    assert shape.shape_type == "text_box"
    # Resolves both ways.
    assert doc.shapes[1].anchor_id == "shape:1"
    assert doc.anchor_by_id("shape:1").shape_type == "text_box"


def test_insert_image_floating_returns_shape_anchor(fake_word, png_file):
    from wordcom._anchors import ShapeAnchor

    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            shape = doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")
    assert isinstance(shape, ShapeAnchor)
    assert shape.shape_type == "picture"


def test_insert_image_inline_returns_none(fake_word, png_file):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            result = doc.bookmarks["Address"].insert_image(str(png_file), wrap="inline")
    assert result is None


# --- mutators -------------------------------------------------------------------


def _text_box(doc):
    return doc.bookmarks["Address"].insert_text_box("Body")


def test_set_wrap(fake_word):
    from wordcom.constants import WdWrapType

    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_wrap("tight")
    assert fake_word.ActiveDocument.Shapes.Item(1).WrapFormat.Type == int(WdWrapType.TIGHT)


def test_set_position_sets_left_top_and_frame(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_position(left="1in", top="2in", relative_to="page")
    com = fake_word.ActiveDocument.Shapes.Item(1)
    assert com.Left == 72.0 and com.Top == 144.0
    assert com.RelativeHorizontalPosition == 1 and com.RelativeVerticalPosition == 1


def test_set_position_center(fake_word):
    from wordcom.constants import WdShapePosition

    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_position(left="center")
    assert fake_word.ActiveDocument.Shapes.Item(1).Left == float(WdShapePosition.CENTER)


def test_set_size_honours_both_dims(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_size(width="3in", height="1in", lock_aspect=False)
    com = fake_word.ActiveDocument.Shapes.Item(1)
    assert com.Width == 216.0 and com.Height == 72.0


def test_set_size_both_dims_restores_prior_lock(fake_word):
    # Both dims, no explicit lock_aspect: aspect-lock is dropped for the set then
    # restored to its prior state — not left permanently FALSE (SHAPES-2).
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        com = fake_word.ActiveDocument.Shapes.Item(1)
        com.LockAspectRatio = -1  # locked beforehand
        shape.set_size(width="3in", height="1in")
    assert com.Width == 216.0 and com.Height == 72.0
    assert com.LockAspectRatio == -1  # restored, not left FALSE


def test_set_position_single_axis_leaves_other_frame(fake_word):
    # A horizontal-only reposition must not re-frame the vertical axis (SHAPES-5).
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        com = fake_word.ActiveDocument.Shapes.Item(1)
        com.RelativeVerticalPosition = 0  # margin
        shape.set_position(left="1in", relative_to="page")
    assert com.RelativeHorizontalPosition == 1  # page frame applied to moved axis
    assert com.RelativeVerticalPosition == 0  # untouched axis keeps its frame


def test_format_fill_and_border(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.format(fill="#eeeeff", border=False)
    com = fake_word.ActiveDocument.Shapes.Item(1)
    com.Fill.Solid.assert_called()
    assert com.Line.Visible == 0  # border=False hides the outline (MsoTriState.FALSE)


def test_set_alt_text(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_alt_text("A pull quote")
    assert fake_word.ActiveDocument.Shapes.Item(1).AlternativeText == "A pull quote"
    assert doc.shapes[1].alt_text == "A pull quote"


def test_set_text_replaces_contents(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_text("New body")
    assert fake_word.ActiveDocument.Shapes.Item(1).TextFrame.TextRange.Text == "New body"
    assert doc.shapes[1].text == "New body"


def test_shape_text_history_mirrors_text_not_paragraph(fake_word):
    # text_final / text_original / revision_segments reflect the shape's own text,
    # not the anchoring paragraph's revision history (SHAPES-4).
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_text("Quote body")
        assert shape.text_final == "Quote body"
        assert shape.text_original == "Quote body"
        assert shape.revision_segments() == [{"text": "Quote body", "change": None}]


def test_shape_mutators_are_polite_and_undo_wrapped(fake_word):
    # Shape mutators never touch Selection, and an edit op rides one UndoRecord (SHAPES-6).
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        # Run mutators OUTSIDE an edit scope so a stray .Select() can't be masked
        # by the scope's selection restore.
        fake_word.Selection.Start = 12
        fake_word.Selection.End = 18
        shape.set_position(left="1in", top="2in", relative_to="page")
        shape.set_size(width="2in", height="1in")
        shape.set_wrap("tight")
        assert fake_word.Selection.Start == 12 and fake_word.Selection.End == 18
        fake_word.Selection.Select.assert_not_called()
        # A representative op still rides exactly one UndoRecord under doc.edit.
        fake_word.UndoRecord.StartCustomRecord.reset_mock()
        fake_word.UndoRecord.EndCustomRecord.reset_mock()
        with doc.edit("reposition"):
            shape.set_position(left="3in")
    fake_word.UndoRecord.StartCustomRecord.assert_called_once_with("reposition")
    fake_word.UndoRecord.EndCustomRecord.assert_called_once()


def test_set_text_on_picture_raises(fake_word, png_file):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            shape = doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")
        with pytest.raises(OpError):
            shape.set_text("nope")


def test_replace_image_needs_a_picture(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        with pytest.raises(OpError):
            shape.replace_image(_PNG)


def test_replace_image_swaps_in_place(fake_word, png_file):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            shape = doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")
        before = fake_word.ActiveDocument.Shapes.Count
        shape.replace_image(_PNG)
    # Delete + reinsert keeps a single shape (no orphan).
    assert fake_word.ActiveDocument.Shapes.Count == before == 1
    assert doc.shapes[1].shape_type == "picture"


def test_delete_removes_the_shape(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        assert fake_word.ActiveDocument.Shapes.Count == 1
        shape.delete()
    assert fake_word.ActiveDocument.Shapes.Count == 0


# --- collections + resolution ---------------------------------------------------


def test_shapes_list_fields(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            doc.bookmarks["Address"].insert_text_box("Quote", width="2in", height="1in")
        rows = doc.shapes.list()
    assert len(rows) == 1
    row = rows[0]
    assert row["anchor_id"] == "shape:1"
    assert row["shape_type"] == "text_box"
    assert row["width"] == 144.0 and row["height"] == 72.0


def test_text_boxes_filter_keeps_unfiltered_id(fake_word, png_file):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("mix"):
            # shape:1 = floating image, shape:2 = text box
            doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")
            doc.bookmarks["Address"].insert_text_box("Quote")
        boxes = list(doc.text_boxes)
    assert len(boxes) == 1
    # The lone text box keeps its canonical shape:N id (its position among ALL shapes).
    assert boxes[0].anchor_id == "shape:2"
    assert doc.text_boxes[1].anchor_id == "shape:2"


def test_unknown_shape_index_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("shape:99")


def test_watermark_excluded_from_body_shapes(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("mix"):
            doc.set_watermark("DRAFT")  # header story
            doc.bookmarks["Address"].insert_text_box("Quote")  # body
        rows = doc.shapes.list()
    # Only the body text box; the header-story watermark is excluded.
    assert [r["shape_type"] for r in rows] == ["text_box"]


def test_body_name_prefixed_watermark_excluded(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            doc.bookmarks["Address"].insert_text_box("Quote")
        # A stray body shape named like Word's watermark is filtered by the guard.
        com = fake_word.ActiveDocument.Shapes
        com.AddTextbox(Anchor=fake_word.ActiveDocument.Range(0, 0))
        com.Item(2).Name = "PowerPlusWaterMarkObject9"
        rows = doc.shapes.list()
    assert len(rows) == 1


# --- exec ops -------------------------------------------------------------------


# --- CLI ------------------------------------------------------------------------


# --- MCP ------------------------------------------------------------------------


def test_mcp_read_shapes(fake_word):
    _write_impl(W, "text_box", {"anchor_id": "bookmark:Address", "text": "Q"})
    rows = _read_impl(W, "shapes", {})
    assert rows[0]["anchor_id"] == "shape:1"


def test_mcp_build_set_shape_position_op():
    op = _build_write_op(
        "set_shape_position", {"anchor_id": "shape:1", "left": "1in", "relative_to": "page"}
    )
    assert op == {
        "op": "set_shape_position",
        "anchor_id": "shape:1",
        "left": "1in",
        "relative_to": "page",
    }


def test_mcp_build_replace_shape_image_requires_one_source():
    with pytest.raises(OpError):
        _build_write_op("replace_shape_image", {"anchor_id": "shape:1"})
    op = _build_write_op("replace_shape_image", {"anchor_id": "shape:1", "image_base64": _B64})
    assert op == {"op": "replace_shape_image", "anchor_id": "shape:1", "base64": _B64}


def test_mcp_write_set_shape_wrap(fake_word):
    _write_impl(W, "text_box", {"anchor_id": "bookmark:Address", "text": "Q"})
    from wordcom.constants import WdWrapType

    _write_impl(W, "set_shape_wrap", {"anchor_id": "shape:1", "wrap": "behind"})
    assert fake_word.ActiveDocument.Shapes.Item(1).WrapFormat.Type == int(WdWrapType.BEHIND)


# --- shape depth: rotation / z-order / text-frame -------------------------------


def test_set_rotation(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_rotation("30")
    assert fake_word.ActiveDocument.Shapes.Item(1).Rotation == 30.0
    assert doc.shapes[1].rotation == 30.0


def test_set_rotation_bad_value_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        with pytest.raises(OpError):
            shape.set_rotation("sideways")


def test_set_z_order_brings_to_front(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("two"):
            doc.bookmarks["Address"].insert_text_box("A")  # shape:1 (backmost)
            doc.bookmarks["Address"].insert_text_box("B")  # shape:2 (frontmost)
        # shape:1 starts at the back; bring it to the front.
        assert doc.shapes[1].z_order == 1
        doc.shapes[1].set_z_order("front")
    # The shape formerly at shape:1 is now frontmost (highest ZOrderPosition).
    assert fake_word.ActiveDocument.Shapes.Item(2).ZOrderPosition == 2


def test_set_text_frame_margins_and_wrap(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_text_frame(margin_left="0.1in", margin_top=5, word_wrap=False)
    frame = fake_word.ActiveDocument.Shapes.Item(1).TextFrame
    assert frame.MarginLeft == 7.2 and frame.MarginTop == 5.0
    assert frame.WordWrap == 0  # word_wrap=False -> MsoTriState.FALSE


def test_set_text_frame_on_picture_raises(fake_word, png_file):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            shape = doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")
        with pytest.raises(OpError):
            shape.set_text_frame(margin_left=5)


def test_shapes_list_includes_rotation_and_z_order(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            _text_box(doc)
        rows = doc.shapes.list()
    assert rows[0]["rotation"] == 0.0
    assert rows[0]["z_order"] == 1


# --- group / ungroup ------------------------------------------------------------


def _two_boxes(doc):
    doc.bookmarks["Address"].insert_text_box("A")
    doc.bookmarks["Address"].insert_text_box("B")


def test_group_collapses_to_one_shape(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("grp"):
            _two_boxes(doc)
            group = doc.group_shapes("shape:1", "shape:2")
    assert group.shape_type == "group"
    # Two boxes became one group shape.
    assert len(doc.shapes) == 1
    assert doc.shapes[1].shape_type == "group"


def test_group_needs_two_shapes(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            _text_box(doc)
        with pytest.raises(OpError):
            doc.group_shapes("shape:1")


def test_group_rejects_non_shape(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            _text_box(doc)
        with pytest.raises(OpError):
            doc.group_shapes("shape:1", "para:1")


def test_ungroup_restores_members(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("grp"):
            _two_boxes(doc)
            group = doc.group_shapes("shape:1", "shape:2")
        with doc.edit("ungrp"):
            members = group.ungroup()
    assert len(members) == 2
    assert {m.anchor_id for m in members} == {"shape:1", "shape:2"}
    assert len(doc.shapes) == 2


def test_ungroup_on_non_group_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        with pytest.raises(OpError):
            shape.ungroup()


# --- inline image:N restyle -----------------------------------------------------


def test_image_set_alt_text(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("alt"):
            doc.images[1].set_alt_text("New caption")
    assert fake_word.ActiveDocument.InlineShapes.Item(1).AlternativeText == "New caption"
    assert doc.anchor_by_id("image:1").alt_text == "New caption"


def test_image_set_size(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("size"):
            doc.images[1].set_size(width="3in")
    assert fake_word.ActiveDocument.InlineShapes.Item(1).Width == 216.0


# --- textbox:N alias ------------------------------------------------------------


def test_textbox_alias_resolves_to_canonical_shape_id(fake_word, png_file):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("mix"):
            doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")  # shape:1
            doc.bookmarks["Address"].insert_text_box("Q")  # shape:2
        anchor = doc.anchor_by_id("textbox:1")
    # The alias resolves to the text box but reports its canonical shape:N id.
    assert anchor.anchor_id == "shape:2"
    assert anchor.shape_type == "text_box"


def test_textbox_alias_unknown_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("textbox:9")


# --- new exec / CLI / MCP wiring ------------------------------------------------


def test_mcp_build_set_shape_rotation_needs_degrees():
    with pytest.raises(OpError):
        _build_write_op("set_shape_rotation", {"anchor_id": "shape:1"})
    op = _build_write_op("set_shape_rotation", {"anchor_id": "shape:1", "degrees": 30})
    assert op == {"op": "set_shape_rotation", "anchor_id": "shape:1", "degrees": 30}


def test_mcp_build_group_shapes_requires_two():
    with pytest.raises(OpError):
        _build_write_op("group_shapes", {"shapes": ["shape:1"]})
    op = _build_write_op("group_shapes", {"shapes": ["shape:1", "shape:2"]})
    assert op == {"op": "group_shapes", "shapes": ["shape:1", "shape:2"]}


def test_mcp_build_set_image_size_op():
    op = _build_write_op("set_image_size", {"anchor_id": "image:1", "width": "2in"})
    assert op == {"op": "set_image_size", "anchor_id": "image:1", "width": "2in"}


# --- wrap side + text distance --------------------------------------------------


def test_set_wrap_side_and_distance(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_wrap(side="left", distance_top="0.1in", distance_right=4)
    wf = fake_word.ActiveDocument.Shapes.Item(1).WrapFormat
    assert wf.Side == 1  # left
    assert wf.DistanceTop == 7.2 and wf.DistanceRight == 4.0


def test_set_wrap_requires_an_argument(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        with pytest.raises(OpError):
            shape.set_wrap()


def test_set_wrap_bad_side_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        with pytest.raises(OpError):
            shape.set_wrap(side="sideways")


def test_shapes_list_includes_wrap_side_and_crop(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_wrap(side="right")
        rows = doc.shapes.list()
    assert rows[0]["wrap_side"] == "right"
    assert rows[0]["crop"] is None  # a text box carries no picture crop


def test_mcp_build_set_shape_wrap_op_optional_wrap():
    op = _build_write_op(
        "set_shape_wrap", {"anchor_id": "shape:1", "side": "left", "distance_top": "0.1in"}
    )
    assert op == {
        "op": "set_shape_wrap",
        "anchor_id": "shape:1",
        "side": "left",
        "distance_top": "0.1in",
    }


# --- cropping -------------------------------------------------------------------


def _floating_image(doc, png_file):
    return doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")


def test_set_crop_on_picture(fake_word, png_file):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            shape = _floating_image(doc, png_file)
        shape.set_crop(left="0.2in", bottom=3)
    pf = fake_word.ActiveDocument.Shapes.Item(1).PictureFormat
    assert pf.CropLeft == 14.4 and pf.CropBottom == 3.0


def test_set_crop_on_non_picture_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        with pytest.raises(OpError):
            shape.set_crop(left=5)


def test_set_crop_requires_an_edge(fake_word, png_file):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            shape = _floating_image(doc, png_file)
        with pytest.raises(OpError):
            shape.set_crop()


def test_shapes_list_includes_crop_for_picture(fake_word, png_file):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            shape = _floating_image(doc, png_file)
        shape.set_crop(left=6)
        rows = doc.shapes.list()
    assert rows[0]["crop"] == {"left": 6.0, "top": 0.0, "right": 0.0, "bottom": 0.0}


def test_image_set_crop(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("crop"):
            doc.images[1].set_crop(right="0.2in")
    assert fake_word.ActiveDocument.InlineShapes.Item(1).PictureFormat.CropRight == 14.4


def test_mcp_build_set_shape_crop_maps_crop_params():
    op = _build_write_op(
        "set_shape_crop", {"anchor_id": "shape:1", "crop_left": "0.2in", "crop_bottom": 3}
    )
    assert op == {"op": "set_shape_crop", "anchor_id": "shape:1", "left": "0.2in", "bottom": 3}


def test_mcp_build_set_image_crop_maps_crop_params():
    op = _build_write_op("set_image_crop", {"anchor_id": "image:1", "crop_right": 4})
    assert op == {"op": "set_image_crop", "anchor_id": "image:1", "right": 4}
