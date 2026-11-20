"""Excel batches: every op is checked first; a failure restores cell contents."""

from __future__ import annotations

import pytest

from office_mcp.excel import dispatch
from office_mcp.excel.dispatch import OpSpec, run_batch, validate_ops
from office_mcp.excel.errors import ExcelError, ExcelOpError, NotFoundError


@pytest.fixture
def fake_ops(monkeypatch, book):
    """Three test ops: `set` (restorable), `struct` (not restorable), `boom` (fails)."""

    def set_cell(c, o):
        sheet = dispatch._sheet_name(c, o.get("sheet"))
        c.book.sheets[sheet].cells[o["range"]] = o["value"]
        return {"success": True, "address": o["range"]}

    def boom(c, o):
        raise ExcelError("Excel said no")

    monkeypatch.setitem(
        dispatch.OPS,
        "set",
        OpSpec(
            set_cell,
            ("range", "value"),
            ("sheet",),
            saves=lambda c, o: [(dispatch._sheet_name(c, o.get("sheet")), o["range"])],
        ),
    )
    monkeypatch.setitem(dispatch.OPS, "struct", OpSpec(lambda c, o: book.events.append("struct")))
    monkeypatch.setitem(dispatch.OPS, "boom", OpSpec(boom))


def test_a_good_batch_reports_outputs_without_success_noise(book, fake_ops):
    out = run_batch(book, [{"op": "set", "range": "A1", "value": 1}])
    assert out == {"ok": True, "ops_run": 1, "outputs": [{"index": 0, "address": "A1"}]}
    assert book.sheets["Data"].cells == {"A1": 1}


def test_failure_restores_overwritten_cells(book, fake_ops):
    book.sheets["Data"].cells["A1"] = "before"
    with pytest.raises(ExcelError) as info:
        run_batch(
            book,
            [
                {"op": "set", "range": "A1", "value": "after"},
                {"op": "set", "range": "B2", "value": 5},
                {"op": "boom"},
            ],
        )
    assert book.sheets["Data"].cells == {"A1": "before", "B2": ""}
    details = info.value.details
    assert details["failure"]["index"] == 2
    assert details["ops_run"] == 2
    assert details["rolled_back"] is True
    assert details["not_rolled_back"] == []


def test_failure_after_a_structural_op_says_what_stays(book, fake_ops):
    with pytest.raises(ExcelError) as info:
        run_batch(
            book, [{"op": "struct"}, {"op": "set", "range": "A1", "value": 1}, {"op": "boom"}]
        )
    details = info.value.details
    assert details["rolled_back"] is False
    assert details["not_rolled_back"] == [0]
    assert book.sheets["Data"].cells == {"A1": ""}


def test_a_missing_sheet_is_not_found_and_still_rolls_back(book, fake_ops):
    with pytest.raises(NotFoundError) as info:
        run_batch(
            book,
            [
                {"op": "set", "range": "A1", "value": 1},
                {"op": "set", "sheet": "Nope", "range": "A1", "value": 2},
            ],
        )
    assert info.value.kind == "sheet"
    assert info.value.details["failure"]["index"] == 1
    assert book.sheets["Data"].cells == {"A1": ""}


def test_sheet_names_match_without_case(book, fake_ops):
    run_batch(book, [{"op": "set", "sheet": "notes", "range": "C3", "value": "x"}])
    assert book.sheets["Notes"].cells == {"C3": "x"}


def test_alerts_are_off_during_the_batch_and_restored_after(book, fake_ops, monkeypatch):
    seen = []
    monkeypatch.setitem(
        dispatch.OPS, "peek", OpSpec(lambda c, o: seen.append(book.app.api.DisplayAlerts))
    )
    run_batch(book, [{"op": "peek"}])
    assert seen == [False]
    assert book.app.api.DisplayAlerts is True
    with pytest.raises(ExcelError):
        run_batch(book, [{"op": "boom"}])
    assert book.app.api.DisplayAlerts is True


@pytest.mark.parametrize(
    ("op", "message"),
    [
        ({"op": "nope"}, "unknown op"),
        ({"op": "write", "range": "A1"}, "needs values"),
        ({"op": "format", "range": "A1", "align": "middle"}, "align must be one of"),
        ({"op": "format", "range": "A1", "fill": "blue"}, "fill must be a color"),
        ({"op": "create_chart", "type": "pie3d", "range": "A1:B3"}, "type must be one of"),
        ({"op": "pivot_add_field", "pivot": "P", "field": "F", "area": "rows"}, "area must be"),
        ("write", "must be an object"),
    ],
)
def test_bad_ops_fail_before_anything_runs(book, fake_ops, op, message):
    with pytest.raises(ExcelOpError) as info:
        run_batch(book, [{"op": "set", "range": "A1", "value": 1}, op])
    assert message in str(info.value)
    assert info.value.details["failure"]["index"] == 1
    assert info.value.details["ops_run"] == 0
    assert book.sheets["Data"].cells == {}


def test_ignored_fields_are_warned_about():
    warnings = validate_ops([{"op": "merge", "range": "A1:B1", "colour": "red"}])
    assert warnings == [{"index": 0, "op": "merge", "ignored_fields": ["colour"]}]


def test_an_empty_batch_is_refused():
    with pytest.raises(ExcelOpError):
        validate_ops([])


def test_colors_accept_hex_and_rgb_lists():
    assert dispatch._rgb("#1F4E79", "fill") == (31, 78, 121)
    assert dispatch._rgb("1f4e79", "fill") == (31, 78, 121)
    assert dispatch._rgb([1, 2, 3], "fill") == (1, 2, 3)
    with pytest.raises(ExcelOpError):
        dispatch._rgb([1, 2, 300], "fill")


def test_values_are_shaped_into_rows():
    assert dispatch._as_rows(5) == [[5]]
    assert dispatch._as_rows([1, 2]) == [[1, 2]]
    assert dispatch._as_rows([[1], [2]]) == [[1], [2]]


def test_com_errors_use_excels_description():
    class com_error(Exception):  # noqa: N801 - mirrors pywintypes.com_error
        pass

    exc = com_error(
        -2147352567, "Exception occurred.", (0, "Excel", "Bad range.", None, 0, 0), None
    )
    assert dispatch.com_message(exc) == "Bad range."
