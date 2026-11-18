"""Tables with vertically merged cells.

Real Word raises error 5991 on `Table.Rows(r)` once any cell is merged
vertically. Reads must fall back to walking `Range.Cells` instead of failing.
"""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock

from office_mcp.word.wordcom._tables import Table
from tests.word.conftest import _make_range


def _vmerged_table_com():
    # 2x2 grid where (1,1) is merged down over (2,1):
    #   | A  | B |
    #   | ^^ | C |
    texts = {(1, 1): "A", (1, 2): "B", (2, 2): "C"}
    cells = {}
    for (r, c), text in texts.items():
        rng = _make_range(0, 0)
        rng.Text = text + "\r\x07"
        cells[(r, c)] = SimpleNamespace(RowIndex=r, ColumnIndex=c, Range=rng)

    def rows(_index):
        raise RuntimeError("Cannot access individual rows: vertically merged cells (5991)")

    rows_view = MagicMock(side_effect=rows)
    rows_view.Count = 2
    com = MagicMock(name="VMergedTable")
    com.Rows = rows_view
    com.Columns.Count = 2
    com.Range.Cells = list(cells.values())
    com.Cell.side_effect = lambda r, c: cells[(r, c)]
    return com


def test_grid_reads_vertically_merged_table():
    table = Table(MagicMock(), _vmerged_table_com(), 1)
    assert table.grid() == [["A", "B"], ["C"]]


def test_read_uses_real_column_indices():
    table = Table(MagicMock(), _vmerged_table_com(), 1)
    dump = table.read()
    assert [c["anchor_id"] for c in dump["cells"][1]] == ["table:1:2:2"]
    assert dump["uniform"] is False
