"""Pin COM enum values to Microsoft's documented numbers.

Other tests compare against the same constants the code uses, so a wrong value
passes them. These assert the literal values from the Word / Excel object-model
reference instead.
"""

from __future__ import annotations

import pytest

from office_mcp.word.wordcom.constants import (
    WdFieldType,
    WdPaperSize,
    WdWrapType,
    XlChartType,
    XlErrorBarInclude,
    XlErrorBarType,
)


@pytest.mark.parametrize(
    ("member", "expected"),
    [
        (WdWrapType.SQUARE, 0),
        (WdWrapType.TIGHT, 1),
        (WdWrapType.THROUGH, 2),
        (WdWrapType.FRONT, 3),
        (WdWrapType.TOP_BOTTOM, 4),
        (WdWrapType.BEHIND, 5),
        (WdWrapType.INLINE, 7),
        (WdPaperSize.LETTER, 2),
        (WdPaperSize.LEGAL, 4),
        (WdPaperSize.A3, 6),
        (WdPaperSize.A4, 7),
        (WdPaperSize.A5, 9),
        (WdPaperSize.TABLOID, 23),
        (WdFieldType.TITLE, 15),
        (WdFieldType.AUTHOR, 17),
        (WdFieldType.PAGE, 33),
        (XlChartType.XY_SCATTER_MARKERS, -4169),
        (XlChartType.LINE, 4),
        (XlChartType.PIE, 5),
        (XlChartType.COLUMN_CLUSTERED, 51),
        (XlErrorBarInclude.BOTH, 1),
        (XlErrorBarInclude.NONE, -4142),
        (XlErrorBarType.FIXED_VALUE, 1),
        (XlErrorBarType.PERCENT, 2),
        (XlErrorBarType.STANDARD_DEVIATION, -4155),
        (XlErrorBarType.STANDARD_ERROR, 4),
        (XlErrorBarType.CUSTOM, -4114),
    ],
)
def test_enum_matches_office_reference(member, expected):
    assert int(member) == expected


def test_wrap_names_round_trip():
    from office_mcp.word.wordcom.content._shapes import WRAP_NAMES, WRAP_TO_NAME

    assert WRAP_TO_NAME[3] == "front"
    assert WRAP_TO_NAME[4] == "top-bottom"
    assert {WRAP_TO_NAME[int(v)] for v in WRAP_NAMES.values()} == set(WRAP_NAMES)
