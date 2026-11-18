"""find/find_replace offsets stay exact around fields and emoji (W18).

A small position model stands in for Word: every document position holds one
UTF-16 unit, and positions that belong to a field's code or its marks are
hidden from `Range.Text`, exactly as Word reports them with field codes off.
"""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock

from office_mcp.word.wordcom._document import Document


class _Model:
    def __init__(self):
        self.units: list[tuple[str, bool]] = []  # (text for this position, visible)
        self.fields: list[SimpleNamespace] = []
        self.writes: list[tuple[int, int, str]] = []

    def text(self, s):
        for ch in s:
            if ord(ch) > 0xFFFF:  # astral: two UTF-16 units, one character
                self.units += [(ch, True), ("", True)]
            else:
                self.units.append((ch, True))

    def field(self, code, result):
        begin = len(self.units)
        self.units.append(("", False))  # begin mark
        code_start = len(self.units)
        self.units += [(c, False) for c in code]
        code_end = len(self.units)
        self.units.append(("", False))  # separator
        result_start = len(self.units)
        self.text(result)
        result_end = len(self.units)
        self.units.append(("", False))  # end mark
        self.fields.append(
            SimpleNamespace(
                Code=SimpleNamespace(Start=code_start, End=code_end),
                Result=SimpleNamespace(Start=result_start, End=result_end),
            )
        )
        return begin

    def range(self, a, b):
        model = self

        class _R:
            Start, End = a, b
            Tables = SimpleNamespace(Count=0)
            Fields = [f for f in model.fields if f.Code.Start - 1 >= a and f.Code.End <= b]

            @property
            def Text(self):
                return "".join(t for t, vis in model.units[a:b] if vis)

            @Text.setter
            def Text(self, value):
                model.writes.append((a, b, value))

        return _R()

    def document(self):
        com = MagicMock()
        com.Range.side_effect = self.range
        com.Content = self.range(0, len(self.units))
        return Document(MagicMock(), com)


def test_match_after_a_field_maps_to_its_real_position():
    m = _Model()
    m.text("Page ")
    m.field(" PAGE ", "7")
    m.text(" of target")
    hits = m.document().find("target")
    assert len(hits) == 1
    start, end = hits[0]["start"], hits[0]["end"]
    assert m.range(start, end).Text == "target"


def test_text_inside_a_field_result_is_found_at_its_position():
    m = _Model()
    m.text("See ")
    m.field(" REF Intro \\h ", "Introduction")
    m.text(".")
    hit = m.document().find("Introduction")[0]
    assert m.range(hit["start"], hit["end"]).Text == "Introduction"


def test_nested_field_result_is_found():
    m = _Model()
    m.text("x ")
    outer_begin = len(m.units)
    m.units.append(("", False))
    code_start = len(m.units)
    m.units += [(c, False) for c in " IF "]
    code_end = len(m.units)
    m.units.append(("", False))
    result_start = len(m.units)
    m.text("a ")
    m.field(" DATE ", "today")
    m.text(" b")
    result_end = len(m.units)
    m.units.append(("", False))
    m.fields.insert(
        0,
        SimpleNamespace(
            Code=SimpleNamespace(Start=code_start, End=code_end),
            Result=SimpleNamespace(Start=result_start, End=result_end),
        ),
    )
    assert outer_begin == code_start - 1
    for word in ("today", "b"):
        hit = m.document().find(word)[0]
        assert m.range(hit["start"], hit["end"]).Text == word


def test_match_after_emoji_counts_utf16_units():
    m = _Model()
    m.text("Launch 🚀 target here")
    hit = m.document().find("target")[0]
    assert m.range(hit["start"], hit["end"]).Text == "target"


def test_find_replace_after_field_writes_the_right_span():
    m = _Model()
    m.text("Total: ")
    m.field(" = 2+2 ", "4")
    m.text(" units in Q3\r")  # a document always ends with a paragraph mark
    m.document().find_replace("Q3", "Q4")
    ((a, b, value),) = m.writes
    assert value == "Q4"
    assert m.range(a, b).Text == "Q3"
