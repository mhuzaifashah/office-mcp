"""The shared result builder: compact JSON, truncation that keeps JSON valid, errors."""

from __future__ import annotations

import json

from office_mcp.server import response


def _texts(result):
    return [c.text for c in result.content]


def test_small_json_is_returned_whole_and_compact():
    result = response.json_result({"a": [1, 2], "b": "é"}, max_chars=1000)
    assert _texts(result) == ['{"a":[1,2],"b":"é"}']
    assert not result.is_error


def test_long_list_keeps_whole_items_and_adds_a_note():
    items = [{"id": f"para:{i}", "text": "x" * 50} for i in range(200)]
    result = response.json_result(items, max_chars=2000, more_hint="Pass start and count.")
    body, note = _texts(result)
    shown = json.loads(body)  # still valid JSON
    assert 0 < len(shown) < 200
    assert shown == items[: len(shown)]
    assert len(body) <= 2000
    assert note.startswith("[truncated] Showing items 1-")
    assert f"of {len(items)}" in note
    assert note.endswith("Pass start and count.")


def test_long_dict_shortens_the_biggest_string_field():
    value = {"anchor_id": "heading:3", "text": "word " * 2000}
    result = response.json_result(value, max_chars=1500)
    body, note = _texts(result)
    parsed = json.loads(body)
    assert parsed["anchor_id"] == "heading:3"
    assert parsed["text"].endswith("…")
    assert len(body) <= 1500
    assert "`text`" in note


def test_long_text_is_cut_with_a_note():
    result = response.text_result("a" * 5000, max_chars=1000, more_hint="Use within.", label="html")
    body, note = _texts(result)
    assert body == "a" * 1000
    assert "1,000 of 5,000 characters of html" in note
    assert note.endswith("Use within.")


def test_error_result_is_a_tool_error_with_code_hint_and_details():
    result = response.error_result(
        "anchor_not_found",
        "heading not found: 'heading:9'",
        hint="Re-read outline.",
        details={"ops_run": 1, "failure": None},
    )
    assert result.is_error
    error = json.loads(result.content[0].text)["error"]
    assert error == {
        "code": "anchor_not_found",
        "message": "heading not found: 'heading:9'",
        "retryable": False,
        "hint": "Re-read outline.",
        "ops_run": 1,
    }
