"""A new paragraph without a style is Normal, not the anchor's style (W15)."""

from __future__ import annotations

import wordcom

# Bookmark "Address" spans (13, 24) in the fake document.
TEXT = "body line"


def _rng(fake_word, start, end):
    return fake_word.ActiveDocument.Range(start, end)


def _style(fake_word, name):
    return fake_word.ActiveDocument.Styles(name)


def _insert(method, **kw):
    with wordcom.attach() as word:
        getattr(word.documents.active.bookmarks["Address"], method)(TEXT, **kw)


def test_unstyled_paragraph_after_gets_normal(fake_word):
    _insert("insert_paragraph_after")
    assert _rng(fake_word, 24, 24 + len(TEXT)).Style == _style(fake_word, "Normal")


def test_unstyled_paragraph_before_gets_normal(fake_word):
    _insert("insert_paragraph_before")
    assert _rng(fake_word, 13, 13 + len(TEXT)).Style == _style(fake_word, "Normal")


def test_explicit_style_still_wins(fake_word):
    _insert("insert_paragraph_after", style="Body Text")
    assert _rng(fake_word, 24, 24 + len(TEXT)).Style == _style(fake_word, "Body Text")


def test_normal_resolved_by_builtin_id_on_localized_word(fake_word):
    # German Word: the Normal style is named "Standard"; only id -1 finds it.
    styles = fake_word.ActiveDocument.Styles
    normal = styles("Normal")
    normal.NameLocal = "Standard"
    original_call = type(styles).__call__
    type(styles).__call__ = lambda self, key: (
        normal if key in (-1, "Standard") else original_call(self, key)
    )
    try:
        _insert("insert_paragraph_after")
    finally:
        type(styles).__call__ = original_call
    assert _rng(fake_word, 24, 24 + len(TEXT)).Style == normal


def test_block_items_without_style_get_normal(fake_word):
    with wordcom.attach() as word:
        word.documents.active.bookmarks["Address"].insert_block(["a", "bb"])
    normal = _style(fake_word, "Normal")
    # Paragraph "a" spans 24-25; "bb" starts after the CR at 26.
    assert _rng(fake_word, 24, 25).Paragraphs(1).Range.Style == normal
    assert _rng(fake_word, 26, 28).Paragraphs(1).Range.Style == normal
