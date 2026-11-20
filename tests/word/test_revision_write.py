"""Revision write surface — accept/reject, bulk accept/reject, and the
revision-aware read model (`text_final` / `text_original` / `revision_segments`).

The pure reconstruction logic (`segment_runs`) is tested directly; the COM-backed
accept/reject paths round-trip against the `fake_word` MagicMock, whose default
document carries one tracked insertion ("Body", range 13–18, inside the body
paragraph para:2 = "Body text here.").
"""

from __future__ import annotations

from unittest.mock import MagicMock

from office_mcp.word import wordcom
from office_mcp.word.wordcom.review._revisions import segment_runs

# --- segment_runs (pure) -------------------------------------------------------
#
# `text` is the range's Range.Text. Offsets are markup-space, where deleted text
# still occupies positions. Live Word returns deleted runs inside Range.Text when
# markup is shown (the default view) and drops them when it is hidden ("No
# Markup"); segment_runs must give the same segments for both.


def _ins(start, end, text):
    return {"change": "insert", "start": start, "end": end, "text": text}


def _del(start, end, text):
    return {"change": "delete", "start": start, "end": end, "text": text}


def test_segment_runs_no_runs_is_one_unchanged_segment():
    segs = segment_runs("hello world", 0, [])
    assert segs == [{"text": "hello world", "change": None}]


def test_segment_runs_pure_delete_restores_text_not_in_final():
    # Nothing is in final_text, but the deleted run carries its own text.
    segs = segment_runs("", 0, [_del(0, 5, "gone")])
    assert segs == [{"text": "gone", "change": "delete"}]


def test_segment_runs_insert_then_unchanged():
    # "Body " (markup 13–18) inserted, "text here." unchanged.
    segs = segment_runs("Body text here.", 13, [_ins(13, 18, "Body ")])
    assert segs == [
        {"text": "Body ", "change": "insert"},
        {"text": "text here.", "change": None},
    ]


def test_segment_runs_replace_interleaves_insert_and_deleted_text():
    # The canonical tracked find_replace("quick" -> "slow"), as live Word reports
    # it: final_text has "slow"; "quick" survives only on the delete run.
    segs = segment_runs(
        "the slow brown fox",
        0,
        [_ins(4, 8, "slow"), _del(8, 13, "quick")],
    )
    assert segs == [
        {"text": "the ", "change": None},
        {"text": "slow", "change": "insert"},
        {"text": "quick", "change": "delete"},
        {"text": " brown fox", "change": None},
    ]
    final = "".join(s["text"] for s in segs if s["change"] != "delete")
    original = "".join(s["text"] for s in segs if s["change"] != "insert")
    assert final == "the slow brown fox"
    assert original == "the quick brown fox"


def _final(segs):
    return "".join(s["text"] for s in segs if s["change"] != "delete")


def _original(segs):
    return "".join(s["text"] for s in segs if s["change"] != "insert")


def test_segment_runs_markup_view_text_still_holds_deleted_run():
    # Exactly what live Word returns for find_replace("quick" -> "slow") under
    # the default All Markup view: Range.Text keeps "quick" next to "slow".
    segs = segment_runs(
        "the quickslow fox\r",
        0,
        [_del(4, 9, "quick"), _ins(9, 13, "slow")],
    )
    assert segs == [
        {"text": "the ", "change": None},
        {"text": "quick", "change": "delete"},
        {"text": "slow", "change": "insert"},
        {"text": " fox\r", "change": None},
    ]
    assert _final(segs) == "the slow fox\r"
    assert _original(segs) == "the quick fox\r"


def test_segment_runs_no_markup_view_matches_markup_view():
    # Same edit, markup hidden: Range.Text drops "quick". Same segments.
    runs = [_del(4, 9, "quick"), _ins(9, 13, "slow")]
    assert segment_runs("the slow fox\r", 0, runs) == segment_runs("the quickslow fox\r", 0, runs)


def test_segment_runs_markup_view_pure_delete_with_offset_base():
    # Range starts at 100; "very " deleted in the middle, markup shown.
    segs = segment_runs("a very big dog", 100, [_del(102, 107, "very ")])
    assert _final(segs) == "a big dog"
    assert _original(segs) == "a very big dog"


def test_segment_runs_markup_view_deleted_paragraph_mark():
    # A deleted run ending in a paragraph mark: the revision's text is cleaned,
    # but the mark still occupies its markup position in Range.Text.
    segs = segment_runs("one\rtwo\r", 0, [_del(0, 4, "one\r")])
    assert _final(segs) == "two\r"


def test_segment_runs_skips_runs_before_the_cursor():
    # A run starting before base_start (already-consumed space) is ignored.
    segs = segment_runs("xyz", 10, [_ins(5, 8, "??")])
    assert segs == [{"text": "xyz", "change": None}]


def test_segment_runs_ignores_non_text_changes():
    segs = segment_runs("abc", 0, [{"change": "format", "start": 0, "end": 3, "text": ""}])
    assert segs == [{"text": "abc", "change": None}]


# --- revision-aware reads on an anchor -----------------------------------------


def test_text_final_keeps_inserted_runs(fake_word):
    with wordcom.attach() as word:
        para = word.documents.active.paragraphs[2]  # "Body text here." + insert 13–18
        # final = as if accepted: the inserted "Body " stays.
        assert para.text_final.startswith("Body text here.")


def test_text_original_drops_inserted_runs(fake_word):
    with wordcom.attach() as word:
        para = word.documents.active.paragraphs[2]
        # original = as if rejected: the inserted "Body " is gone.
        assert para.text_original.startswith("text here.")
        assert "Body " not in para.text_original


def test_revision_segments_classifies_runs(fake_word):
    with wordcom.attach() as word:
        segs = word.documents.active.paragraphs[2].revision_segments()
    assert segs[0] == {"text": "Body ", "change": "insert"}
    assert segs[1]["change"] is None and segs[1]["text"].startswith("text here.")


# --- accept / reject one --------------------------------------------------------


def test_revision_accept_consumes_it(fake_word):
    with wordcom.attach() as word:
        revs = word.documents.active.revisions
        assert len(revs) == 1
        revs[1].accept()
        assert len(word.documents.active.revisions) == 0


def test_revision_reject_consumes_it(fake_word):
    with wordcom.attach() as word:
        word.documents.active.revisions[1].reject()
        assert len(word.documents.active.revisions) == 0


# --- accept-all / reject-all ----------------------------------------------------


def test_accept_all_returns_count_and_clears(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        assert doc.revisions.accept_all() == 1
        assert len(doc.revisions) == 0


def test_reject_all_empty_is_zero(fake_word):
    fake_word.ActiveDocument.Revisions = _empty_revisions()
    with wordcom.attach() as word:
        assert word.documents.active.revisions.reject_all() == 0


def test_accept_all_within_anchor_scopes_to_its_range(fake_word):
    # Seed the body range (13–29) with its own two tracked changes; accept_all
    # scoped to that anchor resolves exactly those.
    fake_word.ActiveDocument.Range(13, 29).Revisions = _revisions_with(2)
    with wordcom.attach() as word:
        doc = word.documents.active
        within = doc.range(13, 29)
        assert doc.revisions.accept_all(within=within) == 2
        # The whole-document collection is untouched by the scoped accept.
        assert len(doc.revisions) == 1


# --- CLI ------------------------------------------------------------------------


# --- exec ops -------------------------------------------------------------------


# --- helpers --------------------------------------------------------------------


def _empty_revisions():
    coll = MagicMock(name="Revisions")
    coll.Count = 0
    coll.__iter__ = MagicMock(side_effect=lambda: iter([]))
    return coll


def _revisions_with(n: int):
    """A range-scoped Revisions stub: Count + AcceptAll/RejectAll that clear it."""
    state = {"count": n}

    coll = MagicMock(name="RangeRevisions")
    type(coll).Count = property(lambda self: state["count"])

    def _clear() -> None:
        state["count"] = 0

    coll.AcceptAll = MagicMock(side_effect=_clear)
    coll.RejectAll = MagicMock(side_effect=_clear)
    return coll
