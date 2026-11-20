"""Time formatting ops against a real Word (its own instance, discarded after).

Builds a long document (headings + body paragraphs), or opens a copy of a real
one with `--file`, then runs exec batches of `apply_style`, `format_paragraph`
and `format_run` and prints milliseconds per op. `--profile` prints the hottest
functions. Never point `--file` at an original: the bench edits the document
(it closes without saving, but use a copy anyway).

    uv run python scripts/bench_word_format.py [--file COPY.docx] [--visible] [--profile]
"""

from __future__ import annotations

import argparse
import cProfile
import pstats
import time

import win32com.client

from office_mcp.word.wordcom.model._application import Word
from office_mcp.word.wordcom.runtime._ops import run_batch

BODY = (
    "The committee reviewed the proposal in detail and agreed that the budget, "
    "the schedule and the staffing plan all needed further work before approval. "
)


def build_doc(app, paras: int) -> None:
    doc_com = app.Documents.Add()
    lines = [f"Section {i // 10 + 1}" if i % 10 == 0 else f"{i}: {BODY * 2}" for i in range(paras)]
    doc_com.Content.Text = "\r".join(lines)
    for i in range(1, paras + 1, 10):
        doc_com.Paragraphs(i).Style = -2  # wdStyleHeading1


def make_ops(doc, n: int, word: str) -> list[dict]:
    range_ids = [m["anchor_id"] for m in doc.find(word)][:n]
    total = int(doc.com.Paragraphs.Count)
    ops: list[dict] = []
    for i in range(n):
        kind = i % 3
        para = 2 + (i * 37) % (total - 2)
        if kind == 0:
            ops.append({"op": "apply_style", "anchor_id": f"para:{para}", "name": "Quote"})
        elif kind == 1:
            ops.append(
                {
                    "op": "format_paragraph",
                    "anchor_id": f"para:{para}",
                    "alignment": "justify",
                    "space_after": 6,
                    "line_spacing": 1.15,
                }
            )
        else:
            target = range_ids[i % len(range_ids)] if range_ids else f"para:{para}"
            ops.append({"op": "format_run", "anchor_id": target, "bold": True, "color": "#1F4E79"})
    return ops


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--paras", type=int, default=400)
    ap.add_argument("--ops", type=int, default=90)
    ap.add_argument("--file", help="open a COPY of a real .docx instead of building one")
    ap.add_argument("--word", default="the", help="word to find for format_run targets")
    ap.add_argument("--visible", action="store_true", help="show Word, as real users run it")
    ap.add_argument("--profile", action="store_true")
    args = ap.parse_args()

    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = args.visible
    try:
        if args.file:
            app.Documents.Open(args.file, False, False, False)
        else:
            build_doc(app, args.paras)
        doc = Word(app).documents.active
        print(
            f"pages: {app.ActiveDocument.ComputeStatistics(2)}  "
            f"paragraphs: {app.ActiveDocument.Paragraphs.Count}"
        )
        ops = make_ops(doc, args.ops, args.word)
        prof = cProfile.Profile() if args.profile else None
        t0 = time.perf_counter()
        if prof:
            prof.enable()
        result, exc = run_batch(doc, ops, label="bench")
        if prof:
            prof.disable()
        dt = time.perf_counter() - t0
        print(f"ok={result['ok']} ops_run={result['ops_run']} error={exc}")
        print(f"{len(ops)} ops in {dt:.2f}s = {dt / len(ops) * 1000:.0f} ms/op")
        for kind in ("apply_style", "format_paragraph", "format_run"):
            sub = [o for o in ops if o["op"] == kind]
            t0 = time.perf_counter()
            run_batch(doc, sub, label=kind)
            d = time.perf_counter() - t0
            print(f"  {kind:17s} {len(sub):3d} ops {d:6.2f}s  {d / len(sub) * 1000:5.0f} ms/op")
        if prof:
            pstats.Stats(prof).sort_stats("cumulative").print_stats(30)
        app.ActiveDocument.Close(False)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
