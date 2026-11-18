"""Live smoke test for the WordController facade.

Run:  .venv/Scripts/python.exe -m office_mcp.word.test_word

The vendored `wordcom` test-suite (`tests/`, 1500+ tests, mocked COM) already
covers the operation logic. This script exercises the *live* round-trip through
`WordController` — the part that needs a real Word instance.

Behaviour:
  * Import + construction always run (cross-platform).
  * Live READ checks (status / outline / stats / snapshot) run only if Word is
    actually running with a document open; otherwise they're skipped, not failed.
  * Snapshot is attempted only when PyMuPDF (the `snapshot` extra) is installed.
  * A reversible WRITE round-trip (append → find → snapshot → delete → verify)
    runs only when ``WORDCOM_LIVE_WRITE=1`` is set, because it edits the document
    the user has open right now. It cleans up after itself using the public
    controller API (no reliance on Word's Undo stack), so the document is left
    exactly as it was found.
"""

from __future__ import annotations

import os
import uuid

from office_mcp.word import WordController
from office_mcp.word.wordcom.exceptions import WordcomError, WordNotRunningError


def _run_read_checks(wc: WordController) -> None:
    print("[ok] outline:", wc.outline())
    print("[ok] stats:", wc.read("stats"))
    try:
        shots = wc.snapshot(pages="1", dpi=96)
        print(f"[ok] snapshot: {len(shots)} page(s), first PNG {len(shots[0][1])} bytes")
    except WordcomError as exc:
        # SnapshotError when PyMuPDF (the `snapshot` extra) isn't installed.
        print(f"[skip] snapshot: {exc}")


def _run_write_roundtrip(wc: WordController) -> int:
    """Append a unique marker paragraph, verify it, then delete it again.

    Returns 0 on success, 1 if a step failed. Uses only the public controller
    surface so the edit is reverted whether or not Word's Undo would cooperate.
    """
    marker = f"wordcom-smoke {uuid.uuid4().hex[:8]}"
    print(f"\n[write] round-trip with marker {marker!r}")

    # 1. append a new final paragraph
    res = wc.write("append", text=marker)
    print("[ok] append:", res)

    # 2. find it back
    hits = wc.read("find", text=marker)
    if not hits:
        print("[FAIL] marker not found after append — leaving document untouched")
        return 1
    anchor_id = hits[0]["anchor_id"]
    print(f"[ok] found marker at {anchor_id}")

    # 3. see it (best-effort)
    try:
        shots = wc.snapshot(anchor=anchor_id, dpi=96)
        print(f"[ok] snapshot(anchor): {len(shots)} image(s)")
    except WordcomError as exc:
        print(f"[skip] snapshot(anchor): {exc}")

    # 4. revert — delete the paragraph we added (mark included → closes up)
    del_res = wc.write("delete_paragraph", anchor_id=anchor_id)
    print("[ok] delete_paragraph:", del_res)

    # 5. verify the revert
    still = wc.read("find", text=marker)
    if still:
        print(f"[FAIL] marker still present after delete: {still}")
        return 1
    print("[ok] reverted — marker gone, document restored")
    return 0


def main() -> int:
    wc = WordController()
    print("[ok] WordController constructed (worker:", type(wc._worker).__name__ + ")")

    try:
        status = wc.status()
    except WordNotRunningError:
        print("[skip] Word is not running - open Word with a document for live checks")
        wc.close()
        return 0
    except WordcomError as exc:
        print(f"[skip] live checks unavailable: {exc}")
        wc.close()
        return 0

    rc = 0
    try:
        print("[ok] status:", status)
        _run_read_checks(wc)
        if os.environ.get("WORDCOM_LIVE_WRITE") == "1":
            rc = _run_write_roundtrip(wc)
        else:
            print("\n[skip] write round-trip — set WORDCOM_LIVE_WRITE=1 to exercise edits")
    finally:
        wc.close()

    print("\n[done] live smoke complete")
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
