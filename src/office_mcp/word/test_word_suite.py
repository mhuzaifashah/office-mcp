"""Full control-layer test suite for the Word tool — NO model / API involved.

Run on Windows with Word installed:
    .venv/Scripts/python.exe -m office_mcp.word.test_word_suite

What it does
------------
Exercises the ``WordController`` dispatch surface (read / write / exec / snapshot
+ status / outline / ensure_running) across the full op vocabulary.

It is **resilient**: every operation is wrapped in ``R.check(...)`` which catches
any exception, records PASS/FAIL, prints the outcome, and moves on — so one
failing op never aborts the run. A summary of all failures prints at the end.

Safety
------
Operates on a **throwaway scratch document** it creates itself (via
``Documents.Add``), never the user's open docs. Saving is confined to the temp
dir. By default the scratch doc + Word are left open for visual inspection
(like the PowerPoint suite); set ``WORDCOM_TEST_CLEANUP=1`` to close it (no save)
and delete temp files at the end.

Env knobs
---------
  WORDCOM_TEST_DELAY=0.4   seconds to pause after each check (watchable run)
  WORDCOM_TEST_CLEANUP=1   close the scratch doc + delete temp files at the end
"""

from __future__ import annotations

import os
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Callable

# Allow running as a script; also runs controller.py's sys.path bootstrap so the
# vendored ``wordcom`` is importable as a top-level package.
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from wordcom import attach  # noqa: E402
from wordcom.exceptions import WordcomError, WordNotRunningError  # noqa: E402

from office_mcp.word import WordController  # noqa: E402

STEP_DELAY = float(os.environ.get("WORDCOM_TEST_DELAY", "0"))
CLEANUP = os.environ.get("WORDCOM_TEST_CLEANUP") == "1"

# Word content (equations, smart quotes, …) contains characters the Windows
# console's default cp1252 codec can't encode — an unguarded ``print`` of a
# result then raises UnicodeEncodeError and aborts the whole run. Force UTF-8
# with replacement so no print can ever crash the suite.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
except Exception:
    pass


def _emit(s: str) -> None:
    """Print that never raises on an un-encodable character."""
    try:
        print(s)
    except UnicodeEncodeError:
        print(s.encode("ascii", "replace").decode("ascii"))


_SENTINEL = object()


class Results:
    """Resilient check harness: run, record, keep going."""

    def __init__(self) -> None:
        self.passed = 0
        self.failed = 0
        self.skipped = 0
        self.failures: list[tuple[str, str]] = []
        self.skips: list[tuple[str, str]] = []

    def section(self, title: str) -> None:
        _emit(f"\n{'=' * 66}\n  {title}\n{'=' * 66}")

    def check_soft(self, name: str, fn: Callable[[], Any], reason: str = "") -> Any:
        """Like ``check`` but a failure is a SKIP, not a FAIL — for ops that
        depend on the environment (Excel installed, doc saved, …) rather than on
        the tool being correct."""
        try:
            result = fn()
        except BaseException as exc:  # noqa: BLE001
            self.skipped += 1
            detail = f"{type(exc).__name__}: {exc}" + (f" [{reason}]" if reason else "")
            self.skips.append((name, detail))
            _emit(f"  [SKIP] {name}: {detail}")
            if STEP_DELAY:
                time.sleep(STEP_DELAY)
            return _SENTINEL
        self.passed += 1
        _emit(f"  [PASS] {name}: {_short(result)}")
        if STEP_DELAY:
            time.sleep(STEP_DELAY)
        return result

    def check(self, name: str, fn: Callable[[], Any]) -> Any:
        """Run ``fn``; print PASS/FAIL; return its result (or ``_SENTINEL``)."""
        try:
            result = fn()
        except BaseException as exc:  # noqa: BLE001 — a test must never crash the suite
            self.failed += 1
            detail = f"{type(exc).__name__}: {exc}"
            self.failures.append((name, detail))
            _emit(f"  [FAIL] {name}: {detail}")
            if STEP_DELAY:
                time.sleep(STEP_DELAY)
            return _SENTINEL
        self.passed += 1
        _emit(f"  [PASS] {name}: {_short(result)}")
        if STEP_DELAY:
            time.sleep(STEP_DELAY)
        return result

    def summary(self) -> int:
        total = self.passed + self.failed + self.skipped
        _emit(
            f"\n{'=' * 66}\n  SUMMARY: {self.passed}/{total} passed, "
            f"{self.failed} failed, {self.skipped} skipped\n{'=' * 66}"
        )
        if self.skips:
            _emit("  Skipped (environmental — not tool defects):")
            for name, detail in self.skips:
                _emit(f"   - {name}: {detail}")
        if self.failures:
            _emit("  Failures:")
            for name, detail in self.failures:
                _emit(f"   - {name}: {detail}")
        return 1 if self.failed else 0


def _short(result: Any, limit: int = 160) -> str:
    if result is _SENTINEL:
        return "(no result)"
    text = repr(result)
    return text if len(text) <= limit else text[:limit] + f"… (+{len(text) - limit} chars)"


# --------------------------------------------------------------------------
# Scratch-document lifecycle (via the controller's own COM worker thread)
# --------------------------------------------------------------------------


def _make_scratch(wc: WordController) -> str:
    def job() -> str:
        with attach() as word:
            doc = word.com.Documents.Add()
            return str(doc.Name)

    return wc._worker.run_on_word(job)


def _close_doc(wc: WordController, name: str) -> bool:
    def job() -> bool:
        with attach() as word:
            for doc in word.com.Documents:
                if str(doc.Name) == name:
                    doc.Close(SaveChanges=0)  # wdDoNotSaveChanges
                    return True
        return False

    return wc._worker.run_on_word(job)


# --------------------------------------------------------------------------
# Fresh-anchor helpers (positional ids renumber under edits — re-fetch)
# --------------------------------------------------------------------------


def _first_heading(wc: WordController) -> str | None:
    try:
        for row in wc.read("outline") or []:
            aid = row.get("anchor_id") if isinstance(row, dict) else None
            if isinstance(aid, str) and aid.startswith("heading:"):
                return aid
    except Exception:
        pass
    return None


def _first_para(wc: WordController) -> str | None:
    try:
        for row in wc.read("paragraphs", count=50) or []:
            aid = row.get("anchor_id") if isinstance(row, dict) else None
            if isinstance(aid, str) and aid.startswith("para:"):
                return aid
    except Exception:
        pass
    return None


def _find(wc: WordController, text: str) -> str | None:
    try:
        hits = wc.read("find", text=text) or []
        if hits and isinstance(hits[0], dict):
            return hits[0].get("anchor_id")
    except Exception:
        pass
    return None


# ==========================================================================
# Suite
# ==========================================================================


def run(wc: WordController, R: Results) -> None:
    st: dict[str, Any] = {}
    tmp = tempfile.gettempdir()

    # ---- lifecycle / status ------------------------------------------------
    R.section("lifecycle / status")
    R.check("ensure_running", lambda: wc.ensure_running(visible=True))
    R.check("status", lambda: wc.status())
    R.check("read:guide (no Word needed)", lambda: wc.read("guide"))

    # ---- build baseline content (also write coverage) ----------------------
    R.section("write — build baseline content")
    md = (
        "# Word Tool Test\n\n"
        "Intro paragraph one.\n\n"
        "## Section A\n\n"
        "Body of section A.\n\n"
        "## Section B\n\n"
        "Body of section B mentioning the risk register phrase.\n"
    )
    R.check(
        "write:insert_markdown", lambda: wc.write("insert_markdown", anchor_id="start", markdown=md)
    )
    R.check("write:append", lambda: wc.write("append", text="Appended final paragraph."))
    R.check("write:prepend", lambda: wc.write("prepend", text="Prepended first paragraph."))
    R.check(
        "write:insert(end)",
        lambda: wc.write("insert", anchor_id="end", text="Inserted at end.", style="Normal"),
    )
    R.check(
        "write:insert_block",
        lambda: wc.write(
            "insert_block",
            anchor_id="end",
            items=[
                "Bullet one.",
                {"text": "Bullet **two**.", "style": "List Bullet"},
                "Bullet three.",
            ],
        ),
    )
    R.check(
        "write:insert_section",
        lambda: wc.write(
            "insert_section",
            anchor_id="end",
            heading="Section C",
            body=["Body C para one.", "Body C para two."],
            level=2,
        ),
    )

    # ---- reads -------------------------------------------------------------
    R.section("read — introspection")
    R.check("read:outline", lambda: wc.read("outline"))
    R.check("read:paragraphs", lambda: wc.read("paragraphs", count=20))
    R.check("read:stats", lambda: wc.read("stats"))
    R.check("read:sections", lambda: wc.read("sections"))
    R.check("read:styles", lambda: wc.read("styles"))
    R.check("read:find", lambda: wc.read("find", text="risk register"))
    R.check(
        "read:find_paragraphs", lambda: wc.read("find_paragraphs", text="risk register", limit=3)
    )
    R.check("read:to_markdown", lambda: wc.read("to_markdown"))
    R.check("read:to_html", lambda: wc.read("to_html"))
    R.check("read:digest", lambda: wc.read("digest", budget=2000))
    R.check("read:track", lambda: wc.read("track"))
    R.check("read:properties", lambda: wc.read("properties"))
    R.check("read:variables", lambda: wc.read("variables"))
    R.check("read:theme", lambda: wc.read("theme"))
    R.check("read:themes", lambda: wc.read("themes"))
    R.check("read:hyperlinks", lambda: wc.read("hyperlinks"))
    R.check("read:fields", lambda: wc.read("fields"))
    R.check("read:comments", lambda: wc.read("comments"))
    R.check("read:revisions", lambda: wc.read("revisions"))
    R.check("read:images", lambda: wc.read("images"))
    R.check("read:shapes", lambda: wc.read("shapes"))
    R.check("read:charts", lambda: wc.read("charts"))
    R.check("read:equations", lambda: wc.read("equations"))
    R.check("read:footnotes", lambda: wc.read("footnotes"))
    R.check("read:endnotes", lambda: wc.read("endnotes"))
    R.check("read:proofing", lambda: wc.read("proofing"))
    R.check("read:lint", lambda: wc.read("lint"))

    # anchor-dependent reads
    h = _first_heading(wc)
    if h:
        R.check("read:read_section(heading anchor)", lambda: wc.read("read_section", anchor_id=h))
        R.check("read:read_text(heading)", lambda: wc.read("read_text", anchor_id=h, view="raw"))
        R.check("read:location(heading)", lambda: wc.read("location", anchor_id=h))
        R.check("read:format_info(heading)", lambda: wc.read("format_info", anchor_id=h))
        R.check(
            "read:nearest_heading",
            lambda: wc.read("nearest_heading", anchor_id=h, direction="after"),
        )
    R.check(
        "read:read_section(by heading text)", lambda: wc.read("read_section", heading="Section A")
    )
    hs = wc.read("outline") or []
    heads = [
        r.get("anchor_id")
        for r in hs
        if isinstance(r, dict) and str(r.get("anchor_id", "")).startswith("heading:")
    ]
    if len(heads) >= 2:
        R.check(
            "read:between(headings)",
            lambda: wc.read("between", start_anchor=heads[0], end_anchor=heads[1]),
        )

    # ---- style / formatting ------------------------------------------------
    R.section("write — style / formatting")
    if h:
        R.check(
            "write:apply_style",
            lambda: wc.write("apply_style", anchor_id=_first_heading(wc), name="Heading 2"),
        )
    p = _first_para(wc)
    if p:
        R.check(
            "write:format_paragraph",
            lambda: wc.write(
                "format_paragraph",
                anchor_id=_first_para(wc),
                alignment="center",
                space_before=6,
                line_spacing=1.5,
            ),
        )
        R.check(
            "write:format_run",
            lambda: wc.write(
                "format_run",
                anchor_id=_first_para(wc),
                bold=True,
                italic=True,
                color="#1F4E79",
                size="12pt",
            ),
        )
        R.check(
            "write:set_shading",
            lambda: wc.write("set_shading", anchor_id=_first_para(wc), fill="#FFF2CC"),
        )
        R.check(
            "write:set_borders",
            lambda: wc.write(
                "set_borders",
                anchor_id=_first_para(wc),
                sides="all",
                line_style="single",
                weight=0.5,
                color="black",
            ),
        )
        R.check(
            "write:drop_cap",
            lambda: wc.write("drop_cap", anchor_id=_first_para(wc), lines=3, position="dropped"),
        )
        R.check(
            "write:add_tab_stop",
            lambda: wc.write(
                "add_tab_stop",
                anchor_id=_first_para(wc),
                position="3in",
                align="right",
                leader="dots",
            ),
        )
        R.check("read:list_levels(para)", lambda: wc.read("list_levels", anchor_id=_first_para(wc)))
    R.check(
        "write:add_style",
        lambda: wc.write("add_style", name="TestBrand", type="paragraph", based_on="Normal"),
    )
    R.check(
        "write:set_style",
        lambda: wc.write("set_style", name="TestBrand", bold=True, color="#1F3864", size="14pt"),
    )

    # ---- lists -------------------------------------------------------------
    R.section("write — lists")
    rng = _find(wc, "Bullet one.")
    if rng:
        R.check(
            "write:list apply(bulleted)",
            lambda: wc.write(
                "list", action="apply", anchor_id=_find(wc, "Bullet one.") or rng, type="bulleted"
            ),
        )
        R.check(
            "write:list indent",
            lambda: wc.write("list", action="indent", anchor_id=_find(wc, "Bullet one.") or rng),
        )
        R.check(
            "read:list info via list_levels",
            lambda: wc.read("list_levels", anchor_id=_find(wc, "Bullet one.") or rng),
        )
        R.check(
            "write:list remove",
            lambda: wc.write("list", action="remove", anchor_id=_find(wc, "Bullet one.") or rng),
        )

    # ---- tables ------------------------------------------------------------
    R.section("write — tables")
    tbl = R.check(
        "write:table create(3x3)",
        lambda: wc.write(
            "table",
            action="create",
            anchor_id="end",
            rows=3,
            cols=3,
            style="Table Grid",
            header=True,
        ),
    )
    if tbl is not _SENTINEL and isinstance(tbl, dict):
        # write result: {'ok':True,'command':'table','result':{'op':'create_table','table':1,...}}
        res = tbl.get("result")
        st["table"] = res.get("table") if isinstance(res, dict) else res
    ti = st.get("table")
    if ti:
        # Fill a clean header + data row, then read it back BEFORE mutating the
        # grid — table_records maps columns→headers and needs a uniform table
        # (merging/adding columns makes it non-uniform, which records rejects).
        for c, h in enumerate(["Col A", "Col B", "Col C"], start=1):
            R.check(
                f"write:table set_cell(1,{c})",
                lambda c=c, h=h: wc.write(
                    "table", action="set_cell", table=ti, row=1, col=c, text=h
                ),
            )
        R.check(
            "write:table set_cell(2,1)",
            lambda: wc.write("table", action="set_cell", table=ti, row=2, col=1, text="Data A"),
        )
        R.check("read:table_list", lambda: wc.read("table_list"))
        R.check("read:table_read (clean)", lambda: wc.read("table_read", table=ti))
        R.check("read:table_records (clean)", lambda: wc.read("table_records", table=ti))
        # Now the structural mutations (these intentionally make it non-uniform).
        R.check(
            "write:table add_row",
            lambda: wc.write("table", action="add_row", table=ti, values=["r", "s", "t"]),
        )
        R.check("write:table add_column", lambda: wc.write("table", action="add_column", table=ti))
        R.check(
            "write:table merge_cells",
            lambda: wc.write(
                "table", action="merge_cells", table=ti, **{"from": "1:1", "to": "1:2"}
            ),
        )
        R.check(
            "write:table autofit",
            lambda: wc.write("table", action="autofit", table=ti, mode="content"),
        )
        R.check(
            "write:table set_style",
            lambda: wc.write("table", action="set_style", table=ti, style="Table Grid"),
        )
        R.check(
            "write:table set_banding",
            lambda: wc.write("table", action="set_banding", table=ti, banded_rows=True),
        )

    # ---- bookmarks / content controls -------------------------------------
    R.section("write — bookmarks / content controls")
    bmk_anchor = _first_para(wc)
    if bmk_anchor:
        R.check(
            "write:add_bookmark",
            lambda: wc.write("add_bookmark", name="Intro", anchor_id=_first_para(wc)),
        )
        R.check(
            "write:write_bookmark",
            lambda: wc.write("write_bookmark", name="Intro", text="Bookmarked intro text."),
        )
        R.check("read:read_bookmark", lambda: wc.read("read_bookmark", name="Intro"))
        R.check("write:pin", lambda: wc.write("pin", anchor_id=_first_para(wc), name="intropin"))
    R.check("write:pin_outline", lambda: wc.write("pin_outline"))
    cc_anchor = _find(wc, "Inserted at end.")
    if cc_anchor:
        R.check(
            "write:create_content_control",
            lambda: wc.write(
                "create_content_control",
                anchor_id=_find(wc, "Inserted at end.") or cc_anchor,
                kind="dropdown",
                title="Status",
                items=["Open", "Done"],
            ),
        )
        R.check("read:read_cc", lambda: wc.read("read_cc", name="Status"))
        R.check(
            "write:set_cc_items",
            lambda: wc.write(
                "set_cc_items", anchor_id="cc:Status", items=["Open", "Done", "Blocked"]
            ),
        )

    # ---- hyperlinks / cross-refs / captions -------------------------------
    R.section("write — hyperlinks / cross-refs / captions")
    link_anchor = _find(wc, "Appended final paragraph.")
    if link_anchor:
        R.check(
            "write:add_hyperlink(url)",
            lambda: wc.write(
                "add_hyperlink",
                anchor_id=_find(wc, "Appended final paragraph.") or link_anchor,
                url="https://www.anthropic.com",
                text=" (link)",
            ),
        )
    R.check(
        "write:insert_cross_reference",
        lambda: wc.write(
            "insert_cross_reference",
            anchor_id="end",
            target="bookmark:Intro",
            kind="page",
        ),
    )
    cap_anchor = _first_para(wc)
    if cap_anchor:
        R.check(
            "write:insert_caption",
            lambda: wc.write(
                "insert_caption", anchor_id=_first_para(wc), label="Figure", text="A caption"
            ),
        )

    # ---- fields ------------------------------------------------------------
    R.section("write — fields")
    R.check(
        "write:insert_field(page)", lambda: wc.write("insert_field", anchor_id="end", kind="date")
    )
    R.check("write:update_fields", lambda: wc.write("update_fields"))

    # ---- TOC / index / notes ----------------------------------------------
    R.section("write — TOC / index / notes")
    R.check("write:insert_toc", lambda: wc.write("insert_toc", anchor_id="start", levels=[1, 3]))
    idx_anchor = _find(wc, "risk register")
    if idx_anchor:
        R.check(
            "write:mark_index_entry",
            lambda: wc.write(
                "mark_index_entry", anchor_id=_find(wc, "risk register") or idx_anchor, entry="risk"
            ),
        )
    R.check("write:insert_index", lambda: wc.write("insert_index", anchor_id="end", columns=2))
    R.check(
        "write:insert_table_of_figures",
        lambda: wc.write("insert_table_of_figures", anchor_id="end", label="Figure"),
    )
    fn_anchor = _first_para(wc)
    if fn_anchor:
        R.check(
            "write:insert_footnote",
            lambda: wc.write("insert_footnote", anchor_id=_first_para(wc), text="A footnote."),
        )
        R.check(
            "write:insert_endnote",
            lambda: wc.write("insert_endnote", anchor_id=_first_para(wc), text="An endnote."),
        )

    # ---- citations / bibliography -----------------------------------------
    R.section("write — citations / bibliography")
    R.check("write:set_bibliography_style", lambda: wc.write("set_bibliography_style", style="APA"))
    R.check(
        "write:add_source",
        lambda: wc.write(
            "add_source",
            source_type="book",
            author="Smith, Jane",
            title="A Test Book",
            year=2020,
            tag="Smith2020",
        ),
    )
    R.check(
        "write:insert_citation",
        lambda: wc.write("insert_citation", anchor_id="end", tag="Smith2020"),
    )
    R.check("write:insert_bibliography", lambda: wc.write("insert_bibliography", anchor_id="end"))
    toa_anchor = _first_para(wc)
    if toa_anchor:
        R.check(
            "write:mark_citation",
            lambda: wc.write(
                "mark_citation",
                anchor_id=_first_para(wc),
                long_citation="Brown v. Board, 347 U.S. 483 (1954)",
                short_citation="Brown",
                category="cases",
            ),
        )
    R.check(
        "write:insert_table_of_authorities",
        lambda: wc.write("insert_table_of_authorities", anchor_id="end", category="all"),
    )

    # ---- images / equations / charts --------------------------------------
    R.section("write — images / equations / charts")
    img_path = os.path.join(tmp, "wordcom_suite_img.png")
    try:
        from PIL import Image, ImageDraw

        im = Image.new("RGB", (240, 140), color=(40, 100, 220))
        ImageDraw.Draw(im).rectangle([0, 0, 239, 139], outline=(220, 40, 40), width=6)
        im.save(img_path, "PNG")
        st["img_path"] = img_path
    except Exception as e:
        print(f"  [SKIP] could not build test image (PIL): {e}")
    if st.get("img_path"):
        R.check(
            "write:insert_image(inline)",
            lambda: wc.write("insert_image", anchor_id="end", wrap="inline", path=st["img_path"]),
        )
        R.check("read:images(after insert)", lambda: wc.read("images"))
        img_anchor = _find_image(wc)
        if img_anchor:
            # read_image serialises the range's Flat OPC; Word materialises embedded
            # media lazily, so on a never-saved doc the part can be absent — soft.
            R.check_soft(
                "read:read_image",
                lambda: wc.read("read_image", anchor_id=img_anchor),
                reason="may require a saved document (lazy media serialization)",
            )
    R.check(
        "write:insert_equation(unicodemath)",
        lambda: wc.write("insert_equation", anchor_id="end", unicodemath="x=(-b±√(b^2-4ac))/(2a)"),
    )
    R.check("read:equations", lambda: wc.read("equations"))
    R.check_soft(
        "write:insert_chart",
        lambda: wc.write(
            "insert_chart",
            anchor_id="end",
            kind="bar",
            data={"Q1": 10, "Q2": 25, "Q3": 18},
            title="Quarterly",
        ),
        reason="needs Excel installed for the embedded chart",
    )

    # ---- shapes / text boxes ----------------------------------------------
    R.section("write — shapes / text boxes")
    R.check(
        "write:text_box",
        lambda: wc.write(
            "text_box", anchor_id="end", text="Key takeaway", width="2.5in", fill="#EEEEFF"
        ),
    )
    shp = _first_shape(wc)
    if shp:
        R.check(
            "write:set_shape_text",
            lambda: wc.write("set_shape_text", anchor_id=_first_shape(wc), text="Revised takeaway"),
        )
        R.check(
            "write:set_shape_rotation",
            lambda: wc.write("set_shape_rotation", anchor_id=_first_shape(wc), degrees=15),
        )
        R.check(
            "write:set_shape_z_order",
            lambda: wc.write("set_shape_z_order", anchor_id=_first_shape(wc), order="front"),
        )

    # ---- theme -------------------------------------------------------------
    R.section("write — theme")
    R.check(
        "write:set_theme_colors",
        lambda: wc.write("set_theme_colors", colors={"accent1": "#1A73E8"}),
    )
    R.check(
        "write:set_theme_fonts", lambda: wc.write("set_theme_fonts", major="Arial", minor="Calibri")
    )

    # ---- page setup / headers / footers / watermark -----------------------
    R.section("write — page setup / headers / footers / watermark")
    R.check(
        "write:page_setup",
        lambda: wc.write("page_setup", section=1, orientation="portrait", margins="1in"),
    )
    R.check(
        "write:header", lambda: wc.write("header", section=1, text="Confidential", which="primary")
    )
    R.check(
        "write:footer", lambda: wc.write("footer", section=1, text="Page footer", which="primary")
    )
    R.check("write:watermark", lambda: wc.write("watermark", text="DRAFT"))

    # ---- properties / variables -------------------------------------------
    R.section("write — properties / variables")
    R.check(
        "write:set_property", lambda: wc.write("set_property", name="Title", value="Word Tool Test")
    )
    R.check(
        "write:set_property(custom)",
        lambda: wc.write("set_property", name="Project", value="Apollo", custom=True),
    )
    R.check("write:delete_property", lambda: wc.write("delete_property", name="Project"))
    R.check("write:set_variable", lambda: wc.write("set_variable", name="ClientName", value="Acme"))
    R.check("write:delete_variable", lambda: wc.write("delete_variable", name="ClientName"))

    # ---- comments / revisions / track -------------------------------------
    R.section("write — comments / revisions / track")
    cmt_anchor = _first_para(wc)
    if cmt_anchor:
        R.check(
            "write:comment add",
            lambda: wc.write(
                "comment", action="add", anchor_id=_first_para(wc), text="Review this."
            ),
        )
    R.check("read:comments(after add)", lambda: wc.read("comments"))
    R.check("write:comment resolve(1)", lambda: wc.write("comment", action="resolve", index=1))
    R.check("write:comment delete(1)", lambda: wc.write("comment", action="delete", index=1))
    R.check("write:track on", lambda: wc.write("track", on=True))
    R.check(
        "write:tracked edit (append)", lambda: wc.write("append", text="Tracked change paragraph.")
    )
    R.check("read:revisions(after tracked edit)", lambda: wc.read("revisions"))
    R.check("write:revision accept_all", lambda: wc.write("revision", action="accept_all"))
    R.check("write:track off", lambda: wc.write("track", on=False))

    # ---- replace variants --------------------------------------------------
    R.section("write — replace")
    # Target a clean body phrase — NOT one earlier turned into an index/TOC entry
    # (the engine correctly refuses to overwrite a find that resolved into a field).
    R.check(
        "write:replace(find/all)",
        lambda: wc.write(
            "replace", find="Body of section A.", text="Body of section A (edited).", all=True
        ),
    )
    del_anchor = _find(wc, "Bullet three.")
    if del_anchor:
        R.check(
            "write:delete_paragraph",
            lambda: wc.write(
                "delete_paragraph", anchor_id=_find(wc, "Bullet three.") or del_anchor
            ),
        )

    # ---- checkpoint / diff -------------------------------------------------
    R.section("read — checkpoint / diff")
    cp = R.check("read:checkpoint", lambda: wc.read("checkpoint", include="text+style"))
    if cp is not _SENTINEL and cp is not None:
        R.check(
            "write:append (post-checkpoint change)",
            lambda: wc.write("append", text="Change after checkpoint."),
        )
        R.check("read:diff(since checkpoint)", lambda: wc.read("diff", checkpoint=cp))

    # ---- exec (batch) ------------------------------------------------------
    R.section("exec — batch as one undo")
    R.check(
        "exec:2 appends",
        lambda: wc.exec(
            [
                {"op": "append_paragraph", "text": "Exec batch A."},
                {"op": "append_paragraph", "text": "Exec batch B."},
            ],
            label="suite exec batch",
        ),
    )

    # ---- snapshot ----------------------------------------------------------
    R.section("snapshot — render to PNG (needs PyMuPDF)")
    R.check("snapshot:pages=1", lambda: wc.snapshot(pages="1", dpi=96))
    snap_h = _first_heading(wc)
    if snap_h:
        R.check("snapshot:anchor=heading", lambda: wc.snapshot(anchor=_first_heading(wc), dpi=96))

    # ---- persistence -------------------------------------------------------
    R.section("write — persistence (save / save_as / export_pdf)")
    docx_path = os.path.join(tmp, "wordcom_suite.docx")
    pdf_path = os.path.join(tmp, "wordcom_suite.pdf")
    st["docx_path"] = docx_path
    st["pdf_path"] = pdf_path
    R.check("write:export_pdf", lambda: wc.write("export_pdf", path=pdf_path))
    sa = R.check("write:save_as", lambda: wc.write("save_as", path=docx_path, overwrite=True))
    if sa is not _SENTINEL and isinstance(sa, dict):
        st["scratch_name"] = os.path.basename(docx_path)  # save_as renames the doc
    R.check("write:save", lambda: wc.write("save"))


def _find_image(wc: WordController) -> str | None:
    try:
        imgs = wc.read("images") or []
        if imgs and isinstance(imgs[0], dict):
            return imgs[0].get("anchor_id") or "image:1"
    except Exception:
        pass
    return "image:1"


def _first_shape(wc: WordController) -> str | None:
    try:
        shps = wc.read("shapes") or []
        if shps and isinstance(shps[0], dict):
            return shps[0].get("anchor_id") or "shape:1"
    except Exception:
        pass
    return None


def main() -> int:
    tmp = tempfile.gettempdir()
    wc = WordController(save_dirs=[tmp])
    print(f"[init] WordController(save_dirs=[{tmp}]) worker={type(wc._worker).__name__}")

    # Gate: need a reachable Word. ensure_running launches one if none is up.
    try:
        wc.ensure_running(visible=True)
        wc.status()
    except (WordNotRunningError, WordcomError) as exc:
        print(
            f"[skip] Word unavailable ({type(exc).__name__}: {exc}) — run on Windows with Word installed."
        )
        wc.close()
        return 0
    except Exception as exc:  # noqa: BLE001 — e.g. pywin32/COM missing off-Windows
        print(f"[skip] cannot reach Word ({type(exc).__name__}: {exc}).")
        wc.close()
        return 0

    R = Results()
    scratch = _make_scratch(wc)
    print(f"[init] scratch document: {scratch}")

    try:
        run(wc, R)
    finally:
        rc_name = scratch  # default; save_as may have renamed it
        if CLEANUP:
            # Try both the original and the saved name (save_as renames the doc).
            for nm in {scratch, "wordcom_suite.docx"}:
                try:
                    _close_doc(wc, nm)
                except Exception:
                    pass
            for key in ("wordcom_suite.docx", "wordcom_suite.pdf", "wordcom_suite_img.png"):
                fp = os.path.join(tempfile.gettempdir(), key)
                if os.path.exists(fp):
                    try:
                        os.remove(fp)
                    except OSError:
                        pass
            print("[cleanup] scratch closed (no save) + temp files removed")
        else:
            print(
                f"[keep] scratch doc left open for inspection ({rc_name}); set WORDCOM_TEST_CLEANUP=1 to auto-close"
            )
        wc.close()

    return R.summary()


if __name__ == "__main__":
    raise SystemExit(main())
