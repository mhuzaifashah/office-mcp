"""
Smoke test for PowerPointController — run on Windows with PowerPoint installed.

Usage:
    .venv/Scripts/python.exe -m office_mcp.ppt.test_ppt

Creates a temp presentation, exercises Stage 1 ops (file + slides),
then cleans up.
"""

from __future__ import annotations

import os
import tempfile
import time

from office_mcp.ppt.controller import PowerPointController, PowerPointError

# Pause between each ok() so a human can watch PowerPoint react in real time.
# Default 0 (fast). Only enabled in the newest test sections via _enable_pause().
# Override globally with env: PPT_TEST_DELAY=2.5
STEP_DELAY = float(os.environ.get("PPT_TEST_DELAY", "1.0"))
_PAUSE = False


def _enable_pause() -> None:
    global _PAUSE
    _PAUSE = True


def section(title: str) -> None:
    print(f"\n{'=' * 60}")
    print(f"  {title}")
    print(f"{'=' * 60}")


def ok(msg: str, result=None) -> None:
    print(f"  [OK] {msg}")
    if isinstance(result, dict):
        for k, v in result.items():
            print(f"        {k}: {v}")
    elif isinstance(result, list):
        print(f"        ({len(result)} items)")
        for item in result[:5]:
            print(f"        - {item}")
    if _PAUSE and STEP_DELAY > 0:
        time.sleep(STEP_DELAY)


def run_tests() -> None:
    tmp_path = os.path.join(tempfile.gettempdir(), "ppt_tool_test.pptx")
    save_as_pdf = os.path.join(tempfile.gettempdir(), "ppt_tool_test.pdf")

    # Clean any leftover from prior runs
    for p in (tmp_path, save_as_pdf):
        if os.path.exists(p):
            try:
                os.remove(p)
            except OSError:
                pass

    ctrl = PowerPointController(visible=True)

    try:
        # ── APP INFO ──
        section("App Info")
        ok("app_info", ctrl.app_info())

        # ── FILE OPS ──
        section("File / Session")
        ok("test_file (nonexistent)", ctrl.test_file(tmp_path))
        ok("create blank 16:9 + saveAs", ctrl.create(tmp_path, preset="16:9"))
        ok("list_open_presentations", ctrl.list_open_presentations())
        ok("save", ctrl.save())

        # ── SLIDE OPS ──
        section("Slides — add / list")
        ok("slide_add (default blank @ end)", ctrl.slide_add())
        ok("slide_add (title layout)", ctrl.slide_add(layout="title"))
        ok(
            "slide_add (section_header @ pos 2)",
            ctrl.slide_add(position=2, layout="section_header"),
        )
        ok("slide_list", ctrl.slide_list())

        section("Slides — duplicate / move / goto")
        dup = ctrl.slide_duplicate(1)
        ok("slide_duplicate(1)", dup)
        ok("slide_move(last → 1)", ctrl.slide_move(ctrl.pres.Slides.Count, 1))
        ok("slide_goto(2)", ctrl.slide_goto(2))

        section("Slides — layout / hidden / notes")
        ok("slide_set_layout(1, 'title_only')", ctrl.slide_set_layout(1, "title_only"))
        ok("slide_set_hidden(1, True)", ctrl.slide_set_hidden(1, True))
        ok("slide_set_notes(1, 'hello notes')", ctrl.slide_set_notes(1, "hello notes"))
        ok("slide_get_notes(1)", ctrl.slide_get_notes(1))

        section("Slides — delete")
        before = ctrl.pres.Slides.Count
        ctrl.slide_delete(before)
        ok(f"slide_delete({before}) → remaining {ctrl.pres.Slides.Count}")

        # ── SHAPES (Stage 2a) ──
        section("Shapes — list / add / mutate / delete")
        # Pick the blank slide we appended at the end of Stage 1 tests.
        target = 1
        ok("shape_list(before)", ctrl.shape_list(target))

        rect = ctrl.shape_add(target, "rectangle", 50, 50, 200, 100, text="hello rect")
        ok("shape_add rectangle", rect)
        rect_id = rect["shape_id"]

        oval = ctrl.shape_add(target, "oval", 300, 50, 150, 150)
        ok("shape_add oval", oval)

        tb = ctrl.shape_add_text_box(target, 50, 200, 400, 60, text="text box body")
        ok("shape_add_text_box", tb)
        tb_id = tb["shape_id"]

        ok("shape_move(rect → 400,250)", ctrl.shape_move(target, rect_id, 400, 250))
        ok("shape_resize(rect → 250,150)", ctrl.shape_resize(target, rect_id, 250, 150))
        ok("shape_list(after add+mutate)", ctrl.shape_list(target))

        ok("shape_delete(text_box)", ctrl.shape_delete(target, tb_id))
        ok("shape_list(after delete)", ctrl.shape_list(target))

        # ── SHAPES (Stage 2b) ──
        section("Shapes — fill / line / rotate")
        oval_id = oval["shape_id"]
        ok("set_fill_color(rect, #FF6600)", ctrl.shape_set_fill_color(target, rect_id, "#FF6600"))
        ok(
            "set_fill_transparency(rect, 0.4)",
            ctrl.shape_set_fill_transparency(target, rect_id, 0.4),
        )
        ok(
            "set_line_color(rect, [0,128,255])",
            ctrl.shape_set_line_color(target, rect_id, [0, 128, 255]),
        )
        ok("set_line_weight(rect, 4)", ctrl.shape_set_line_weight(target, rect_id, 4))
        ok(
            "set_line_dash_style(rect, dash_dot)",
            ctrl.shape_set_line_dash_style(target, rect_id, "dash_dot"),
        )
        ok("set_rotation(rect, 30)", ctrl.shape_set_rotation(target, rect_id, 30))
        ok("set_fill_color(oval, #00AA66)", ctrl.shape_set_fill_color(target, oval_id, "#00AA66"))
        ok("remove_fill(oval)", ctrl.shape_remove_fill(target, oval_id))

        section("Shapes — z-order / group / align")
        ok("z_order(rect, send_to_back)", ctrl.shape_z_order(target, rect_id, "send_to_back"))
        ok("z_order(rect, bring_forward)", ctrl.shape_z_order(target, rect_id, "bring_forward"))

        # Add two more for align/distribute/group tests
        sq1 = ctrl.shape_add(target, "rectangle", 100, 400, 80, 60)
        sq2 = ctrl.shape_add(target, "rectangle", 250, 420, 80, 60)
        sq3 = ctrl.shape_add(target, "rectangle", 400, 380, 80, 60)
        ok("3 squares for align test", [sq1["shape_id"], sq2["shape_id"], sq3["shape_id"]])

        ok(
            "align(tops, 3 squares)",
            ctrl.shape_align(target, [sq1["shape_id"], sq2["shape_id"], sq3["shape_id"]], "top"),
        )
        ok(
            "distribute(horizontal, 3 squares)",
            ctrl.shape_distribute(
                target,
                [sq1["shape_id"], sq2["shape_id"], sq3["shape_id"]],
                "horizontal",
            ),
        )

        grp = ctrl.shape_group(target, [sq1["shape_id"], sq2["shape_id"], sq3["shape_id"]])
        ok("group(3 squares)", grp)
        ok("ungroup", ctrl.shape_ungroup(target, grp["group_shape_id"]))

        # ── RICH TEXT (Stage 3a) ──
        section("Text — rich (set/append/font/paragraph)")
        # Reuse the rectangle which already has 'hello rect' text
        ok(
            "text_set(rect, 'Multi-line\\rsecond line')",
            ctrl.text_set(target, rect_id, "Multi-line\rsecond line"),
        )
        ok(
            "text_append(rect, ' third', new_paragraph=True)",
            ctrl.text_append(target, rect_id, "third", new_paragraph=True),
        )
        ok(
            "text_set_font(rect, size=24 bold color=#FFFFFF)",
            ctrl.text_set_font(target, rect_id, size=24, bold=True, color="#FFFFFF"),
        )
        ok(
            "text_set_font(rect, paragraph=1 italic=True)",
            ctrl.text_set_font(target, rect_id, italic=True, paragraph=1),
        )
        ok(
            "text_set_paragraph(rect, alignment=center, all paragraphs)",
            ctrl.text_set_paragraph(target, rect_id, alignment="center"),
        )
        ok(
            "text_set_paragraph(rect, paragraph=2, bullet=bullet, indent=2)",
            ctrl.text_set_paragraph(target, rect_id, bullet="bullet", indent_level=2, paragraph=2),
        )
        ok("text_get_runs(rect)", ctrl.text_get_runs(target, rect_id))

        # ── PICTURES (Stage 3b) ──
        section("Pictures — add / crop / border / replace")
        # Build a throwaway test image (200x120 solid blue with red border)
        from PIL import Image, ImageDraw

        img_path = os.path.join(tempfile.gettempdir(), "ppt_test_img.png")
        img_path2 = os.path.join(tempfile.gettempdir(), "ppt_test_img2.png")
        im = Image.new("RGB", (200, 120), color=(40, 100, 220))
        ImageDraw.Draw(im).rectangle([0, 0, 199, 119], outline=(220, 40, 40), width=6)
        im.save(img_path, "PNG")
        im2 = Image.new("RGB", (200, 120), color=(40, 200, 80))
        im2.save(img_path2, "PNG")

        pic = ctrl.picture_add_from_file(target, img_path, left=550, top=50, width=240, height=140)
        ok("picture_add_from_file", pic)
        pic_id = pic["shape_id"]

        ok(
            "picture_set_border(red, 3pt, dash)",
            ctrl.picture_set_border(target, pic_id, color="#FF0000", weight=3, dash_style="dash"),
        )
        ok("picture_crop(left=10,top=10)", ctrl.picture_crop(target, pic_id, left=10, top=10))

        rep = ctrl.picture_replace(target, pic_id, img_path2)
        ok("picture_replace(→ green img)", rep)

        # URL fetch (skip on failure — offline boxes shouldn't break the suite)
        try:
            url_pic = ctrl.picture_add_from_url(
                target,
                "https://httpbin.org/image/png",
                left=550,
                top=220,
                width=180,
                height=120,
            )
            ok("picture_add_from_url", url_pic)
        except Exception as e:
            print(f"  [SKIP] picture_add_from_url: {e}")

        # Cleanup local test images (PNG inside .pptx is now embedded)
        for p in (img_path, img_path2):
            try:
                os.remove(p)
            except OSError:
                pass

        # ── TABLES (Stage 3c) ──
        section("Tables — add / cell text / cell fill / sizes / merge / rows+cols")
        # Drop a fresh slide so the table has space
        ctrl.slide_add()
        tbl_slide = ctrl.pres.Slides.Count
        tbl = ctrl.table_add(tbl_slide, rows=3, cols=4, left=50, top=80, width=600, height=300)
        ok("table_add(3x4)", tbl)
        tbl_id = tbl["shape_id"]

        # Header row
        for c, h in enumerate(["Quarter", "Plan", "Actual", "Δ"], start=1):
            ctrl.table_set_cell_text(tbl_slide, tbl_id, 1, c, h)
        # Two data rows
        rows_data = [("Q1", "100", "110", "+10"), ("Q2", "120", "115", "-5")]
        for r, vals in enumerate(rows_data, start=2):
            for c, v in enumerate(vals, start=1):
                ctrl.table_set_cell_text(tbl_slide, tbl_id, r, c, v)
        ok("table_set_cell_text(× 12 cells)", "filled")

        ok(
            "table_set_cell_fill(1,1..4 = blue header)",
            [ctrl.table_set_cell_fill(tbl_slide, tbl_id, 1, c, "#2A5DB0") for c in range(1, 5)],
        )
        ok("table_set_column_width(1, 100)", ctrl.table_set_column_width(tbl_slide, tbl_id, 1, 100))
        ok("table_set_row_height(1, 50)", ctrl.table_set_row_height(tbl_slide, tbl_id, 1, 50))

        ok("table_add_row (append)", ctrl.table_add_row(tbl_slide, tbl_id))
        ok("table_add_row (position=2)", ctrl.table_add_row(tbl_slide, tbl_id, position=2))
        ok("table_add_column (append)", ctrl.table_add_column(tbl_slide, tbl_id))

        # Merge two header cells
        ok("table_merge_cells((1,2)-(1,3))", ctrl.table_merge_cells(tbl_slide, tbl_id, 1, 2, 1, 3))

        ok("table_get_info", ctrl.table_get_info(tbl_slide, tbl_id))

        # ── CONNECTORS (Stage 4a) ──
        section("Connectors — add / connect / reroute / disconnect")
        ctrl.slide_add()
        conn_slide = ctrl.pres.Slides.Count
        # Two boxes to connect
        box_a = ctrl.shape_add(conn_slide, "rectangle", 100, 100, 140, 80, text="A")
        box_b = ctrl.shape_add(conn_slide, "rectangle", 500, 320, 140, 80, text="B")
        ok("two boxes", [box_a["shape_id"], box_b["shape_id"]])

        conn = ctrl.connector_add(conn_slide, "elbow", 240, 140, 500, 360)
        ok("connector_add(elbow)", conn)
        conn_id = conn["shape_id"]
        ok(
            "connector_connect(A→B)",
            ctrl.connector_connect(
                conn_slide,
                conn_id,
                begin_shape_id=box_a["shape_id"],
                end_shape_id=box_b["shape_id"],
            ),
        )
        # Move box A — connector should auto-reroute
        ctrl.shape_move(conn_slide, box_a["shape_id"], 80, 60)
        ok("connector_reroute", ctrl.connector_reroute(conn_slide, conn_id))
        ok("connector_disconnect(both)", ctrl.connector_disconnect(conn_slide, conn_id, end="both"))

        # ── CHARTS (Stage 4b) ──
        section("Charts — add / set_data / set_title / set_type")
        ctrl.slide_add()
        chart_slide = ctrl.pres.Slides.Count
        try:
            chart = ctrl.chart_add(
                chart_slide,
                "column",
                80,
                80,
                600,
                380,
                categories=["Q1", "Q2", "Q3", "Q4"],
                series={"Plan": [100, 120, 140, 160], "Actual": [110, 115, 150, 155]},
                title="Quarterly Plan vs Actual",
            )
            ok("chart_add(column + data + title)", chart)
            chart_id = chart["shape_id"]

            ok(
                "chart_set_data(new series)",
                ctrl.chart_set_data(
                    chart_slide,
                    chart_id,
                    categories=["A", "B", "C"],
                    series={"Score": [10, 25, 18]},
                ),
            )
            ok("chart_set_title('Updated')", ctrl.chart_set_title(chart_slide, chart_id, "Updated"))
            ok("chart_set_type(bar)", ctrl.chart_set_type(chart_slide, chart_id, "bar"))
        except PowerPointError as e:
            print(f"  [SKIP] charts (Excel may not be installed): {e}")

        # ── HYPERLINKS (Stage 4c) ──
        section("Hyperlinks — set / get / remove")
        # Add a box on the connector slide to attach a link to
        link_box = ctrl.shape_add(
            conn_slide, "rounded_rectangle", 100, 450, 220, 60, text="Open Anthropic"
        )
        link_id = link_box["shape_id"]
        ok(
            "hyperlink_set(URL)",
            ctrl.hyperlink_set(
                conn_slide, link_id, "https://www.anthropic.com", screen_tip="Anthropic homepage"
            ),
        )
        ok("hyperlink_get", ctrl.hyperlink_get(conn_slide, link_id))
        ok("hyperlink_set(slide://1)", ctrl.hyperlink_set(conn_slide, link_id, "slide://1"))
        ok("hyperlink_remove", ctrl.hyperlink_remove(conn_slide, link_id))

        # ── POLISH (Stage 5) ──
        # Switch to slow-step mode so the new polish section is watchable.
        _enable_pause()
        section("Polish — chart styling")
        try:
            ctrl.slide_add()
            polish_slide = ctrl.pres.Slides.Count
            polish_chart = ctrl.chart_add(
                polish_slide,
                "column",
                60,
                60,
                660,
                380,
                categories=["Q1", "Q2", "Q3", "Q4"],
                series={"Plan": [100, 120, 140, 160], "Actual": [110, 115, 150, 155]},
                title="Quarterly Revenue ($M)",
            )
            pc_id = polish_chart["shape_id"]
            ok("chart_set_style(2)", ctrl.chart_set_style(polish_slide, pc_id, 2))
            ok(
                "chart_set_series_color(Plan, grey)",
                ctrl.chart_set_series_color(polish_slide, pc_id, "Plan", "#A0A0A0"),
            )
            ok(
                "chart_set_series_color(Actual, brand blue)",
                ctrl.chart_set_series_color(polish_slide, pc_id, "Actual", "#1F4E79"),
            )
            ok(
                "chart_set_data_labels(value, $#,##0)",
                ctrl.chart_set_data_labels(
                    polish_slide, pc_id, show="value", number_format="$#,##0"
                ),
            )
            ok(
                "chart_set_axis(y, min=0, format=$M, title='Revenue ($M)')",
                ctrl.chart_set_axis(
                    polish_slide,
                    pc_id,
                    axis="y",
                    minimum=0,
                    number_format="$#,##0",
                    title="Revenue ($M)",
                ),
            )
            ok(
                "chart_set_legend(bottom)",
                ctrl.chart_set_legend(polish_slide, pc_id, position="bottom"),
            )
            ok(
                "chart_set_gridlines(y major=on minor=off)",
                ctrl.chart_set_gridlines(polish_slide, pc_id, axis="y", major=True, minor=False),
            )
        except PowerPointError as e:
            print(f"  [SKIP] chart polish (Excel may not be installed): {e}")

        section("Polish — table styling")
        ctrl.slide_add()
        polish_tbl_slide = ctrl.pres.Slides.Count
        ptbl = ctrl.table_add(
            polish_tbl_slide, rows=3, cols=4, left=50, top=80, width=600, height=200
        )
        ptbl_id = ptbl["shape_id"]
        for c, h in enumerate(["Quarter", "Plan", "Actual", "Δ"], start=1):
            ctrl.table_set_cell_text(polish_tbl_slide, ptbl_id, 1, c, h)
        for r, vals in enumerate(
            [("Q1", "$100", "$110", "+10"), ("Q2", "$120", "$115", "-5")], start=2
        ):
            for c, v in enumerate(vals, start=1):
                ctrl.table_set_cell_text(polish_tbl_slide, ptbl_id, r, c, v)
        ok(
            "table_apply_style(medium_style_2_accent_1)",
            ctrl.table_apply_style(polish_tbl_slide, ptbl_id, "medium_style_2_accent_1"),
        )
        ok(
            "table_set_cell_font(1,1 bold size=14)",
            ctrl.table_set_cell_font(
                polish_tbl_slide, ptbl_id, 1, 1, bold=True, size=14, color="#FFFFFF"
            ),
        )
        ok(
            "table_set_cell_alignment(2..3, 2..4 → right)",
            [
                ctrl.table_set_cell_alignment(
                    polish_tbl_slide, ptbl_id, r, c, horizontal="right", vertical="middle"
                )
                for r in (2, 3)
                for c in (2, 3, 4)
            ],
        )

        section("Polish — shape format copy + shadow")
        ctrl.slide_add()
        polish_sh_slide = ctrl.pres.Slides.Count
        kpi_a = ctrl.shape_add(
            polish_sh_slide, "rounded_rectangle", 60, 120, 200, 120, text="$2.3M\nRevenue"
        )
        kpi_b = ctrl.shape_add(
            polish_sh_slide, "rounded_rectangle", 300, 120, 200, 120, text="+12%\nGrowth"
        )
        kpi_c = ctrl.shape_add(
            polish_sh_slide, "rounded_rectangle", 540, 120, 200, 120, text="42\nCustomers"
        )
        # Style the first tile fully
        ctrl.shape_set_fill_color(polish_sh_slide, kpi_a["shape_id"], "#1F4E79")
        ctrl.shape_set_line_color(polish_sh_slide, kpi_a["shape_id"], "#1F4E79")
        ok(
            "shape_set_shadow(kpi_a, offset 0/4 blur=8 transp=0.6)",
            ctrl.shape_set_shadow(
                polish_sh_slide,
                kpi_a["shape_id"],
                visible=True,
                color="#000000",
                transparency=0.6,
                blur=8,
                offset_x=0,
                offset_y=4,
            ),
        )
        ok(
            "shape_copy_format(kpi_a → [kpi_b, kpi_c])",
            ctrl.shape_copy_format(
                polish_sh_slide, kpi_a["shape_id"], [kpi_b["shape_id"], kpi_c["shape_id"]]
            ),
        )

        section("Polish — slide background + footer")
        ok(
            "slide_set_background(color=#1F4E79)",
            ctrl.slide_set_background(polish_sh_slide, color="#1F4E79"),
        )
        ok(
            "slide_set_footer(deck-wide)",
            ctrl.slide_set_footer(
                footer="Confidential — 2026 Board Meeting",
                slide_number=True,
                date=True,
                date_format="May 18, 2026",
            ),
        )

        # ── SAVE AS PDF ──
        section("Save As (PDF)")
        ok("save(as pdf)", ctrl.save(file_path=save_as_pdf, format="pdf"))

        # ── CONNECT / CLOSE ──
        section("Connect / Close / Quit")
        # Test reconnect-by-name path
        name_only = os.path.basename(ctrl.pres.FullName)
        ctrl_ref = ctrl.pres.Name  # noqa: F841
        ok("connect (by name)", ctrl.connect(name_only))
        ok("close (no save)", ctrl.close(save=False))

        # Re-open to verify
        section("Open from disk")
        ok("open(tmp_path)", ctrl.open(tmp_path))
        ok("close", ctrl.close(save=False))

        print("\nAll Stage 1 tests completed.\n")

    except PowerPointError as e:
        print(f"\n[FAIL] PowerPointError: {e}\n")
        raise
    finally:
        # Clean up files (PowerPoint app left running for inspection)
        for p in (tmp_path, save_as_pdf):
            if os.path.exists(p):
                try:
                    os.remove(p)
                except OSError:
                    pass


if __name__ == "__main__":
    run_tests()
