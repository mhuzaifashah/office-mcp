"""Live smoke test of the PowerPoint tools' backend against real PowerPoint.

Works in a new presentation whose window is minimized at once (PowerPoint
charts need a window), so the user's open decks are untouched. Quits
PowerPoint afterwards only if it wasn't running before.

    uv run python scripts/smoke_ppt.py
"""

from __future__ import annotations

import json
import os
import tempfile
import time

import pywintypes
import win32com.client

from office_mcp.ppt import service as svc_mod
from office_mcp.ppt.service import PowerPointService

state: dict = {}


def find_presentation(name=None):
    if "pres" not in state:
        try:
            win32com.client.GetActiveObject("PowerPoint.Application")
            state["was_running"] = True
        except pywintypes.com_error:
            state["was_running"] = False
        app = win32com.client.Dispatch("PowerPoint.Application")
        pres = app.Presentations.Add(-1)  # WithWindow: charts need one
        pres.Windows(1).WindowState = 2  # ppWindowMinimized
        pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts(1))
        state["app"], state["pres"] = app, pres
    return state["pres"]


svc_mod.find_presentation = find_presentation
out_dir = tempfile.mkdtemp(prefix="omcp-ppt-live-")
s = PowerPointService(save_dirs=(out_dir,))
failures: list[str] = []


def step(name, fn, *, expect_error=False):
    t0 = time.perf_counter()
    try:
        result, ok = fn(), True
    except Exception as e:  # noqa: BLE001
        result, ok = f"{type(e).__name__}: {e} | {getattr(e, 'details', '')}", False
    text = result if isinstance(result, str) else json.dumps(result, default=str)
    print(f"[{'OK ' if ok else 'ERR'}] {name} ({time.perf_counter() - t0:.2f}s): {text[:300]}")
    if ok == expect_error:
        failures.append(name)
    return result


try:
    step("status", lambda: s.read("status"))
    step("layouts", lambda: s.read("layouts"))
    step(
        "title slide + content slide",
        lambda: s.exec(
            [
                {
                    "op": "set_placeholder",
                    "slide": 1,
                    "placeholder": "title",
                    "text": "Quarterly review",
                },
                {"op": "set_placeholder", "slide": 1, "placeholder": "subtitle", "text": "Q2 2026"},
                {"op": "add_slide", "layout": "title_only"},
                {
                    "op": "set_placeholder",
                    "slide": 2,
                    "placeholder": "title",
                    "text": "Sales by region",
                },
                {
                    "op": "add_chart",
                    "slide": 2,
                    "chart_type": "column_clustered",
                    "left": 60,
                    "top": 120,
                    "width": 520,
                    "height": 360,
                    "categories": ["Q1", "Q2"],
                    "series": [
                        {"name": "North", "values": [120, 135]},
                        {"name": "South", "values": [98, 110]},
                    ],
                },
                {
                    "op": "add_text_box",
                    "slide": 2,
                    "left": 600,
                    "top": 140,
                    "width": 300,
                    "height": 120,
                    "text": "North grew 12%",
                },
                {"op": "set_notes", "slide": 2, "notes": "Mention the North region first."},
            ]
        ),
    )
    info = step("slide 2", lambda: s.read("slide", slide=2))
    box = next(sh["id"] for sh in info["shapes"] if sh.get("text") == "North grew 12%")
    step(
        "format the text box",
        lambda: s.exec(
            [
                {
                    "op": "format_text",
                    "slide": 2,
                    "shape": box,
                    "size": 24,
                    "bold": True,
                    "color": "#1F4E79",
                },
                {"op": "fill", "slide": 2, "shape": box, "color": "#DDEBF7"},
                {
                    "op": "add_table",
                    "slide": 1,
                    "rows": 2,
                    "cols": 2,
                    "left": 300,
                    "top": 400,
                    "width": 360,
                    "height": 80,
                },
            ]
        ),
    )
    table = next(
        sh["id"] for sh in s.read("slide", slide=1)["shapes"] if sh["shape_type_name"] == "table"
    )
    step(
        "fill the table",
        lambda: s.exec(
            [
                {"op": "set_cell", "slide": 1, "shape": table, "row": 1, "col": 1, "text": "Owner"},
                {"op": "set_cell", "slide": 1, "shape": table, "row": 1, "col": 2, "text": "Due"},
            ]
        ),
    )
    step("table read", lambda: s.read("table", slide=1, shape=table))
    step("text of every slide", lambda: s.read("text"))
    step("notes", lambda: s.read("notes", slide=2))
    step("runs", lambda: s.read("runs", slide=2, shape=box))
    step(
        "bad shape id is caught first (expected ERR)",
        lambda: s.exec([{"op": "set_text", "slide": 2, "shape": 9999, "text": "x"}]),
        expect_error=True,
    )
    before = {
        "slides": s.read("slides"),
        "notes": s.read("notes", slide=1),
        "slide2": s.read("slide", slide=2),
    }
    err = step(
        "failure part-way undoes the whole batch (expected ERR)",
        lambda: s.exec(
            [
                {"op": "set_notes", "slide": 1, "notes": "should be undone"},
                {"op": "add_slide", "layout": "blank"},
                {
                    "op": "add_text_box",
                    "slide": 2,
                    "left": 10,
                    "top": 10,
                    "width": 100,
                    "height": 30,
                    "text": "should be undone",
                },
                {"op": "format_paragraph", "slide": 2, "shape": box, "indent_level": 9},
            ]
        ),
        expect_error=True,
    )
    if "'rolled_back': True" not in str(err):
        failures.append("rollback reported")

    def unchanged():
        after = {
            "slides": s.read("slides"),
            "notes": s.read("notes", slide=1),
            "slide2": s.read("slide", slide=2),
        }
        if after != before:
            raise AssertionError(f"deck changed: {before} -> {after}")
        return "deck matches the state before the failed batch"

    step("rollback restored notes, slides and shapes", unchanged)

    def snap():
        shots, total = s.snapshot(slides="1-2", max_dim=1200)
        for n, png in shots:
            with open(os.path.join(out_dir, f"slide_{n}.png"), "wb") as f:
                f.write(png)
        return [(n, len(png)) for n, png in shots], total

    step("snapshot", snap)
    step("save_as", lambda: s.write("save_as", path=os.path.join(out_dir, "deck.pptx")))
    step("save", lambda: s.write("save"))
    step("export_pdf", lambda: s.write("export_pdf", path=os.path.join(out_dir, "deck.pdf")))
    step(
        "save_as outside the folder (expected ERR)",
        lambda: s.write("save_as", path="C:\\Windows\\Temp\\x.pptx"),
        expect_error=True,
    )
    print("files:", sorted(os.listdir(out_dir)))
finally:

    def cleanup():
        if "pres" in state:
            state["pres"].Saved = True
            state["pres"].Close()
            if not state["was_running"] and state["app"].Presentations.Count == 0:
                state["app"].Quit()

    s._run(cleanup)
    s.close()
print("output folder:", out_dir)
print("unexpected results:", failures or "none")
