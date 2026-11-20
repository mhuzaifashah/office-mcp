"""Live smoke test of the Excel tools' backend against real Excel.

Starts its own hidden Excel (never touches the user's workbooks), runs every
area once (cells, formulas, formatting, tables, charts, pivots, reads, a failed
batch with rollback, snapshot, save_as, export_pdf), then closes it.

    uv run python scripts/smoke_excel.py

The snapshot is written next to the temp output folder it prints.
"""

import json
import os
import tempfile
import time

import xlwings as xw

from office_mcp.excel import service as svc_mod
from office_mcp.excel.service import ExcelService

state = {}


def find_book(name=None):
    if "book" not in state:
        app = xw.App(visible=False, add_book=False)
        state["app"] = app
        state["book"] = app.books.add()
        state["book"].sheets[0].name = "Sales"
    return state["book"]


svc_mod.find_book = find_book
out_dir = tempfile.mkdtemp(prefix="omcp-xl-live-")
s = ExcelService(save_dirs=(out_dir,))
results = []


def step(name, fn):
    t0 = time.perf_counter()
    try:
        r = fn()
        ok = True
    except Exception as e:  # noqa: BLE001
        r = f"{type(e).__name__}: {e} | {getattr(e, 'details', '')}"
        ok = False
    dt = time.perf_counter() - t0
    short = r if isinstance(r, str) else json.dumps(r, default=str)[:300]
    print(f"[{'OK ' if ok else 'ERR'}] {name} ({dt:.2f}s): {short}")
    results.append(ok)
    return r


try:
    step(
        "write+formula+format+chart",
        lambda: s.exec(
            [
                {
                    "op": "write",
                    "range": "A1",
                    "values": [
                        ["Region", "Q1", "Q2"],
                        ["North", 120, 135],
                        ["South", 98, 110],
                        ["East", 143, 150],
                    ],
                },
                {"op": "write", "range": "D1", "values": "Total"},
                {"op": "fill_formula", "range": "D2:D4", "formula": "=B2+C2"},
                {
                    "op": "format",
                    "range": "A1:D1",
                    "bold": True,
                    "fill": "#DDEBF7",
                    "align": "center",
                },
                {"op": "format", "range": "B2:D4", "number_format": "#,##0"},
                {"op": "autofit", "range": "A:D"},
                {
                    "op": "create_chart",
                    "range": "A1:C4",
                    "type": "column_clustered",
                    "title": "Sales by region",
                    "at": "F2:L16",
                    "name": "SalesChart",
                },
            ]
        ),
    )
    step("read with formulas", lambda: s.read("read", range="A1:D4", formulas=True))
    step("read used range (no range)", lambda: s.read("read"))
    step("status", lambda: s.read("status"))
    step("sheets", lambda: s.read("sheets"))
    step("find", lambda: s.read("find", text="South"))
    step("summary", lambda: s.read("summary", range="D2:D4"))
    step("charts", lambda: s.read("charts"))
    step(
        "table + pivot",
        lambda: s.exec(
            [
                {
                    "op": "create_table",
                    "range": "A1:D4",
                    "name": "SalesTbl",
                    "style": "TableStyleMedium2",
                },
                {"op": "add_sheet", "name": "Pivot"},
                {
                    "op": "create_pivot",
                    "name": "SalesPivot",
                    "table": "SalesTbl",
                    "dest_sheet": "Pivot",
                    "dest_cell": "A3",
                },
                {"op": "pivot_add_field", "pivot": "SalesPivot", "field": "Region", "area": "row"},
                {
                    "op": "pivot_add_field",
                    "pivot": "SalesPivot",
                    "field": "Q1",
                    "area": "value",
                    "function": "sum",
                },
            ]
        ),
    )
    step("tables", lambda: s.read("tables"))
    step("table_data", lambda: s.read("table_data", table="SalesTbl"))
    step("pivots", lambda: s.read("pivots"))
    step("pivot_fields", lambda: s.read("pivot_fields", pivot="SalesPivot"))
    # rollback: overwrite B2 then fail on a missing table
    step(
        "rollback batch (expected ERR)",
        lambda: s.exec(
            [
                {"op": "write", "sheet": "Sales", "range": "B2", "values": 999},
                {"op": "format", "sheet": "Sales", "range": "B2", "bold": True},
                {"op": "table_style", "table": "Nope", "style": "TableStyleLight1"},
            ]
        ),
    )
    r = step("B2 restored?", lambda: s.read("read", sheet="Sales", range="B2"))
    results.append(isinstance(r, dict) and r["values"] == [[120]])
    print("   B2 restored to 120:", results[-1])
    shots = step(
        "snapshot",
        lambda: [
            (p, open(os.path.join(out_dir, f"snapshot_{p}.png"), "wb").write(b))
            for p, b in s.snapshot(sheet="Sales", address="A1:L18")
        ],
    )
    step("save_as", lambda: s.write("save_as", path=os.path.join(out_dir, "live.xlsx")))
    step(
        "export_pdf",
        lambda: s.write("export_pdf", path=os.path.join(out_dir, "live.pdf"), sheet="Sales"),
    )
    step("save", lambda: s.write("save"))
    r = step(
        "snapshot keeps Saved",
        lambda: (
            s.snapshot(sheet="Sales", address="A1:D4"),
            s._run(lambda: bool(state["book"].api.Saved)),
        )[1],
    )
    results.append(r is True)
    step(
        "save_as outside dir (expected ERR)",
        lambda: s.write("save_as", path="C:\\Windows\\Temp\\x.xlsx"),
    )
    print("files:", os.listdir(out_dir))
finally:

    def cleanup():
        if "book" in state:
            state["book"].close()
            state["app"].quit()

    s._run(cleanup)
    s.close()
print("output folder:", out_dir)
# Two steps fail on purpose (the rollback batch and the save outside the folder).
print("unexpected failures:", results.count(False) - 2)
