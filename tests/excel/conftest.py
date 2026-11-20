"""A minimal fake workbook: sheets of cells addressed by A1 strings.

Enough for the batch logic (validation, saving and restoring cell contents,
alerts) without Excel. The real controller's COM calls are covered by the
live smoke test (`office_mcp.excel.test_excel`) instead.
"""

from __future__ import annotations

from types import SimpleNamespace
from typing import Any

import pytest


class FakeRangeApi:
    Count = 1

    def __init__(self, sheet: FakeSheet, address: str) -> None:
        self._sheet = sheet
        self.Address = address

    @property
    def Formula2(self) -> Any:
        return self._sheet.cells.get(self.Address, "")

    @Formula2.setter
    def Formula2(self, value: Any) -> None:
        self._sheet.cells[self.Address] = value


class FakeRange:
    def __init__(self, sheet: FakeSheet, address: str) -> None:
        self.address = address
        self.api = FakeRangeApi(sheet, address)


class FakeSheet:
    def __init__(self, name: str) -> None:
        self.name = name
        self.cells: dict[str, Any] = {}

    def range(self, address: str) -> FakeRange:
        return FakeRange(self, address)


class FakeSheets:
    def __init__(self, names: list[str]) -> None:
        self._items = [FakeSheet(n) for n in names]

    def __iter__(self):
        return iter(self._items)

    def __getitem__(self, key: str | int) -> FakeSheet:
        if isinstance(key, int):
            return self._items[key]
        for s in self._items:
            if s.name == key:
                return s
        raise KeyError(key)

    @property
    def active(self) -> FakeSheet:
        return self._items[0]


class FakeBook:
    def __init__(self, sheets: list[str] | None = None) -> None:
        self.name = "Book1.xlsx"
        self.fullname = "C:\\data\\Book1.xlsx"
        self.sheets = FakeSheets(sheets or ["Data", "Notes"])
        self.app = SimpleNamespace(api=SimpleNamespace(DisplayAlerts=True))
        self.api = SimpleNamespace(Saved=True)
        self.events: list[str] = []


@pytest.fixture
def book() -> FakeBook:
    return FakeBook()
