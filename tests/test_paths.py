"""Path rules shared by the Excel and PowerPoint tools."""

from __future__ import annotations

import pytest

from office_mcp.paths import PathNotAllowed, SavePolicy, reject_nonlocal


@pytest.mark.parametrize(
    "path", ["\\\\host\\share\\a.xlsx", "//host/share/a.xlsx", "http://x/a.xlsx", "file:///C:/a"]
)
def test_network_paths_and_urls_are_refused(path):
    with pytest.raises(PathNotAllowed):
        reject_nonlocal(path)


def test_drive_paths_pass():
    reject_nonlocal("C:\\data\\a.xlsx")


def test_saving_is_off_without_folders(tmp_path):
    with pytest.raises(PathNotAllowed, match="--save-dir"):
        SavePolicy().resolve(tmp_path / "a.xlsx")


def test_only_allowed_folders_are_written(tmp_path):
    allowed = tmp_path / "out"
    allowed.mkdir()
    policy = SavePolicy([allowed])
    assert policy.resolve(allowed / "a.xlsx") == (allowed / "a.xlsx").resolve()
    with pytest.raises(PathNotAllowed):
        policy.resolve(tmp_path / "a.xlsx")
    with pytest.raises(PathNotAllowed):
        policy.resolve(allowed / ".." / "a.xlsx")
