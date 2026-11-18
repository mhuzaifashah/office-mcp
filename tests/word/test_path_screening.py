"""Non-local (UNC / URL) paths are refused before any filesystem access."""

from __future__ import annotations

import base64
from pathlib import Path

import pytest

from office_mcp.word.wordcom.exceptions import PathNotAllowedError
from office_mcp.word.wordcom.runtime import _ops
from office_mcp.word.wordcom.runtime._path_policy import PathPolicy, reject_nonlocal_path

UNC = "\\\\attacker\\share\\x.png"


@pytest.mark.parametrize("raw", [UNC, "//attacker/share/x", "http://h/x", "file:///c:/x"])
def test_reject_nonlocal_path(raw):
    with pytest.raises(PathNotAllowedError):
        reject_nonlocal_path(raw)


@pytest.mark.parametrize("raw", ["C:\\docs\\a.docx", "relative\\a.docx", "a.docx"])
def test_local_paths_pass(raw):
    reject_nonlocal_path(raw)


def test_save_target_unc_refused_without_resolving(monkeypatch):
    def boom(self, *a, **k):
        raise AssertionError("resolve() must not run for a UNC target")

    monkeypatch.setattr(Path, "resolve", boom)
    with pytest.raises(PathNotAllowedError):
        PathPolicy().resolve_save_target(UNC)


def test_save_disabled_refused_without_resolving(monkeypatch):
    def boom(self, *a, **k):
        raise AssertionError("resolve() must not run when saving is disabled")

    monkeypatch.setattr(Path, "resolve", boom)
    with pytest.raises(PathNotAllowedError, match="saving is disabled"):
        PathPolicy().resolve_save_target("C:\\out\\a.docx")


@pytest.mark.parametrize(
    "op",
    [
        {"op": "insert_image", "anchor_id": "para:1", "path": UNC},
        {"op": "replace_shape_image", "anchor_id": "shape:1", "path": UNC},
        {"op": "apply_theme", "theme": UNC},
        {"op": "set_theme_colors", "scheme": UNC},
        {"op": "set_theme_fonts", "scheme": UNC},
    ],
)
def test_batch_screening_covers_every_path_field(op):
    with pytest.raises(PathNotAllowedError):
        PathPolicy().screen_op_image_paths([op])


def test_builtin_theme_names_pass():
    PathPolicy().screen_op_image_paths([{"op": "apply_theme", "theme": "Office"}])


@pytest.mark.parametrize("kind", ["insert_image", "replace_shape_image"])
def test_base64_field_is_decoded_not_probed(fake_word, monkeypatch, kind):
    """A `base64` value that is really a path must never reach is_file()."""
    from office_mcp.word import wordcom
    from office_mcp.word.wordcom.runtime._ops import run_batch

    probed: list[str] = []
    real_is_file = Path.is_file

    def spy(self):
        probed.append(str(self))
        return real_is_file(self)

    monkeypatch.setattr(Path, "is_file", spy)
    anchor = "para:1" if kind == "insert_image" else "shape:1"
    with wordcom.attach() as word:
        doc = word.documents.active
        _result, exc = run_batch(doc, [{"op": kind, "anchor_id": anchor, "base64": UNC}], label="t")
    assert exc is not None
    assert not any("attacker" in p for p in probed)


def test_base64_field_still_accepts_real_base64():
    assert _ops._decode_base64(base64.b64encode(b"png").decode()) == b"png"


def test_controller_open_refuses_unc(monkeypatch):
    import os

    from office_mcp.word import WordController

    monkeypatch.setattr(os.path, "exists", lambda p: pytest.fail("exists() probed a UNC path"))
    with pytest.raises(PathNotAllowedError):
        WordController.open(object.__new__(WordController), UNC)
