"""Persistence (save / save-as / export-pdf), the path policy, and image hardening.

The Python API is ungated; the CLI / MCP surfaces gate writes behind a
default-deny directory whitelist and reject non-local image-source paths.
Also covers the verb-first bookmark / section CLI consolidation.
"""

from __future__ import annotations

import pytest

from office_mcp.word.wordcom import attach
from office_mcp.word.wordcom.exceptions import OpError, PathNotAllowedError
from office_mcp.word.wordcom.runtime._dispatch import dispatch_write
from office_mcp.word.wordcom.runtime._path_policy import PathPolicy, reject_nonlocal_image_path
from office_mcp.word.wordcom.runtime._worker import InlineWorker

_BS = chr(92)  # a single backslash, unambiguous through every quoting layer
_UNC = _BS + _BS + "host" + _BS + "share" + _BS + "a.png"


def _active():
    """The Document wrapper for the fake_word active document."""
    with attach() as word:
        return word.documents.active


# ---------------------------------------------------------------------------
# PathPolicy
# ---------------------------------------------------------------------------


class TestPathPolicy:
    def test_deny_all_by_default(self, tmp_path):
        with pytest.raises(PathNotAllowedError, match="saving is disabled"):
            PathPolicy().resolve_save_target(tmp_path / "x.docx")

    def test_inside_whitelist(self, tmp_path):
        pol = PathPolicy(save_dirs=[tmp_path])
        assert pol.resolve_save_target(tmp_path / "out.docx") == (tmp_path / "out.docx").resolve()

    def test_escape_via_dotdot_blocked(self, tmp_path):
        allowed = tmp_path / "allowed"
        allowed.mkdir()
        pol = PathPolicy(save_dirs=[allowed])
        with pytest.raises(PathNotAllowedError, match="outside"):
            pol.resolve_save_target(allowed / ".." / "evil.docx")

    def test_outside_whitelist(self, tmp_path):
        pol = PathPolicy(save_dirs=[tmp_path / "a"])
        with pytest.raises(PathNotAllowedError, match="outside"):
            pol.resolve_save_target(tmp_path / "b" / "x.docx")

    def test_saving_enabled_flag(self, tmp_path):
        assert PathPolicy().saving_enabled is False
        assert PathPolicy(save_dirs=[tmp_path]).saving_enabled is True

    def test_from_env(self, monkeypatch, tmp_path):
        monkeypatch.setenv("WORDCOM_SAVE_DIRS", str(tmp_path))
        pol = PathPolicy.from_env()
        assert pol.saving_enabled
        assert pol.resolve_save_target(tmp_path / "x.docx") == (tmp_path / "x.docx").resolve()

    def test_from_env_merges_extra(self, monkeypatch, tmp_path):
        monkeypatch.delenv("WORDCOM_SAVE_DIRS", raising=False)
        pol = PathPolicy.from_env(extra_save=[tmp_path])
        assert pol.resolve_save_target(tmp_path / "x.docx")


@pytest.mark.parametrize(
    "bad",
    [_UNC, "//host/share/a.png", "http://example.com/a.png", "https://x/a.png", "file:///etc/x"],
)
def test_reject_nonlocal_image_path(bad):
    with pytest.raises(PathNotAllowedError):
        reject_nonlocal_image_path(bad)


@pytest.mark.parametrize("ok", [r"C:\img\a.png", "rel/a.png", "/abs/unix/a.png", "a.png"])
def test_local_image_paths_pass(ok):
    reject_nonlocal_image_path(ok)  # must not raise


class TestImageScreening:
    def test_allowlist_inside_ok(self, tmp_path):
        PathPolicy(image_dirs=[tmp_path]).screen_image_path(tmp_path / "a.png")

    def test_allowlist_outside_blocked(self, tmp_path):
        outside = tmp_path / "other.png"
        pol = PathPolicy(image_dirs=[tmp_path / "imgs"])
        with pytest.raises(PathNotAllowedError, match="outside"):
            pol.screen_image_path(outside)

    def test_no_allowlist_accepts_local(self, tmp_path):
        PathPolicy().screen_image_path(tmp_path / "a.png")  # local always ok

    def test_op_scan_skips_base64(self):
        PathPolicy().screen_op_image_paths([{"op": "insert_image", "base64": "AAAA"}])

    def test_op_scan_rejects_unc_path(self):
        with pytest.raises(PathNotAllowedError):
            PathPolicy().screen_op_image_paths([{"op": "insert_image", "path": _UNC}])

    def test_op_scan_ignores_other_ops(self):
        PathPolicy().screen_op_image_paths([{"op": "append", "text": "hi"}])


# ---------------------------------------------------------------------------
# Document API (ungated) — against the fake document
# ---------------------------------------------------------------------------


class TestDocumentPersistence:
    def test_saved_property(self, fake_word):
        doc = _active()
        doc.com.Saved = True
        assert doc.saved is True
        doc.com.Saved = False
        assert doc.saved is False

    def test_save_writes_when_path_exists(self, fake_word, tmp_path):
        doc = _active()
        doc.com.Path = str(tmp_path)
        doc.com.FullName = str(tmp_path / "Test.docx")
        assert doc.save() == str(tmp_path / "Test.docx")
        doc.com.Save.assert_called_once()

    def test_save_refuses_never_saved(self, fake_word):
        doc = _active()
        doc.com.Path = ""
        with pytest.raises(OpError, match="never been saved"):
            doc.save()

    def test_save_as_calls_saveas2(self, fake_word, tmp_path):
        doc = _active()
        written = doc.save_as(tmp_path / "out.docx")
        assert written == str((tmp_path / "out.docx").resolve())
        doc.com.SaveAs2.assert_called_once()
        # docx format constant (wdFormatDocumentDefault == 16)
        assert doc.com.SaveAs2.call_args.kwargs["FileFormat"] == 16

    def test_save_as_refuses_overwrite(self, fake_word, tmp_path):
        existing = tmp_path / "out.docx"
        existing.write_text("x")
        doc = _active()
        with pytest.raises(OpError, match="overwrite"):
            doc.save_as(existing)
        doc.save_as(existing, overwrite=True)  # explicit override is allowed

    def test_save_as_rejects_pdf_format(self, fake_word, tmp_path):
        doc = _active()
        with pytest.raises(OpError, match="export_pdf"):
            doc.save_as(tmp_path / "out.pdf", fmt="pdf")

    def test_export_pdf_calls_com(self, fake_word, tmp_path):
        doc = _active()
        written = doc.export_pdf(tmp_path / "out.pdf")
        assert written == str((tmp_path / "out.pdf").resolve())
        doc.com.ExportAsFixedFormat.assert_called_once()


# ---------------------------------------------------------------------------
# CLI — gated save / save-as / export-pdf and image screening
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# MCP — save commands honour the policy
# ---------------------------------------------------------------------------


class TestPersistenceMcp:
    def test_save_as_deny_all(self, fake_word):
        with pytest.raises(PathNotAllowedError):
            dispatch_write(InlineWorker(), "save_as", {"path": "C:/whatever/x.docx"})

    def test_save_as_whitelisted(self, fake_word, tmp_path):
        pol = PathPolicy(save_dirs=[tmp_path])
        out = dispatch_write(
            InlineWorker(), "save_as", {"path": str(tmp_path / "x.docx")}, policy=pol
        )
        assert out["ok"] is True
        assert out["command"] == "save_as"

    def test_export_pdf_whitelisted(self, fake_word, tmp_path):
        pol = PathPolicy(save_dirs=[tmp_path])
        out = dispatch_write(
            InlineWorker(), "export_pdf", {"path": str(tmp_path / "x.pdf")}, policy=pol
        )
        assert out["ok"] is True
        assert out["command"] == "export_pdf"


# ---------------------------------------------------------------------------
# Bookmark / section CLI consolidation (verb-first)
# ---------------------------------------------------------------------------
